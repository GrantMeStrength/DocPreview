/* global marked, hljs, DOMPurify */
(function () {
  'use strict';

  // ── DOM refs ────────────────────────────────────────────────
  const urlInput     = document.getElementById('urlInput');
  const loadBtn      = document.getElementById('loadBtn');
  const browseBtn    = document.getElementById('browseBtn');
  const folderInput  = document.getElementById('folderInput');
  const contentArea  = document.getElementById('content-area');
  const breadcrumb   = document.getElementById('breadcrumb');
  const bcRepo       = document.getElementById('bc-repo');
  const bcFile       = document.getElementById('bc-file');
  const sidebar      = document.getElementById('sidebar');
  const sidebarTabs  = document.getElementById('sidebar-tabs');
  const tabFiles     = document.getElementById('tab-files');
  const tabToc       = document.getElementById('tab-toc');
  const panelFiles   = document.getElementById('panel-files');
  const panelToc     = document.getElementById('panel-toc');
  const tocList      = document.getElementById('toc-list');
  const fileTreeEl   = document.getElementById('file-tree-el');
  const folderNameEl = document.getElementById('folder-name-el');

  // ── State ────────────────────────────────────────────────────
  let localFolderLoaded = false;
  let activeFileItem    = null;
  let localFileMap      = new Map(); // normalised relPath → File
  let blobUrls          = [];        // blob: URLs to revoke on next render
  let fetchController   = null;      // current AbortController for fetch
  let renderGen         = 0;         // increments on every load; stale FileReader results ignored

  // Cache the initial welcome-screen HTML so back-navigation can restore it
  const welcomeHtml = contentArea.innerHTML;

  // ── Configure marked ────────────────────────────────────────
  marked.use({
    gfm: true,
    breaks: false,
    renderer: {
      // Syntax-highlight fenced code blocks and add data-lang label
      code({ text, lang }) {
        let highlighted;
        let detectedLang = null;
        try {
          if (lang && hljs.getLanguage(lang)) {
            highlighted  = hljs.highlight(text, { language: lang, ignoreIllegals: true }).value;
            detectedLang = lang;
          } else {
            const result = hljs.highlightAuto(text);
            highlighted  = result.value;
            detectedLang = result.language || null;
          }
        } catch (_) {
          highlighted = escapeHtml(text);
        }
        const langAttr = detectedLang ? ` data-lang="${detectedLang}"` : '';
        return `<pre${langAttr}><code class="hljs">${highlighted}</code></pre>\n`;
      },
      // Wrap images in mx-imgBorder span (matches learn.microsoft.com HTML)
      image({ href, title, text }) {
        const alt       = text  ? ` alt="${escapeHtml(text)}"` : '';
        const titleAttr = title ? ` title="${escapeHtml(title)}"` : '';
        return `<span class="mx-imgBorder"><img src="${href}"${alt}${titleAttr} loading="lazy"></span>`;
      }
    }
  });

  // ── OFM preprocessor ────────────────────────────────────────
  /**
   * Convert OFM (Open File Markdown) extensions used by learn.microsoft.com
   * into standard markdown or HTML that marked.js can handle.
   */
  function preprocessOFM(md) {
    // :::image type="content" source="path" alt="text"::: → standard image
    md = md.replace(/^:::\s*image\b([^:]*?):::\s*$/gm, (_, attrs) => {
      const src = /source="([^"]*)"/.exec(attrs)?.[1] ?? '';
      const alt = /alt="([^"]*)"/.exec(attrs)?.[1] ?? '';
      return src ? `\n![${alt}](${src})\n` : '';
    });

    // [!INCLUDE [text](path)] → strip (file includes can't be resolved)
    md = md.replace(/\[!INCLUDE\s*\[[^\]]*\]\([^)]+\)\]/gi, '');

    return md;
  }

  // ── URL helpers ─────────────────────────────────────────────

  /** Convert any GitHub markdown URL to a raw.githubusercontent.com URL */
  function toRawUrl(input) {
    const url = input.trim()
      .replace(/#[^?]*$/, '')   // strip fragment
      .replace(/\?[^#]*$/, ''); // strip query string

    if (url.includes('raw.githubusercontent.com')) return url;

    // github.com/org/repo/blob/...  OR  github.com/org/repo/raw/...
    const m = url.match(/^https?:\/\/github\.com\/([^/]+\/[^/]+)\/(?:blob|raw)\/(.+)$/);
    if (m) return `https://raw.githubusercontent.com/${m[1]}/${m[2]}`;

    return url;
  }

  /**
   * Resolve a relative href/src against a base directory URL.
   * Returns the input unchanged if it is already absolute/protocol-relative.
   */
  function resolveUrl(rel, baseDir, rawBase) {
    if (!rel) return rel;
    // Absolute, protocol-relative, mailto, data, fragments – leave alone
    if (/^([a-z][a-z0-9+\-.]*:|\/\/|#|mailto:|data:)/i.test(rel)) return rel;

    if (rel.startsWith('/')) {
      return rawBase + rel.replace(/^\//, '');
    }
    try { return new URL(rel, baseDir).href; } catch (_) { return rel; }
  }

  /**
   * Resolve a relative path against a relative directory (for local files).
   * Handles `.` and `..` segments. Returns normalised forward-slash path.
   */
  function resolveLocalPath(rel, relDir) {
    const base  = relDir ? `${relDir}/${rel}` : rel;
    const parts = base.replace(/\\/g, '/').split('/');
    const out   = [];
    for (const p of parts) {
      if (p === '..') out.pop();
      else if (p && p !== '.') out.push(p);
    }
    return out.join('/');
  }

  // ── YAML front-matter stripper ──────────────────────────────
  function stripFrontMatter(md) {
    // Only strip if the document starts at byte 0 with "---"
    return md.replace(/^---\r?\n[\s\S]*?\r?\n---\r?\n/, '');
  }

  // ── Callout post-processor ──────────────────────────────────
  const CALLOUT_META = {
    NOTE:      { title: 'Note',      icon: 'ℹ️'  },
    TIP:       { title: 'Tip',       icon: '💡'  },
    WARNING:   { title: 'Warning',   icon: '⚠️'  },
    IMPORTANT: { title: 'Important', icon: '❗'  },
    CAUTION:   { title: 'Caution',   icon: '🔴' },
  };

  /**
   * Turn  <blockquote><p>[!NOTE] …</p>…</blockquote>
   * into  <div class="callout callout-note">…</div>
   */
  function postProcessCallouts(container) {
    container.querySelectorAll('blockquote').forEach(bq => {
      const firstP = bq.querySelector('p:first-child');
      if (!firstP) return;

      const match = firstP.textContent.trim().match(/^\[!(NOTE|TIP|WARNING|IMPORTANT|CAUTION)\]/i);
      if (!match) return;

      const type = match[1].toUpperCase();
      const meta = CALLOUT_META[type];

      const div = document.createElement('div');
      div.className = `callout callout-${type.toLowerCase()}`;

      const titleEl = document.createElement('p');
      titleEl.className = 'callout-title';
      titleEl.setAttribute('aria-label', meta.title);
      titleEl.innerHTML = `${meta.icon} ${meta.title}`;
      div.appendChild(titleEl);

      // Strip the [!TYPE] token from firstP
      const strippedHTML = firstP.innerHTML
        .replace(/^\[!(NOTE|TIP|WARNING|IMPORTANT|CAUTION)\](<br\s*\/?>|\s)*/i, '')
        .trim();

      Array.from(bq.childNodes).forEach((child, idx) => {
        if (idx === 0) {
          if (strippedHTML) {
            const newP = document.createElement('p');
            newP.innerHTML = strippedHTML;
            div.appendChild(newP);
          }
        } else {
          div.appendChild(child);
        }
      });

      bq.replaceWith(div);
    });
  }

  // ── Remote relative URL rewriter ────────────────────────────
  function rewriteRelativeUrls(container, rawDocUrl) {
    let baseDir, rawBase, repoGHBase;
    try {
      const u     = new URL(rawDocUrl);
      const parts = u.pathname.split('/').filter(Boolean);
      const org    = parts[0];
      const repo   = parts[1];
      const branch = parts[2];
      const pathParts = parts.slice(0, -1);

      baseDir    = `${u.origin}/${pathParts.join('/')}/`;
      rawBase    = `https://raw.githubusercontent.com/${org}/${repo}/${branch}/`;
      repoGHBase = `https://github.com/${org}/${repo}/blob/${branch}/`;
    } catch (_) {
      return;
    }

    container.querySelectorAll('img[src]').forEach(img => {
      const src = img.getAttribute('src');
      const resolved = resolveUrl(src, baseDir, rawBase);
      if (resolved !== src) img.src = resolved;
    });

    container.querySelectorAll('a[href]').forEach(a => {
      const href = a.getAttribute('href');
      if (!href) return;
      if (href.startsWith('#')) return;
      if (/^[a-z][a-z0-9+\-.]*:/i.test(href)) {
        if (!href.startsWith(window.location.origin)) {
          a.target = '_blank';
          a.rel    = 'noopener noreferrer';
        }
        return;
      }

      if (/\.md(#[^?]*)?$/i.test(href)) {
        const fragment  = href.match(/(#[^?]*)$/)?.[1] || '';
        const cleanHref = href.replace(/(#[^?]*)$/, '');
        const absRaw    = resolveUrl(cleanHref, baseDir, rawBase);
        a.href          = `${window.location.pathname}?url=${encodeURIComponent(absRaw)}${fragment}`;
        a.target        = '_self';
        a.addEventListener('click', e => {
          e.preventDefault();
          const newUrl = new URL(a.href, window.location.href);
          const target = newUrl.searchParams.get('url');
          if (target) {
            urlInput.value = target;
            loadMarkdown(target, true, fragment);
          }
        });
        return;
      }

      const abs = resolveUrl(href, baseDir, repoGHBase);
      a.href   = abs;
      a.target = '_blank';
      a.rel    = 'noopener noreferrer';
    });
  }

  // ── Local relative URL rewriter ─────────────────────────────
  /**
   * Rewrites relative img src to blob: URLs using files from localFileMap,
   * and rewrites relative .md links to open via loadLocalFile.
   */
  function postProcessLocalUrls(container, relFilePath) {
    // Revoke blob URLs from previous render to free memory
    blobUrls.forEach(u => URL.revokeObjectURL(u));
    blobUrls = [];

    const parts  = relFilePath.replace(/\\/g, '/').split('/');
    const relDir = parts.slice(0, -1).join('/'); // directory of current file

    container.querySelectorAll('img[src]').forEach(img => {
      const src = img.getAttribute('src');
      if (!src || /^(https?:|blob:|data:|\/\/)/i.test(src)) return;

      const resolved = resolveLocalPath(src, relDir);
      const file     = localFileMap.get(resolved);
      if (file) {
        const blobUrl = URL.createObjectURL(file);
        blobUrls.push(blobUrl);
        img.src = blobUrl;
      }
    });

    container.querySelectorAll('a[href]').forEach(a => {
      const href = a.getAttribute('href');
      if (!href || /^(https?:|#|mailto:)/i.test(href)) return;
      if (!/\.md(#[^?]*)?$/i.test(href)) return;

      const fragment  = href.match(/(#[^?]*)$/)?.[1] || '';
      const cleanHref = href.replace(/(#[^?]*)$/, '');
      const resolved  = resolveLocalPath(cleanHref, relDir);
      const targetFile = localFileMap.get(resolved);

      if (targetFile) {
        a.href = fragment || '#';
        a.addEventListener('click', e => {
          e.preventDefault();
          // Try to highlight the target in the file tree
          activateLocalFileInTree(resolved);
          loadLocalFile(targetFile, fragment);
        });
      }
    });
  }

  /**
   * Mark a file in the rendered tree as active based on its repo-relative path.
   * Opens parent directories as needed.
   */
  function activateLocalFileInTree(relPath) {
    // Walk the rendered tree looking for the matching file button
    const allFileBtns = fileTreeEl.querySelectorAll('.ft-file-btn');
    allFileBtns.forEach(btn => {
      const li = btn.parentElement;
      if (li._relPath === relPath) {
        if (activeFileItem) activeFileItem.classList.remove('active');
        activeFileItem = li;
        li.classList.add('active');
      }
    });
  }

  // ── Heading ID slugger (collision-aware) ────────────────────
  function buildHeadingIds(container) {
    const seen = Object.create(null);
    container.querySelectorAll('h1, h2, h3, h4, h5, h6').forEach(h => {
      const base = h.textContent
        .trim()
        .toLowerCase()
        .replace(/[^\w\s-]/g, '')
        .replace(/\s+/g, '-')
        .replace(/-+/g, '-')
        .replace(/^-|-$/g, '') || 'section';
      seen[base] = (seen[base] || 0) + 1;
      h.id = seen[base] === 1 ? base : `${base}-${seen[base]}`;
    });
  }

  // ── Sidebar tab switcher ─────────────────────────────────────
  function showTab(name) {
    const isFiles = name === 'files';
    tabFiles.setAttribute('aria-selected', isFiles ? 'true' : 'false');
    tabToc.setAttribute('aria-selected', isFiles ? 'false' : 'true');
    panelFiles.hidden = !isFiles;
    panelToc.hidden   =  isFiles;
  }

  // ── TOC builder ─────────────────────────────────────────────
  function buildTOC(container) {
    tocList.innerHTML = '';
    container.querySelectorAll('h2, h3, h4').forEach(h => {
      const li = document.createElement('li');
      li.className = `toc-${h.tagName.toLowerCase()}`;
      const a = document.createElement('a');
      a.href = `#${h.id}`;
      a.textContent = h.textContent;
      li.appendChild(a);
      tocList.appendChild(li);
    });

    if (localFolderLoaded) {
      sidebar.hidden = false;
      // Switch to TOC so the user can navigate headings after opening a file
      if (tocList.children.length > 0) showTab('toc');
    } else {
      sidebar.hidden = tocList.children.length === 0;
    }
  }

  // ── Copy buttons ─────────────────────────────────────────────
  function addCopyButtons(container) {
    container.querySelectorAll('pre').forEach(pre => {
      const btn = document.createElement('button');
      btn.className = 'code-copy-btn';
      btn.textContent = 'Copy';
      btn.setAttribute('aria-label', 'Copy code');
      btn.addEventListener('click', () => {
        const text = pre.querySelector('code')?.textContent ?? pre.textContent;
        navigator.clipboard.writeText(text).then(() => {
          btn.textContent = 'Copied!';
          btn.classList.add('copied');
          btn.setAttribute('aria-label', 'Copied');
          setTimeout(() => {
            btn.textContent = 'Copy';
            btn.classList.remove('copied');
            btn.setAttribute('aria-label', 'Copy code');
          }, 2000);
        }).catch(() => {
          btn.textContent = 'Failed';
          setTimeout(() => { btn.textContent = 'Copy'; }, 2000);
        });
      });
      pre.appendChild(btn);
    });
  }

  // ── Scroll spy ───────────────────────────────────────────────
  let activeObserver = null;

  function setupScrollSpy() {
    if (activeObserver) { activeObserver.disconnect(); activeObserver = null; }

    const tocLinks = Array.from(tocList.querySelectorAll('a'));
    if (!tocLinks.length) return;

    activeObserver = new IntersectionObserver(entries => {
      entries.forEach(entry => {
        if (entry.isIntersecting) {
          const id = entry.target.id;
          tocLinks.forEach(a => a.classList.toggle('active', a.getAttribute('href') === `#${id}`));
        }
      });
    }, { rootMargin: '-15% 0px -75% 0px' });

    contentArea.querySelectorAll('h2, h3, h4').forEach(h => activeObserver.observe(h));
  }

  // ── Breadcrumb ───────────────────────────────────────────────
  function updateBreadcrumb(rawUrl) {
    try {
      const u     = new URL(rawUrl);
      const parts = u.pathname.replace(/^\//, '').split('/');
      const org    = parts[0];
      const repo   = parts[1];
      const file   = parts[parts.length - 1];

      bcRepo.textContent = `${org}/${repo}`;
      bcRepo.href        = `https://github.com/${org}/${repo}`;
      bcFile.textContent = file;
      breadcrumb.hidden  = false;
    } catch (_) {
      breadcrumb.hidden = true;
    }
  }

  // ── HTML escape helper ───────────────────────────────────────
  function escapeHtml(str) {
    return str
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  // ── Scroll to fragment ───────────────────────────────────────
  function scrollToFragment(fragment) {
    if (!fragment) return;
    const id = fragment.replace(/^#/, '');
    // Small delay lets the browser finish painting
    setTimeout(() => {
      const el = document.getElementById(id);
      if (el) el.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }, 60);
  }

  // ── Broken image placeholders ────────────────────────────────
  function addImagePlaceholders(container) {
    container.querySelectorAll('img').forEach(img => {
      img.addEventListener('error', () => {
        const label = img.alt || img.src.split('/').pop() || 'image';
        const span  = document.createElement('span');
        span.className   = 'img-placeholder';
        span.textContent = label;
        span.title       = img.src;
        img.replaceWith(span);
      }, { once: true });
    });
  }

  // ── Shared render function ───────────────────────────────────
  /**
   * Parse, sanitize, and render markdown into contentArea.
   * @param {string}      md           Raw markdown source
   * @param {string}      displayName  Used in document title if no h1 found
   * @param {string|null} rawUrl       Source URL (remote) — null for local files
   * @param {string|null} localRelPath Repo-relative path of local file (null for remote)
   */
  function renderContent(md, displayName, rawUrl, localRelPath) {
    const stripped  = preprocessOFM(stripFrontMatter(md));
    const dirtyHtml = marked.parse(stripped);
    const cleanHtml = DOMPurify.sanitize(dirtyHtml, {
      ADD_ATTR: ['target', 'rel', 'data-lang', 'loading'],
    });

    contentArea.innerHTML = cleanHtml;

    postProcessCallouts(contentArea);
    buildHeadingIds(contentArea);

    if (rawUrl)       rewriteRelativeUrls(contentArea, rawUrl);
    if (localRelPath) postProcessLocalUrls(contentArea, localRelPath);

    addImagePlaceholders(contentArea);

    buildTOC(contentArea);
    addCopyButtons(contentArea);
    setupScrollSpy();
    if (rawUrl) updateBreadcrumb(rawUrl);

    document.title = contentArea.querySelector('h1')?.textContent
      ? `${contentArea.querySelector('h1').textContent} – Doc Preview`
      : `${displayName} – Doc Preview`;
  }

  // ── GitHub URL load function ─────────────────────────────────
  async function loadMarkdown(inputUrl, pushHistory, fragment) {
    const trimmed = (inputUrl || '').trim();
    if (!trimmed) return;

    // Cancel any in-flight fetch
    if (fetchController) {
      fetchController.abort();
      fetchController = null;
    }
    const controller = new AbortController();
    fetchController  = controller;

    const rawUrl = toRawUrl(trimmed);

    contentArea.innerHTML = '<div class="loading-state"><div class="loading-spinner" aria-hidden="true"></div><span>Loading…</span></div>';
    if (!localFolderLoaded) sidebar.hidden = true;
    breadcrumb.hidden = true;

    if (activeFileItem) {
      activeFileItem.classList.remove('active');
      activeFileItem = null;
    }

    try {
      const res = await fetch(rawUrl, { signal: controller.signal });
      if (!res.ok) throw new Error(`${res.status} ${res.statusText}`);
      const raw = await res.text();

      if (controller.signal.aborted) return; // superseded by a newer load

      renderContent(raw, rawUrl.split('/').pop() || 'document', rawUrl, null);
      scrollToFragment(fragment);

      const pageUrl = new URL(window.location.href);
      pageUrl.searchParams.set('url', trimmed);
      const historyEntry = `${pageUrl.pathname}${pageUrl.search}${fragment || ''}`;
      if (pushHistory) {
        history.pushState({ url: trimmed, fragment: fragment || '' }, '', historyEntry);
      } else {
        history.replaceState({ url: trimmed, fragment: fragment || '' }, '', historyEntry);
      }

    } catch (err) {
      if (err.name === 'AbortError') return; // cancelled, ignore

      contentArea.innerHTML =
        `<div class="error-state">` +
        `<h3>Failed to load document</h3>` +
        `<p>${escapeHtml(err.message)}</p>` +
        `<p>Make sure the URL points to a <strong>public</strong> GitHub markdown file.</p>` +
        `</div>`;
      if (!localFolderLoaded) sidebar.hidden = true;
      breadcrumb.hidden = true;
    } finally {
      if (fetchController === controller) fetchController = null;
    }
  }

  // ── Local file load function ─────────────────────────────────
  function loadLocalFile(file, fragment) {
    // Cancel any in-flight remote fetch
    if (fetchController) { fetchController.abort(); fetchController = null; }

    const gen = ++renderGen; // stale-load guard

    contentArea.innerHTML = '<div class="loading-state"><div class="loading-spinner" aria-hidden="true"></div><span>Loading…</span></div>';
    breadcrumb.hidden = true;

    // repo-relative path (strip root folder prefix)
    const relPath = file.webkitRelativePath.split('/').slice(1).join('/');

    const reader = new FileReader();
    reader.onload = e => {
      if (gen !== renderGen) return; // superseded by a newer load
      renderContent(e.target.result, file.name, null, relPath);
      scrollToFragment(fragment);
    };
    reader.onerror = () => {
      if (gen !== renderGen) return;
      contentArea.innerHTML =
        `<div class="error-state">` +
        `<h3>Failed to read file</h3>` +
        `<p>Could not read ${escapeHtml(file.name)}.</p>` +
        `</div>`;
    };
    reader.readAsText(file);
  }

  // ── File tree model builder ──────────────────────────────────
  function buildTreeModel(files) {
    const root = { dirs: {}, files: [] };

    Array.from(files)
      .filter(f =>
        f.name.endsWith('.md') &&
        !f.webkitRelativePath.split('/').some(p => p.startsWith('.'))
      )
      .sort((a, b) => a.webkitRelativePath.localeCompare(b.webkitRelativePath))
      .forEach(file => {
        const inner = file.webkitRelativePath.split('/').slice(1); // drop root folder
        let node = root;
        for (let i = 0; i < inner.length - 1; i++) {
          const d = inner[i];
          if (!node.dirs[d]) node.dirs[d] = { dirs: {}, files: [] };
          node = node.dirs[d];
        }
        const fileName = inner[inner.length - 1];
        if (fileName) node.files.push(file);
      });

    return root;
  }

  // ── File tree DOM renderer ───────────────────────────────────
  function renderFileTree(node, ulEl) {
    const dirNames = Object.keys(node.dirs).sort((a, b) =>
      a.localeCompare(b, undefined, { sensitivity: 'base' })
    );
    const sortedFiles = node.files.slice().sort((a, b) =>
      a.name.localeCompare(b.name, undefined, { sensitivity: 'base' })
    );

    dirNames.forEach(dirName => {
      const dirNode = node.dirs[dirName];
      const li = document.createElement('li');
      li.className = 'ft-dir';

      const btn = document.createElement('button');
      btn.className = 'ft-dir-toggle';
      btn.setAttribute('aria-expanded', 'false');
      btn.innerHTML =
        `<span class="ft-chevron" aria-hidden="true">›</span>` +
        `<span class="ft-dir-name">${escapeHtml(dirName)}</span>`;

      const childUl = document.createElement('ul');
      childUl.hidden = true;

      let rendered = false;
      btn.addEventListener('click', () => {
        const expanding = childUl.hidden;
        childUl.hidden = !expanding;
        btn.setAttribute('aria-expanded', String(expanding));
        li.classList.toggle('expanded', expanding);
        if (expanding && !rendered) {
          renderFileTree(dirNode, childUl);
          rendered = true;
        }
      });

      li.appendChild(btn);
      li.appendChild(childUl);
      ulEl.appendChild(li);
    });

    sortedFiles.forEach(file => {
      const relPath = file.webkitRelativePath.split('/').slice(1).join('/');
      const li = document.createElement('li');
      li.className  = 'ft-file';
      li._relPath   = relPath; // used by activateLocalFileInTree

      const btn = document.createElement('button');
      btn.className = 'ft-file-btn';
      btn.setAttribute('aria-label', `Preview ${file.name}`);
      btn.innerHTML =
        `<span class="ft-file-icon" aria-hidden="true">📄</span>` +
        `<span class="ft-file-name">${escapeHtml(file.name)}</span>`;

      btn.addEventListener('click', () => {
        if (activeFileItem) activeFileItem.classList.remove('active');
        activeFileItem = li;
        li.classList.add('active');
        loadLocalFile(file);
      });

      li.appendChild(btn);
      ulEl.appendChild(li);
    });
  }

  // ── Event listeners ──────────────────────────────────────────
  loadBtn.addEventListener('click', () => loadMarkdown(urlInput.value, true));

  urlInput.addEventListener('keydown', e => {
    if (e.key === 'Enter') loadMarkdown(urlInput.value, true);
  });

  browseBtn.addEventListener('click', () => folderInput.click());

  folderInput.addEventListener('change', () => {
    const { files } = folderInput;
    if (!files || files.length === 0) return;

    const rootName = files[0].webkitRelativePath.split('/')[0];
    folderNameEl.textContent = rootName;

    // Build lookup map: repo-relative path → File (for all file types, for image resolution)
    localFileMap = new Map();
    Array.from(files).forEach(f => {
      const relPath = f.webkitRelativePath.split('/').slice(1).join('/');
      localFileMap.set(relPath, f);
    });

    fileTreeEl.innerHTML = '';
    renderFileTree(buildTreeModel(files), fileTreeEl);

    localFolderLoaded = true;
    sidebarTabs.hidden = false;
    sidebar.hidden = false;
    showTab('files');

    folderInput.value = ''; // allow re-opening the same folder
  });

  tabFiles.addEventListener('click', () => showTab('files'));
  tabToc.addEventListener('click',   () => showTab('toc'));

  // Back / forward navigation
  window.addEventListener('popstate', e => {
    const state   = e.state;
    const params  = new URLSearchParams(window.location.search);
    const url     = state?.url || params.get('url');
    const fragment = state?.fragment || window.location.hash || '';

    if (url) {
      urlInput.value = url;
      loadMarkdown(url, false, fragment);
    } else {
      // Back to start — restore welcome screen
      urlInput.value        = '';
      contentArea.innerHTML = welcomeHtml;
      if (!localFolderLoaded) sidebar.hidden = true;
      breadcrumb.hidden = true;
      document.title    = 'Doc Preview – Microsoft Learn Style';
    }
  });

  // ── Auto-load from ?url= param on startup ────────────────────
  const startUrl = new URLSearchParams(window.location.search).get('url');
  const startFragment = window.location.hash || '';
  if (startUrl) {
    urlInput.value = startUrl;
    loadMarkdown(startUrl, false, startFragment);
  }

})();
