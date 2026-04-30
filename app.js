/* global marked, hljs, DOMPurify, jsyaml */
(function () {
  'use strict';

  const APP_VERSION = '1.5.0';
  document.getElementById('app-version').textContent = `v${APP_VERSION}`;

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
  const tabContents  = document.getElementById('tab-contents');
  const tabFiles     = document.getElementById('tab-files');
  const panelContents = document.getElementById('panel-contents');
  const panelFiles   = document.getElementById('panel-files');
  const tocSidebar   = document.getElementById('toc-sidebar');
  const tocNavEl     = document.getElementById('toc-nav-el');
  const tocList      = document.getElementById('toc-list');
  const fileTreeEl   = document.getElementById('file-tree-el');
  const folderNameEl = document.getElementById('folder-name-el');

  // ── State ────────────────────────────────────────────────────
  let localFolderLoaded = false;
  let activeFileItem    = null;
  let localFileMap      = new Map(); // normalised relPath → File
  let loadedRootName    = '';        // name of the folder the user selected
  let blobUrls          = [];        // blob: URLs to revoke on next render
  let fetchController   = null;      // current AbortController for fetch
  let renderGen         = 0;         // increments on every load; stale FileReader results ignored
  let tocGen            = 0;         // increments on every folder load; guards toc.yml FileReader
  let activeTocNavLink  = null;      // currently highlighted toc.yml nav link

  // Cache the initial welcome-screen HTML so back-navigation can restore it
  const welcomeHtml = contentArea.innerHTML;

  // ── Callout metadata ────────────────────────────────────────
  const CALLOUT_META = {
    NOTE:      { title: 'Note',      icon: 'ℹ️'  },
    TIP:       { title: 'Tip',       icon: '💡'  },
    WARNING:   { title: 'Warning',   icon: '⚠️'  },
    IMPORTANT: { title: 'Important', icon: '❗'  },
    CAUTION:   { title: 'Caution',   icon: '🔴' },
  };

  // ── Configure marked ────────────────────────────────────────
  marked.use({
    gfm: true,
    breaks: false,
    renderer: {
      // Render blockquote callouts (> [!NOTE] …) as styled callout divs
      blockquote({ tokens }) {
        const list = tokens ?? [];
        const firstPara = list.find(t => t.type === 'paragraph');
        if (firstPara) {
          const rawText = firstPara.text ?? '';
          const match   = rawText.match(/^\[!(NOTE|TIP|WARNING|IMPORTANT|CAUTION)\]\s*/i);
          if (match) {
            const type       = match[1].toUpperCase();
            const meta       = CALLOUT_META[type];
            const inlineText = rawText.slice(match[0].length).trim();
            const bodyTokens = list.filter(t => t !== firstPara);
            const bodyHtml   = this.parser.parse(bodyTokens);

            let html = `<div class="callout callout-${type.toLowerCase()}">`;
            html += `<p class="callout-title" aria-label="${meta.title}">${meta.icon} ${meta.title}</p>`;
            if (inlineText) html += `<p>${marked.parseInline(inlineText)}</p>`;
            html += bodyHtml;
            html += `</div>\n`;
            return html;
          }
        }
        const body = this.parser.parse(list);
        return `<blockquote>\n${body}</blockquote>\n`;
      },
      // Syntax-highlight fenced code blocks and add data-lang label
      code({ text, lang }) {
        const safeText = text ?? '';
        let highlighted;
        let detectedLang = null;
        try {
          if (lang && hljs.getLanguage(lang)) {
            highlighted  = hljs.highlight(safeText, { language: lang, ignoreIllegals: true }).value;
            detectedLang = lang;
          } else {
            const result = hljs.highlightAuto(safeText);
            highlighted  = result.value;
            detectedLang = result.language || null;
          }
        } catch (_) {
          highlighted = escapeHtml(safeText);
        }
        const langAttr = detectedLang ? ` data-lang="${detectedLang}"` : '';
        return `<pre${langAttr}><code class="hljs">${highlighted}</code></pre>\n`;
      },
      // Wrap images in mx-imgBorder span (matches learn.microsoft.com HTML)
      image({ href, title, text }) {
        if (!href) return `<span class="img-placeholder">[image: ${escapeHtml(text || 'no src')}]</span>`;
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
    return (md ?? '').replace(/^---\r?\n[\s\S]*?\r?\n---\r?\n/, '');
  }

  /**
   * Detect mis-formatted callouts written without a blockquote ">" prefix
   * and replace them with an authoring-error notice.
   * (Correctly formatted callouts are handled by the marked blockquote renderer.)
   */
  function postProcessCallouts(container) {
    const CALLOUT_RE = /^\[!(NOTE|TIP|WARNING|IMPORTANT|CAUTION)\]/i;
    container.querySelectorAll('p').forEach(p => {
      // Skip paragraphs already inside a rendered callout div
      if (p.closest('.callout')) return;
      const match = p.textContent.trim().match(CALLOUT_RE);
      if (!match) return;

      const type  = match[1].toUpperCase();
      const raw   = p.textContent.trim();
      const div   = document.createElement('div');
      div.className = 'authoring-error';
      div.innerHTML =
        `<p class="authoring-error-title">⚠️ Authoring error — missing <code>&gt;</code> prefix</p>` +
        `<p>This callout was written as plain text instead of a blockquote. Change it to:</p>` +
        `<pre><code>&gt; [!${escapeHtml(type)}]\n&gt; ${escapeHtml(raw.replace(CALLOUT_RE, '').trim())}</code></pre>`;
      p.replaceWith(div);
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
    [
      { id: 'contents', tab: tabContents, panel: panelContents },
      { id: 'files',    tab: tabFiles,    panel: panelFiles    },
    ].forEach(({ id, tab, panel }) => {
      const active = id === name;
      tab.setAttribute('aria-selected', active ? 'true' : 'false');
      panel.hidden = !active;
    });
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
      // Scroll without touching the URL — prevents popstate from firing
      // and resetting the page to the welcome screen.
      a.addEventListener('click', e => {
        e.preventDefault();
        document.getElementById(h.id)?.scrollIntoView({ behavior: 'smooth', block: 'start' });
      });
      li.appendChild(a);
      tocList.appendChild(li);
    });

    if (localFolderLoaded) {
      sidebar.hidden = false;
      // Show right TOC column whenever there are headings
      tocSidebar.hidden = tocList.children.length === 0;
    } else {
      // GitHub URL mode — only the right column is used; left sidebar stays hidden
      tocSidebar.hidden = tocList.children.length === 0;
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
    return String(str ?? '')
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

  // ── toc.yml navigation ───────────────────────────────────────

  /**
   * Case-insensitive lookup in localFileMap.
   * Also tries appending ".md" for extension-less hrefs common in toc.yml.
   * Returns { file, path } or null.
   */
  function findLocalFile(relPath) {
    if (!relPath) return null;
    const candidates = [relPath];
    // Extension-less hrefs: try appending .md
    if (!/\.\w+$/.test(relPath)) {
      candidates.push(relPath + '.md');
      // Directory-style hrefs (e.g. "overview" or "overview/"): try index.md inside
      candidates.push(relPath + '/index.md');
    }
    for (const candidate of candidates) {
      if (localFileMap.has(candidate)) return { file: localFileMap.get(candidate), path: candidate };
      const lower = candidate.toLowerCase();
      for (const [key, file] of localFileMap) {
        if (key.toLowerCase() === lower) return { file, path: key };
      }
    }
    return null;
  }

  /**
   * Build a nav link (or a plain span for unresolvable hrefs) for a toc.yml entry.
   */
  function buildTocNavLink(name, href, baseDir) {
    // External URL — open in new tab
    if (/^https?:\/\//i.test(href)) {
      const a = document.createElement('a');
      a.className = 'toc-nav-link';
      a.textContent = name;
      a.href = href;
      a.target = '_blank';
      a.rel = 'noopener noreferrer';
      return a;
    }

    const fragment  = href.match(/(#[^?]*)$/)?.[1] || '';
    const cleanHref = href.replace(/(#[^?]*)$/, '').trim();

    if (!cleanHref) {
      const span = document.createElement('span');
      span.className = 'toc-nav-label';
      span.textContent = name;
      return span;
    }

    // "~/" prefix means repo-root-relative in Learn/DocFX repos — skip baseDir join
    let resolved;
    if (cleanHref.startsWith('~/')) {
      resolved = cleanHref.slice(2);
    } else {
      resolved = resolveLocalPath(cleanHref, baseDir);
      // Strip loadedRootName prefix (case-insensitive). Happens when toc.yml uses
      // "../<rootFolder>/file.md" paths and the user loaded <rootFolder> directly:
      // resolveLocalPath produces "<rootFolder>/file.md" but localFileMap keys have
      // no root prefix.
      if (loadedRootName) {
        const prefix = loadedRootName.toLowerCase() + '/';
        if (resolved.toLowerCase().startsWith(prefix)) {
          resolved = resolved.slice(loadedRootName.length + 1);
        }
      }
    }
    const found = findLocalFile(resolved);

    if (!found) {
      // Detect whether the href escapes the loaded folder root via ".." traversal
      const upCount      = cleanHref.replace(/\\/g, '/').split('/').filter(p => p === '..').length;
      const baseDirDepth = baseDir ? baseDir.split('/').length : 0;
      const aboveRoot    = upCount > baseDirDepth;

      console.warn('[DocPreview] toc.yml: could not find file for href:', cleanHref, '(resolved:', resolved + ')');
      const a = document.createElement('a');
      a.className = 'toc-nav-link toc-nav-link--missing';
      a.textContent = name;
      a.href = '#';
      a.title = aboveRoot
        ? `Outside loaded folder — load the parent of "${loadedRootName}"`
        : `File not found in loaded folder: ${cleanHref}`;
      a.addEventListener('click', e => {
        e.preventDefault();
        if (aboveRoot) {
          contentArea.innerHTML =
            `<div class="error-state">` +
            `<h3>File is outside the loaded folder</h3>` +
            `<p><code>${escapeHtml(cleanHref)}</code> references a file in a parent or sibling folder.</p>` +
            `<p>To navigate this TOC, load the <strong>parent folder</strong> of ` +
            `<strong>${escapeHtml(loadedRootName)}</strong> instead of just ` +
            `<strong>${escapeHtml(loadedRootName)}</strong> itself.</p>` +
            `</div>`;
        } else {
          contentArea.innerHTML =
            `<div class="error-state">` +
            `<h3>File not loaded</h3>` +
            `<p><code>${escapeHtml(cleanHref)}</code> was referenced in toc.yml but is not in the loaded folder.</p>` +
            `<p>Try loading the folder that contains this file.</p>` +
            `</div>`;
        }
        sidebar.hidden = false;
        breadcrumb.hidden = true;
      });
      return a;
    }

    const a = document.createElement('a');
    a.className = 'toc-nav-link';
    a.textContent = name;
    a.href = '#';
    a.dataset.tocPath = found.path;
    if (fragment) a.dataset.tocFragment = fragment;

    a.addEventListener('click', e => {
      e.preventDefault();
      syncTocNavToPath(found.path);
      loadLocalFile(found.file, fragment || undefined);
    });
    return a;
  }

  /**
   * Recursively render toc.yml items into a <ul>.
   * depth 0 = root level (expanded by default); deeper levels follow YAML `expanded` field.
   */
  function renderTocYmlItems(items, ulEl, baseDir, depth) {
    if (!Array.isArray(items)) return;
    depth = depth || 0;

    items.forEach(item => {
      if (!item || typeof item !== 'object') return;

      const name    = String(item.name || item.displayName || item.title || '').trim();
      if (!name) return;

      // Prefer `href`; fall back to `topicHref` (used on group nodes in DocFX/Learn)
      const rawHref    = (item.href != null ? item.href : item.topicHref) || null;
      const children   = Array.isArray(item.items) && item.items.length > 0 ? item.items : null;
      const shouldExpand = depth === 0 || item.expanded === true;

      const li = document.createElement('li');
      li.className = 'toc-nav-item';

      if (children) {
        li.classList.add('has-children');
        if (shouldExpand) li.classList.add('expanded');

        const row = document.createElement('div');
        row.className = 'toc-nav-row';

        const toggleBtn = document.createElement('button');
        toggleBtn.className = 'toc-nav-toggle';
        toggleBtn.setAttribute('aria-expanded', String(shouldExpand));
        toggleBtn.title = shouldExpand ? 'Collapse section' : 'Expand section';
        toggleBtn.innerHTML = '<span class="toc-nav-chevron" aria-hidden="true">›</span>';

        const childUl = document.createElement('ul');
        childUl.className = 'toc-nav-children';
        childUl.hidden = !shouldExpand;

        toggleBtn.addEventListener('click', () => {
          const expanding = childUl.hidden;
          childUl.hidden = !expanding;
          toggleBtn.setAttribute('aria-expanded', String(expanding));
          toggleBtn.title = expanding ? 'Collapse section' : 'Expand section';
          li.classList.toggle('expanded', expanding);
        });

        row.appendChild(toggleBtn);

        if (rawHref) {
          row.appendChild(buildTocNavLink(name, String(rawHref), baseDir));
        } else {
          const span = document.createElement('span');
          span.className = 'toc-nav-label';
          span.textContent = name;
          row.appendChild(span);
        }

        li.appendChild(row);
        renderTocYmlItems(item.items, childUl, baseDir, depth + 1);
        li.appendChild(childUl);

      } else if (rawHref) {
        li.appendChild(buildTocNavLink(name, String(rawHref), baseDir));
      } else {
        const span = document.createElement('span');
        span.className = 'toc-nav-label';
        span.textContent = name;
        li.appendChild(span);
      }

      ulEl.appendChild(li);
    });
  }

  /**
   * Highlight all toc.yml nav links matching relPath and expand their ancestors.
   * Clears any previous active state first.
   */
  function syncTocNavToPath(relPath) {
    tocNavEl.querySelectorAll('a.toc-nav-link.active').forEach(a => a.classList.remove('active'));
    activeTocNavLink = null;
    if (!relPath) return;

    const lower = relPath.toLowerCase().replace(/\\/g, '/');
    let firstMatch = null;

    tocNavEl.querySelectorAll('a[data-toc-path]').forEach(a => {
      if (a.dataset.tocPath.toLowerCase() !== lower) return;
      a.classList.add('active');
      if (!firstMatch) firstMatch = a;

      // Expand all ancestor toc-nav-children
      let node = a.parentElement;
      while (node && node !== tocNavEl) {
        if (node.classList.contains('toc-nav-children')) {
          node.hidden = false;
          const parentLi = node.parentElement;
          if (parentLi && parentLi.classList.contains('toc-nav-item')) {
            parentLi.classList.add('expanded');
            const toggle = parentLi.querySelector(':scope > .toc-nav-row > .toc-nav-toggle');
            if (toggle) {
              toggle.setAttribute('aria-expanded', 'true');
              toggle.title = 'Collapse section';
            }
          }
        }
        node = node.parentElement;
      }
    });

    activeTocNavLink = firstMatch;
    if (firstMatch) firstMatch.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
  }

  /** Show a toc.yml error in the Contents panel and make it visible. */
  function showTocYmlError(msg) {
    tocNavEl.innerHTML = `<p class="toc-error">${escapeHtml(msg)}</p>`;
    tabContents.hidden = false;
    showTab('contents');
  }

  /**
   * Find the most root-level toc.yml in localFileMap, parse it with js-yaml,
   * and populate the Contents panel. Shows the Contents tab if successful.
   */
  function findAndLoadTocYml() {
    if (typeof jsyaml === 'undefined') return;

    let tocFile = null;
    let tocPath = null;

    for (const [path, file] of localFileMap) {
      const name = path.split('/').pop().toLowerCase();
      if (name === 'toc.yml' || name === 'toc.yaml') {
        const depth = path.split('/').length;
        if (tocPath === null || depth < tocPath.split('/').length) {
          tocFile = file;
          tocPath = path;
        }
      }
    }

    if (!tocFile) {
      tabContents.hidden = true;
      return;
    }

    const gen     = ++tocGen;
    const baseDir = tocPath.split('/').slice(0, -1).join('/');
    const reader  = new FileReader();

    reader.onload = e => {
      if (gen !== tocGen) return; // superseded by a newer folder load
      try {
        let parsed = jsyaml.load(e.target.result);
        // Handle top-level { items: [...] } wrapper used by some DocFX repos
        if (parsed && !Array.isArray(parsed) && Array.isArray(parsed.items)) {
          parsed = parsed.items;
        }
        if (!Array.isArray(parsed)) {
          showTocYmlError('Unsupported toc.yml format (expected a YAML list)');
          return;
        }

        const ul = document.createElement('ul');
        ul.className = 'toc-nav-list';
        renderTocYmlItems(parsed, ul, baseDir, 0);

        // Warn if any entries couldn't be resolved (dimmed links)
        // Count them after rendering so the user knows upfront
        tocNavEl.innerHTML = '';

        const missingLinks = ul.querySelectorAll('.toc-nav-link--missing');
        if (missingLinks.length > 0) {
          // Check if any of them escape the loaded root (above-root hrefs)
          const aboveRootCount = Array.from(missingLinks)
            .filter(a => a.title.startsWith('Outside loaded folder')).length;
          if (aboveRootCount > 0) {
            const banner = document.createElement('p');
            banner.className = 'toc-warning-banner';
            banner.innerHTML =
              `⚠️ ${aboveRootCount} TOC ${aboveRootCount === 1 ? 'entry references a file' : 'entries reference files'} ` +
              `outside <strong>${escapeHtml(loadedRootName)}</strong>. ` +
              `Load the <strong>parent folder</strong> of <strong>${escapeHtml(loadedRootName)}</strong> to access them.`;
            tocNavEl.appendChild(banner);
          }
        }

        tocNavEl.appendChild(ul);
        tabContents.hidden = false;
        showTab('contents');
      } catch (err) {
        console.warn('Failed to parse toc.yml:', err);
        showTocYmlError('Could not parse toc.yml: ' + err.message);
      }
    };
    reader.onerror = () => {
      if (gen !== tocGen) return;
      showTocYmlError('Could not read toc.yml');
    };
    reader.readAsText(tocFile);
  }

  // ── Shared render function ───────────────────────────────────
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
    tocSidebar.hidden = true;
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
      try {
        renderContent(e.target.result, file.name, null, relPath);
        scrollToFragment(fragment);
        syncTocNavToPath(relPath);
      } catch (err) {
        console.error('renderContent failed:', err);
        contentArea.innerHTML =
          `<div class="error-state">` +
          `<h3>Render error</h3>` +
          `<p>${escapeHtml(String(err))}</p>` +
          `</div>`;
      }
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

  browseBtn.addEventListener('click', () => {
    folderInput.value = ''; // reset BEFORE opening so re-selecting same folder works
    folderInput.click();    // and File objects from previous selection stay valid
  });

  folderInput.addEventListener('change', () => {
    const { files } = folderInput;
    if (!files || files.length === 0) return;

    // Single file selected (some OS/browsers allow picking one file even with
    // webkitdirectory). webkitRelativePath has no "/" so there's no folder root.
    const firstRelPath = files[0].webkitRelativePath || '';
    if (files.length === 1 && !firstRelPath.includes('/')) {
      loadLocalFile(files[0]);
      return;
    }

    // Reset toc.yml state from any previous folder
    ++tocGen; // invalidate any in-flight toc FileReader
    tocNavEl.innerHTML = '';
    tabContents.hidden = true;
    activeTocNavLink = null;

    const rootName = firstRelPath.split('/')[0];
    loadedRootName = rootName;
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

    // Async: shows the Contents tab if a toc.yml is found
    findAndLoadTocYml();
    // ← do NOT clear folderInput.value here; that invalidates File objects in Safari
  });

  tabContents.addEventListener('click', () => showTab('contents'));
  tabFiles.addEventListener(   'click', () => showTab('files'));

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
      tocSidebar.hidden = true;
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
