/* global marked, hljs, DOMPurify */
(function () {
  'use strict';

  // ── DOM refs ────────────────────────────────────────────────
  const urlInput   = document.getElementById('urlInput');
  const loadBtn    = document.getElementById('loadBtn');
  const contentArea= document.getElementById('content-area');
  const breadcrumb = document.getElementById('breadcrumb');
  const bcRepo     = document.getElementById('bc-repo');
  const bcFile     = document.getElementById('bc-file');
  const sidebar    = document.getElementById('sidebar');
  const tocList    = document.getElementById('toc-list');

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
      }
    }
  });

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
   * Resolve a relative href/src found in the document
   * against the base directory URL of the document being previewed.
   */
  function resolveUrl(rel, baseDir, rawBase, repoGHBase) {
    if (!rel) return rel;
    // Absolute, protocol-relative, mailto, data, fragments – leave alone
    if (/^([a-z][a-z0-9+\-.]*:|\/\/|#|mailto:|data:)/i.test(rel)) return rel;

    let abs;
    if (rel.startsWith('/')) {
      // Root-relative – resolve against repo root
      abs = rawBase + rel.replace(/^\//, '');
    } else {
      try { abs = new URL(rel, baseDir).href; } catch (_) { return rel; }
    }
    return abs;
  }

  // ── YAML front-matter stripper ──────────────────────────────
  function stripFrontMatter(md) {
    return md.replace(/^\s*---[\s\S]*?---\s*\n/, '');
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
   *
   * Works for both:
   *   > [!NOTE]
   *   > text on next line   (renders as single <p> with newline)
   *
   *   > [!NOTE]
   *   >
   *   > text after blank    (renders as separate <p> elements)
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

      // Strip the [!TYPE] token (and any following <br> / whitespace) from firstP
      const strippedHTML = firstP.innerHTML
        .replace(/^\[!(NOTE|TIP|WARNING|IMPORTANT|CAUTION)\](<br\s*\/?>|\s)*/i, '')
        .trim();

      // Snapshot children before moving them (live NodeList would change)
      const children = Array.from(bq.childNodes);
      children.forEach((child, idx) => {
        if (idx === 0) {
          // This is firstP – replace with stripped version if non-empty
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

  // ── Relative URL rewriter ───────────────────────────────────
  /**
   * Rewrites relative img[src] to absolute raw GitHub URLs,
   * and relative a[href] to either absolute GitHub URLs or
   * ?url=... preview links for .md files.
   */
  function rewriteRelativeUrls(container, rawDocUrl) {
    let baseDir, rawBase, repoGHBase;
    try {
      const u     = new URL(rawDocUrl);
      const parts = u.pathname.split('/').filter(Boolean);
      // raw.githubusercontent.com / org / repo / branch / ...path
      const org    = parts[0];
      const repo   = parts[1];
      const branch = parts[2];
      const pathParts = parts.slice(0, -1); // drop filename

      baseDir    = `${u.origin}/${pathParts.join('/')}/`;
      rawBase    = `https://raw.githubusercontent.com/${org}/${repo}/${branch}/`;
      repoGHBase = `https://github.com/${org}/${repo}/blob/${branch}/`;
    } catch (_) {
      return; // Can't parse base URL – skip rewriting
    }

    // Images: always resolve to raw content URL
    container.querySelectorAll('img[src]').forEach(img => {
      const src = img.getAttribute('src');
      const resolved = resolveUrl(src, baseDir, rawBase, repoGHBase);
      if (resolved !== src) img.src = resolved;
    });

    // Links: .md → in-app preview; others → absolute GitHub URL in new tab
    container.querySelectorAll('a[href]').forEach(a => {
      const href = a.getAttribute('href');
      if (!href) return;

      // Fragment-only links: map to our heading IDs (already set)
      if (href.startsWith('#')) return;

      // Already absolute
      if (/^[a-z][a-z0-9+\-.]*:/i.test(href)) {
        // Open external links in a new tab safely
        if (!href.startsWith(window.location.origin)) {
          a.target = '_blank';
          a.rel    = 'noopener noreferrer';
        }
        return;
      }

      // Relative .md file link → convert to preview URL
      if (/\.md(#[^?]*)?$/i.test(href)) {
        const fragment = href.match(/(#[^?]*)$/)?.[1] || '';
        const cleanHref = href.replace(/(#[^?]*)$/, '');
        const absRaw = resolveUrl(cleanHref, baseDir, rawBase, repoGHBase);
        a.href   = `${window.location.pathname}?url=${encodeURIComponent(absRaw)}${fragment}`;
        a.target = '_self';
        // Clicking a preview link should re-render, not navigate
        a.addEventListener('click', e => {
          e.preventDefault();
          const newUrl = new URL(a.href);
          const target = newUrl.searchParams.get('url');
          if (target) {
            urlInput.value = target;
            loadMarkdown(target, true);
          }
        });
        return;
      }

      // Other relative links → resolve to GitHub
      const abs = resolveUrl(href, baseDir, repoGHBase, repoGHBase);
      a.href   = abs;
      a.target = '_blank';
      a.rel    = 'noopener noreferrer';
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

  // ── TOC builder ─────────────────────────────────────────────
  function buildTOC(container) {
    tocList.innerHTML = '';
    // Only show h2, h3, h4 in TOC (h1 is the page title)
    container.querySelectorAll('h2, h3, h4').forEach(h => {
      const li = document.createElement('li');
      li.className = `toc-${h.tagName.toLowerCase()}`;
      const a = document.createElement('a');
      a.href = `#${h.id}`;
      a.textContent = h.textContent;
      li.appendChild(a);
      tocList.appendChild(li);
    });
    sidebar.hidden = tocList.children.length === 0;
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
      // raw.githubusercontent.com / org / repo / branch / ...path
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

  // ── Main load function ───────────────────────────────────────
  async function loadMarkdown(inputUrl, pushHistory) {
    const trimmed = (inputUrl || '').trim();
    if (!trimmed) return;

    const rawUrl = toRawUrl(trimmed);

    // Show loading state
    contentArea.innerHTML = '<div class="loading-state"><div class="loading-spinner" aria-hidden="true"></div><span>Loading…</span></div>';
    sidebar.hidden    = true;
    breadcrumb.hidden = true;

    try {
      const res = await fetch(rawUrl);
      if (!res.ok) throw new Error(`${res.status} ${res.statusText}`);
      const raw = await res.text();

      // Pre-process
      const md = stripFrontMatter(raw);

      // Parse markdown → HTML
      const dirtyHtml = marked.parse(md);

      // Sanitize (allow data-lang and target/rel attrs we'll set ourselves)
      const cleanHtml = DOMPurify.sanitize(dirtyHtml, {
        ADD_ATTR: ['target', 'rel', 'data-lang'],
      });

      contentArea.innerHTML = cleanHtml;

      // Post-processing (order matters)
      postProcessCallouts(contentArea);
      buildHeadingIds(contentArea);
      rewriteRelativeUrls(contentArea, rawUrl);
      buildTOC(contentArea);
      addCopyButtons(contentArea);
      setupScrollSpy();
      updateBreadcrumb(rawUrl);
      document.title = contentArea.querySelector('h1')?.textContent
        ? `${contentArea.querySelector('h1').textContent} – Doc Preview`
        : 'Doc Preview';

      // History
      const pageUrl = new URL(window.location.href);
      pageUrl.searchParams.set('url', trimmed);
      if (pushHistory) {
        history.pushState({ url: trimmed }, '', pageUrl);
      } else {
        history.replaceState({ url: trimmed }, '', pageUrl);
      }

      window.scrollTo({ top: 0, behavior: 'instant' });

    } catch (err) {
      contentArea.innerHTML =
        `<div class="error-state">` +
        `<h3>Failed to load document</h3>` +
        `<p>${escapeHtml(err.message)}</p>` +
        `<p>Make sure the URL points to a <strong>public</strong> GitHub markdown file.</p>` +
        `</div>`;
      sidebar.hidden    = true;
      breadcrumb.hidden = true;
    }
  }

  // ── Event listeners ──────────────────────────────────────────
  loadBtn.addEventListener('click', () => loadMarkdown(urlInput.value, true));

  urlInput.addEventListener('keydown', e => {
    if (e.key === 'Enter') loadMarkdown(urlInput.value, true);
  });

  // Back / forward navigation
  window.addEventListener('popstate', e => {
    const url = e.state?.url || new URLSearchParams(window.location.search).get('url');
    if (url) {
      urlInput.value = url;
      loadMarkdown(url, false);
    } else {
      // Reset to welcome screen
      urlInput.value    = '';
      contentArea.innerHTML = document.querySelector('.welcome-screen')?.outerHTML || '';
      sidebar.hidden    = true;
      breadcrumb.hidden = true;
      document.title    = 'Doc Preview – Microsoft Learn Style';
    }
  });

  // ── Auto-load from ?url= param on startup ────────────────────
  const startUrl = new URLSearchParams(window.location.search).get('url');
  if (startUrl) {
    urlInput.value = startUrl;
    // replaceState so the initial load doesn't add a history entry
    loadMarkdown(startUrl, false);
  }

})();
