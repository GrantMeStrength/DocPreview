# Doc Preview

Preview GitHub markdown files styled like [Microsoft Learn](https://learn.microsoft.com), without needing to publish them first.

**[Open the app →](https://your-org.github.io/DocPreview/)**

---

## Usage

1. Paste a GitHub markdown file URL into the bar at the top
2. Click **Preview**
3. The rendered preview loads with Microsoft Learn styling
4. Share the URL in your browser — it includes `?url=…` so others see the same document

### Supported URL formats

```
https://github.com/{org}/{repo}/blob/{branch}/{path/to/file.md}
https://raw.githubusercontent.com/{org}/{repo}/{branch}/{path/to/file.md}
```

> [!NOTE]
> Only **public** GitHub repositories are supported. Private repos will return a 404.

---

## Features

- **Microsoft Learn styling** — typography, colors, callout boxes, and code blocks that match learn.microsoft.com
- **Callouts** — renders `[!NOTE]`, `[!TIP]`, `[!WARNING]`, `[!IMPORTANT]`, and `[!CAUTION]` blockquotes as styled callout boxes
- **Syntax highlighting** — code fences are highlighted using highlight.js
- **Auto-generated TOC** — in-page table of contents from headings, with scroll-tracking
- **Relative link resolution** — images and links resolve correctly against the source document's location; `.md` links open in the preview
- **Shareable URLs** — the `?url=` parameter is automatically set, so any preview link is shareable
- **XSS protection** — rendered HTML is sanitized via DOMPurify before display

---

## Deploying to GitHub Pages

1. Fork or clone this repo
2. Push to GitHub
3. Go to **Settings → Pages** and set the source to **Deploy from a branch → `main` / `/ (root)`**
4. Your app will be live at `https://{your-username}.github.io/{repo-name}/`

---

## Local development

No build step required. Open `index.html` directly in a browser, or serve with any static server:

```bash
npx serve .
# or
python3 -m http.server 8080
```
