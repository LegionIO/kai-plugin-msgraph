import React, { useMemo, useState, useEffect, useRef } from 'react';
import DOMPurify from 'dompurify';

// 1×1 transparent placeholder shown until the user opts into external content.
const BLANK_PX = 'data:image/gif;base64,R0lGODlhAQABAAAAACH5BAEKAAEALAAAAAABAAEAAAICTAEAOw==';
const isExternal = (u: string) => /^https?:\/\//i.test(u);

/** Rewrite near-black inline `color:` values to `inherit` so mail authored for a
 *  white background stays readable in dark mode; leave chromatic colours alone. */
function neutralizeDarkText(style: string): string {
  return style.replace(
    /(^|;)\s*color\s*:\s*([^;]+)/gi,
    (m, sep: string, val: string) => (isNearBlack(val.trim()) ? `${sep}color:inherit` : m),
  );
}
function isNearBlack(v: string): boolean {
  const s = v.toLowerCase().replace(/\s+/g, '');
  if (s === 'black' || s === 'currentcolor' || s === 'windowtext') return true;
  let r = 255, g = 255, b = 255;
  let m = /^#([0-9a-f]{3})$/.exec(s);
  if (m) { r = parseInt(m[1][0], 16) * 17; g = parseInt(m[1][1], 16) * 17; b = parseInt(m[1][2], 16) * 17; }
  else if ((m = /^#([0-9a-f]{6})([0-9a-f]{2})?$/.exec(s))) { r = parseInt(m[1].slice(0, 2), 16); g = parseInt(m[1].slice(2, 4), 16); b = parseInt(m[1].slice(4, 6), 16); }
  else if ((m = /^rgba?\((\d+),(\d+),(\d+)/.exec(s))) { r = +m[1]; g = +m[2]; b = +m[3]; }
  else return false;
  // Rec.709 luma; ~0.25 threshold catches greys up to ~#404040 while keeping brand colours.
  return (0.2126 * r + 0.7152 * g + 0.0722 * b) / 255 < 0.25;
}

let cssInjected = false;
function ensureCss() {
  if (cssInjected || typeof document === 'undefined') return;
  const s = document.createElement('style');
  s.setAttribute('data-msgraph-mail', '');
  s.textContent = `
.msg-mailbody{color:inherit;font-family:inherit;font-size:13px;line-height:1.5;overflow-wrap:break-word}
.msg-mailbody *{max-width:100% !important}
.msg-mailbody body,.msg-mailbody html{background:transparent !important;color:inherit !important}
.msg-mailbody table{border-collapse:collapse;width:auto !important;max-width:100%}
.msg-mailbody td,.msg-mailbody th{border-color:rgba(127,127,127,.3)}
.msg-mailbody img{max-width:100%;height:auto}
.msg-mailbody a{color:#2563eb;text-decoration:underline}
.dark .msg-mailbody a{color:#60a5fa}
.msg-mailbody p{margin:0 0 .6em}
.msg-mailbody blockquote{border-left:2px solid rgba(127,127,127,.4);margin:.5em 0;padding:.2em 0 .2em .8em;color:inherit}
.msg-mailbody pre,.msg-mailbody code{font-family:ui-monospace,SFMono-Regular,Menlo,monospace}
.msg-mailbody font[color]{color:inherit !important}
`;
  document.head.appendChild(s);
  cssInjected = true;
}

let hooked = false;
function ensureHooks() {
  if (hooked) return;
  hooked = true;
  DOMPurify.addHook('afterSanitizeAttributes', (node) => {
    if (!(node instanceof Element)) return;
    if (node.tagName === 'A') {
      node.setAttribute('target', '_blank');
      node.setAttribute('rel', 'noopener noreferrer');
    }
    if (node.hasAttribute('color')) node.removeAttribute('color');
    const style = node.getAttribute('style');
    if (style && /color\s*:/i.test(style)) {
      node.setAttribute('style', neutralizeDarkText(style));
    }
  });
}

export function MailBody({
  html,
  inlineAttachments,
  onOpenLink,
}: {
  html: string;
  inlineAttachments: Record<string, string | null>;
  onOpenLink: (url: string) => void;
}) {
  ensureCss();
  ensureHooks();
  const [loadExternal, setLoadExternal] = useState(false);
  const hostRef = useRef<HTMLDivElement>(null);
  useEffect(() => setLoadExternal(false), [html]);

  // Sanitize once per (html, loadExternal). cid: images are tagged as data-cid and
  // resolved via a DOM patch effect below so incoming inline attachments don't
  // replace the whole innerHTML (which would drop the user's text selection).
  const { sanitized, blockedCount } = useMemo(() => {
    let src = html.replace(
      /\bsrc\s*=\s*(["'])cid:([^"']+)\1/gi,
      (_m, q: string, cid: string) => `src=${q}${BLANK_PX}${q} data-cid=${q}${cid.replace(/^<|>$/g, '')}${q}`,
    );
    let blocked = 0;
    if (!loadExternal) {
      src = src.replace(
        /(<img\b[^>]*\bsrc\s*=\s*)(["'])(https?:\/\/[^"']+)\2/gi,
        (_m, pre: string, q: string, url: string) => {
          blocked++;
          return `${pre}${q}${BLANK_PX}${q} data-blocked-src=${q}${url}${q}`;
        },
      );
      src = src.replace(/\b(background(?:-image)?\s*:\s*)url\((["']?)(https?:\/\/[^)"']+)\2\)/gi, (_m, pre: string) => {
        blocked++;
        return `${pre}none`;
      });
    }
    const clean = DOMPurify.sanitize(src, {
      USE_PROFILES: { html: true },
      FORBID_TAGS: ['style', 'script', 'iframe', 'object', 'embed', 'form', 'meta', 'link'],
      FORBID_ATTR: ['onerror', 'onload', 'onclick', 'srcset'],
      ADD_ATTR: ['data-blocked-src', 'data-cid'],
      ALLOW_DATA_ATTR: false,
    });
    return { sanitized: clean, blockedCount: blocked };
  }, [html, loadExternal]);

  // Patch cid: images in place as their bytes arrive — no innerHTML churn.
  useEffect(() => {
    const host = hostRef.current;
    if (!host) return;
    host.querySelectorAll<HTMLImageElement>('img[data-cid]').forEach((img) => {
      const cid = img.getAttribute('data-cid') ?? '';
      const url = inlineAttachments[cid] ?? inlineAttachments[`<${cid}>`];
      if (url && img.src !== url) img.src = url;
    });
  }, [sanitized, inlineAttachments]);

  return (
    <>
      {!loadExternal && blockedCount > 0 && (
        <div className="mb-3 flex items-center justify-between gap-3 rounded-lg border border-border bg-muted/40 px-3 py-2 text-[11px]">
          <span className="text-muted-foreground">
            {blockedCount} external image{blockedCount === 1 ? '' : 's'} blocked to prevent tracking.
          </span>
          <button
            type="button"
            onClick={() => setLoadExternal(true)}
            className="px-2.5 py-1 text-[11px] font-medium bg-primary text-primary-foreground rounded-md hover:bg-primary/90"
          >
            Load images
          </button>
        </div>
      )}
      <div
        ref={hostRef}
        className="msg-mailbody"
        onClick={(e) => {
          const a = (e.target as HTMLElement).closest('a');
          if (a && a.getAttribute('href')) {
            e.preventDefault();
            const href = a.getAttribute('href')!;
            if (isExternal(href) || href.startsWith('mailto:')) onOpenLink(href);
          }
        }}
        dangerouslySetInnerHTML={{ __html: sanitized }}
      />
    </>
  );
}
