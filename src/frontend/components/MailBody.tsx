import React, { useMemo, useState, useEffect, useRef } from 'react';
import DOMPurify from 'dompurify';

// 1×1 transparent placeholder shown until the user opts into external content.
const BLANK_PX = 'data:image/gif;base64,R0lGODlhAQABAAAAACH5BAEKAAEALAAAAAABAAEAAAICTAEAOw==';
const isExternal = (u: string) => /^https?:\/\//i.test(u);

let cssInjected = false;
function ensureCss() {
  if (cssInjected || typeof document === 'undefined') return;
  const s = document.createElement('style');
  s.setAttribute('data-msgraph-mail', '');
  // Mail HTML is authored for a white page; rendering it on the app's dark
  // theme is a losing game (dark-on-dark, light-on-white, invisible borders).
  // Do what every mail client does: give the body its own light card.
  s.textContent = `
.msg-mailbody{background:#ffffff;color:#1f2937;border-radius:10px;padding:16px 18px;font-family:inherit;font-size:13px;line-height:1.5;overflow-wrap:break-word;color-scheme:light}
.msg-mailbody *{max-width:100% !important}
.msg-mailbody body,.msg-mailbody html{background:transparent !important}
.msg-mailbody table{border-collapse:collapse;width:auto !important;max-width:100%}
.msg-mailbody td,.msg-mailbody th{border-color:rgba(0,0,0,.15)}
.msg-mailbody img{max-width:100%;height:auto}
.msg-mailbody a{color:#2563eb;text-decoration:underline}
.msg-mailbody p{margin:0 0 .6em}
.msg-mailbody blockquote{border-left:2px solid rgba(0,0,0,.2);margin:.5em 0;padding:.2em 0 .2em .8em}
.msg-mailbody pre,.msg-mailbody code{font-family:ui-monospace,SFMono-Regular,Menlo,monospace}
`;
  document.head.appendChild(s);
  cssInjected = true;
}

let hooked = false;
function ensureHooks() {
  if (hooked) return;
  hooked = true;
  DOMPurify.addHook('afterSanitizeAttributes', (node) => {
    if (node instanceof Element && node.tagName === 'A') {
      node.setAttribute('target', '_blank');
      node.setAttribute('rel', 'noopener noreferrer');
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
  const [blockedCount, setBlockedCount] = useState(0);
  const hostRef = useRef<HTMLDivElement>(null);
  useEffect(() => { setLoadExternal(false); setBlockedCount(0); }, [html]);

  // Sanitize once per message body. All <img> src values are moved to data-src so
  // nothing loads before we've classified it; the effect below decides per-image.
  const sanitized = useMemo(() => {
    const src = html
      .replace(
        /(<img\b[^>]*?\bsrc\s*=\s*)(?:(["'])([^"']*)\2|([^\s>]+))/gi,
        (_m, pre: string, _q, quoted?: string, bare?: string) =>
          `${pre}"${BLANK_PX}" data-src="${(quoted ?? bare ?? '').replace(/"/g, '&quot;')}"`,
      )
      .replace(
        /\b(background(?:-image)?\s*:\s*)url\((["']?)(https?:\/\/[^)"']+)\2\)/gi,
        (_m, p: string, q: string, url: string) => `${p}none/*blocked:${url}*/`,
      );
    return DOMPurify.sanitize(src, {
      USE_PROFILES: { html: true },
      FORBID_TAGS: ['style', 'script', 'iframe', 'object', 'embed', 'form', 'meta', 'link'],
      FORBID_ATTR: ['onerror', 'onload', 'onclick', 'srcset'],
      ADD_ATTR: ['data-src'],
      ALLOW_DATA_ATTR: false,
    });
  }, [html]);

  // Write innerHTML imperatively and only when the sanitized string actually
  // differs, then classify each <img> in place: cid: → inline attachment,
  // http(s) → gated behind loadExternal, everything else (data:, relative) → allow.
  // React re-renders from unrelated state churn never touch this subtree, so text
  // selection survives.
  const lastHtmlRef = useRef<string | null>(null);
  useEffect(() => {
    const host = hostRef.current;
    if (!host) return;
    if (lastHtmlRef.current !== sanitized) {
      host.innerHTML = sanitized;
      lastHtmlRef.current = sanitized;
      // Preserve author-declared aspect ratio so the 1×1 placeholder doesn't
      // inflate to a square under height:auto while the real image loads.
      host.querySelectorAll<HTMLImageElement>('img[data-src]').forEach((img) => {
        const w = Number(img.getAttribute('width'));
        const h = Number(img.getAttribute('height'));
        if (w > 0 && h > 0) img.style.aspectRatio = `${w} / ${h}`;
        else img.style.maxHeight = '160px';
      });
    }
    let blocked = 0;
    host.querySelectorAll<HTMLImageElement>('img[data-src]').forEach((img) => {
      const orig = img.getAttribute('data-src') ?? '';
      let target: string | null = null;
      const cidMatch = /^cid:(.+)$/i.exec(orig);
      if (cidMatch) {
        const key = cidMatch[1].replace(/^<|>$/g, '');
        target = inlineAttachments[key] ?? inlineAttachments[`<${key}>`] ?? null;
      } else if (isExternal(orig)) {
        if (loadExternal) target = orig;
        else blocked++;
      } else if (orig) {
        target = orig;
      }
      if (target && img.getAttribute('src') !== target) {
        img.src = target;
        img.style.maxHeight = '';
      }
    });
    host.querySelectorAll<HTMLElement>('[style*="/*blocked:"]').forEach((el) => {
      if (loadExternal) {
        el.setAttribute('style', (el.getAttribute('style') ?? '').replace(/none\/\*blocked:([^*]+)\*\//g, 'url("$1")'));
      } else {
        blocked++;
      }
    });
    setBlockedCount((prev) => (prev === blocked ? prev : blocked));
  });

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
      />
    </>
  );
}
