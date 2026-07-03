import React, { useMemo, useState, useEffect } from 'react';
import DOMPurify from 'dompurify';

// 1×1 transparent placeholder shown until the user opts into external content.
const BLANK_PX = 'data:image/gif;base64,R0lGODlhAQABAAAAACH5BAEKAAEALAAAAAABAAEAAAICTAEAOw==';
const isExternal = (u: string) => /^https?:\/\//i.test(u);

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
.msg-mailbody [style*="background"],.msg-mailbody [bgcolor]{color:#1f2937}
`;
  document.head.appendChild(s);
  cssInjected = true;
}

let hooked = false;
function ensureHooks() {
  if (hooked) return;
  hooked = true;
  DOMPurify.addHook('afterSanitizeAttributes', (node) => {
    if (node.tagName === 'A') {
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
  useEffect(() => setLoadExternal(false), [html]);

  const { sanitized, blockedCount } = useMemo(() => {
    let src = html;
    // Resolve cid: references to fetched inline attachments.
    if (Object.keys(inlineAttachments).length) {
      src = src.replace(/\bsrc\s*=\s*(["'])cid:([^"']+)\1/gi, (_m, q: string, cid: string) => {
        const key = cid.replace(/^<|>$/g, '');
        const url = inlineAttachments[key] ?? inlineAttachments[`<${key}>`];
        return url ? `src=${q}${url}${q}` : `data-cid=${q}${cid}${q}`;
      });
    }
    // Gate external http(s) images (tracking pixels, remote content) until opted in.
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
      ADD_ATTR: ['data-blocked-src'],
      ALLOW_DATA_ATTR: false,
    });
    return { sanitized: clean, blockedCount: blocked };
  }, [html, inlineAttachments, loadExternal]);

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
