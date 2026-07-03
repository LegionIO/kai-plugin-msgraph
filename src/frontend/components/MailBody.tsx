import React, { useMemo } from 'react';
import DOMPurify from 'dompurify';

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

  const sanitized = useMemo(() => {
    let src = html;
    // Resolve cid: references to fetched inline attachments.
    if (Object.keys(inlineAttachments).length) {
      src = src.replace(/\bsrc\s*=\s*(["'])cid:([^"']+)\1/gi, (_m, q: string, cid: string) => {
        const key = cid.replace(/^<|>$/g, '');
        const url = inlineAttachments[key] ?? inlineAttachments[`<${key}>`];
        return url ? `src=${q}${url}${q}` : `data-cid=${q}${cid}${q}`;
      });
    }
    return DOMPurify.sanitize(src, {
      USE_PROFILES: { html: true },
      FORBID_TAGS: ['style', 'script', 'iframe', 'object', 'embed', 'form', 'meta', 'link'],
      FORBID_ATTR: ['onerror', 'onload', 'onclick', 'srcset'],
      ALLOW_DATA_ATTR: false,
    });
  }, [html, inlineAttachments]);

  return (
    <div
      className="msg-mailbody"
      onClick={(e) => {
        const a = (e.target as HTMLElement).closest('a');
        if (a && a.getAttribute('href')) {
          e.preventDefault();
          onOpenLink(a.getAttribute('href')!);
        }
      }}
      dangerouslySetInnerHTML={{ __html: sanitized }}
    />
  );
}
