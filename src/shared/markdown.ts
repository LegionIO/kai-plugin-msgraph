// Minimal markdown → Teams-safe HTML for outgoing chat messages.
// Supports **bold**, *italic* / _italic_, ~~strike~~, `inline code`,
// fenced ``` code blocks ```, and line breaks. Everything else is escaped.
// Returns null when the input contains no markdown syntax so callers can
// send plain text instead.

const esc = (s: string) =>
  s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

function inline(src: string): string {
  let out = '';
  let i = 0;
  while (i < src.length) {
    // inline code
    if (src[i] === '`') {
      const j = src.indexOf('`', i + 1);
      if (j > i) {
        out += `<code>${esc(src.slice(i + 1, j))}</code>`;
        i = j + 1;
        continue;
      }
    }
    // **bold**
    if (src.startsWith('**', i)) {
      const j = src.indexOf('**', i + 2);
      if (j > i + 1) {
        out += `<strong>${inline(src.slice(i + 2, j))}</strong>`;
        i = j + 2;
        continue;
      }
    }
    // ~~strike~~
    if (src.startsWith('~~', i)) {
      const j = src.indexOf('~~', i + 2);
      if (j > i + 1) {
        out += `<s>${inline(src.slice(i + 2, j))}</s>`;
        i = j + 2;
        continue;
      }
    }
    // *italic* or _italic_ (single-char, non-greedy to next matching delim on same line)
    if ((src[i] === '*' || src[i] === '_') && src[i + 1] && src[i + 1] !== src[i]) {
      const d = src[i];
      let j = i + 1;
      while (j < src.length && src[j] !== d && src[j] !== '\n') j++;
      if (j < src.length && src[j] === d && j > i + 1) {
        out += `<em>${inline(src.slice(i + 1, j))}</em>`;
        i = j + 1;
        continue;
      }
    }
    out += esc(src[i]);
    i++;
  }
  return out;
}

const MENTION_TOKEN = /@\[([^\]]+)\]\(aad:([0-9a-fA-F-]{36})\)/g;
const MD_TOKENS = /```|`[^`]+`|\*\*|~~|(?<!\*)\*(?!\*)|(?<!_)_(?!_)/;

export function mdToHtml(src: string): string | null {
  if (!MD_TOKENS.test(src)) return null;

  const parts: string[] = [];
  const fence = /```([A-Za-z0-9_+-]*)\n?([\s\S]*?)```/g;
  let last = 0;
  let m: RegExpExecArray | null;
  const flushText = (text: string) => {
    if (!text) return;
    parts.push(inline(text).replace(/\n/g, '<br>'));
  };
  while ((m = fence.exec(src))) {
    flushText(src.slice(last, m.index));
    const lang = m[1] ? ` class="language-${esc(m[1])}"` : '';
    const body = esc(m[2].replace(/\n$/, '')).replace(/\n/g, '<br>');
    parts.push(`<codeblock${lang}><code>${body}</code></codeblock>`);
    last = fence.lastIndex;
  }
  flushText(src.slice(last));
  return parts.join('');
}

export interface PendingImage {
  id: string;
  contentType: string;
  /** base64 (no data: prefix) */
  contentBytes: string;
  name?: string;
}

export interface OutgoingMention {
  id: number;
  mentionText: string;
  mentioned: { user: { id: string; displayName: string; userIdentityType: 'aadUser' } };
}

/** Build a Graph chatMessage body from text (markdown- and @mention-aware) + inline images. */
export function buildMessageBody(text: string, images: PendingImage[] = []): {
  body: { contentType: 'text' | 'html'; content: string };
  hostedContents?: Array<{ '@microsoft.graph.temporaryId': string; contentBytes: string; contentType: string }>;
  mentions?: OutgoingMention[];
} {
  // Extract @[Name](aad:guid) tokens first, replacing with private-use placeholders
  // so the markdown pass doesn't mangle them.
  const mentions: OutgoingMention[] = [];
  const PH = '';
  const withPh = text.replace(MENTION_TOKEN, (_m, name: string, id: string) => {
    const idx = mentions.length;
    mentions.push({
      id: idx,
      mentionText: name,
      mentioned: { user: { id, displayName: name, userIdentityType: 'aadUser' } },
    });
    return `${PH}${idx}${PH}`;
  });

  const html = mdToHtml(withPh);
  const needsHtml = html !== null || mentions.length > 0 || images.length > 0;

  if (!needsHtml) {
    return { body: { contentType: 'text', content: text } };
  }

  let textHtml = html ?? esc(withPh).replace(/\n/g, '<br>');
  textHtml = textHtml.replace(new RegExp(`${PH}(\\d+)${PH}`, 'g'), (_m, i) => {
    const mm = mentions[Number(i)];
    return `<at id="${mm.id}">${esc(mm.mentionText)}</at>`;
  });

  const imgHtml = images
    .map((img) => `<img src="../hostedContents/${img.id}/$value" style="max-width:100%">`)
    .join('');
  const content = [textHtml, imgHtml].filter(Boolean).join(textHtml && imgHtml ? '<br>' : '');

  return {
    body: { contentType: 'html', content },
    ...(mentions.length ? { mentions } : {}),
    ...(images.length
      ? {
          hostedContents: images.map((img) => ({
            '@microsoft.graph.temporaryId': img.id,
            contentBytes: img.contentBytes,
            contentType: img.contentType,
          })),
        }
      : {}),
  };
}
