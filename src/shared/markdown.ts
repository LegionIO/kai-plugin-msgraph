// Markdown → Teams-safe HTML for outgoing chat messages.
// Span-level: **bold**, *italic* / _italic_, ~~strike~~, `inline code`,
//   [text](url) links, ![alt](url) images, <autolinks>, and bare http(s) URLs.
// Block-level: fenced ``` code blocks ```, # / setext headings, > blockquotes
//   (with lazy continuation), -/*/+ and 1. lists, - [ ] / - [x] task lists,
//   GFM | pipe | tables |, --- horizontal rules, hard line breaks (2+ trailing
//   spaces or \), and soft line breaks. Everything else is escaped, and every
//   URL is scheme-gated (http/https/mailto/tel/relative only). Returns null when
//   the input contains no markdown syntax so callers can send plain text instead.

import { highlight } from './highlight.js';

const esc = (s: string) =>
  s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

// Private-use character used by buildMessageBody() to hold @mention slots while
// markdown is rendered. It must never survive into an HTML attribute (href/src/
// alt/title), or the later placeholder→<at> restore would inject a tag inside a
// quoted attribute and produce malformed HTML. Attribute escaping strips it, and
// safeUrl() rejects URLs containing it.
const MENTION_PH = '\uE100';

// Regex matching a full mention placeholder token (PH + digits + PH) plus any
// stray lone sentinel chars, so attributes never keep a numeric artifact like
// alt="hi 0" from a mention that appeared inside a URL/alt/title.
const MENTION_PH_TOKEN = new RegExp(`${MENTION_PH}\\d+${MENTION_PH}|${MENTION_PH}`, 'g');

// Escape for use inside a double-quoted HTML attribute. In addition to the
// text escapes, this neutralizes quotes so a value can never break out of the
// surrounding attribute (defense-in-depth alongside safeUrl's rejection of
// quotes/whitespace in URLs), and drops any mention placeholder tokens/chars.
export const escAttr = (s: string) =>
  esc(s.replace(MENTION_PH_TOKEN, '')).replace(/"/g, '&quot;').replace(/'/g, '&#39;');

// Only allow safe URL schemes in links/autolinks. Anything else (javascript:,
// data:, vbscript:, etc.) is rendered as plain text to prevent injection.
// A URL is also rejected outright if it contains characters that must never
// appear raw in an href/src attribute (quotes, whitespace, angle brackets,
// backtick, C0 control chars, or a mention placeholder) — those indicate an
// attribute-breakout attempt or a nonsensical mention-inside-URL.
export function safeUrl(raw: string): string | null {
  const url = raw.trim();
  if (!url) return null;
  // Reject anything that could break out of a quoted attribute or inject markup.
  // eslint-disable-next-line no-control-regex
  if (/["'`<>\s]|[\u0000-\u001f]/.test(url) || url.includes(MENTION_PH)) return null;
  // Protocol-relative (//host) → force https so it's an absolute, openable URL.
  // Must be checked BEFORE the generic "/" relative case below (which would
  // otherwise match "//..." first and emit an href the app can't open).
  if (/^\/\//.test(url)) return `https:${url}`;
  // Relative / anchor links are fine.
  if (/^(#|\/|\.\/|\.\.\/)/.test(url)) return url;
  const scheme = /^([a-zA-Z][a-zA-Z0-9+.-]*):/.exec(url);
  if (!scheme) {
    // No scheme and not an obvious path → treat www.* as https, else reject.
    return /^www\./i.test(url) ? `https://${url}` : null;
  }
  const s = scheme[1].toLowerCase();
  return s === 'http' || s === 'https' || s === 'mailto' || s === 'tel' ? url : null;
}

// Stricter allowlist for image `src` values: only sources that make sense (and
// are safe) as an image — http(s), protocol-relative, Graph hostedContents-style
// relative paths, and `cid:` (mail inline). Rejects mailto:/tel:/#anchors/data:/
// javascript:/file: etc. that safeUrl would otherwise permit for links.
export function safeImageUrl(raw: string): string | null {
  const url = raw.trim();
  if (!url) return null;
  // eslint-disable-next-line no-control-regex
  if (/["'`<>\s]|[\u0000-\u001f]/.test(url) || url.includes(MENTION_PH)) return null;
  if (/^\/\//.test(url)) return `https:${url}`;             // protocol-relative → https
  if (/^(\/|\.\/|\.\.\/)/.test(url)) return url;            // relative path (hostedContents)
  const scheme = /^([a-zA-Z][a-zA-Z0-9+.-]*):/.exec(url);
  if (!scheme) return null;                                 // bare/anchor → not an image src
  const s = scheme[1].toLowerCase();
  return s === 'http' || s === 'https' || s === 'cid' ? url : null;
}

// Max characters a single label/URL scan will look ahead. Bounding the scan
// keeps the hand-written parser linear even on adversarial bracket/paren runs
// (each start does O(SCAN_WINDOW) work, so total is O(n)) without the fragile
// "failure watermark" heuristic, which mis-fired on inputs like "[[ok](#x)".
const SCAN_WINDOW = 2048;
// Cap inline recursion (link labels render their inner markdown) so deeply
// nested links like "[[[…](#x)](#x)…" cannot overflow the stack.
const MAX_INLINE_DEPTH = 24;

function anchor(text: string, url: string, depth = 0, title = ''): string {
  const safe = safeUrl(url);
  // Render the label with link/autolink parsing disabled so a URL-like label
  // (e.g. [https://a](https://b)) can't emit a nested <a> inside this anchor.
  const label = depth >= MAX_INLINE_DEPTH ? esc(text) : inline(text, depth + 1, true);
  if (!safe) return label; // unsafe scheme → drop href, keep the visible text
  // Only emit a clickable link for schemes that actually resolve when opened from
  // a chat message. Relative/anchor targets (#, /, ./, ../) have no base document
  // here, so the app's link opener would just error — render them as text instead.
  if (!/^(https?:|mailto:|tel:)/i.test(safe)) return label;
  const titleAttr = title ? ` title="${escAttr(title)}"` : '';
  return `<a href="${escAttr(safe)}"${titleAttr}>${label}</a>`;
}


function inline(src: string, depth = 0, noLinks = false): string {
  let out = '';
  let i = 0;
  // Monotonic lookahead memos for the sentinels `](`, `)`, `>`. Each caches the
  // next sentinel position at/after the last query; because `i` only advances,
  // indexOf is recomputed only after `i` passes a cached hit, and a miss (-1) is
  // cached permanently. Total lookahead work is therefore O(n) — a single
  // distant sentinel can't make repeated delimiters O(n^2).
  let nextBracketLink = -2; // -2 = stale/unknown; -1 = none for rest of string
  let nextParenClose = -2;
  let nextAngle = -2;
  const hasBracketLink = (from: number): boolean => {
    if (nextBracketLink === -1) return false;
    if (nextBracketLink < from) nextBracketLink = src.indexOf('](', from);
    return nextBracketLink !== -1;
  };
  const hasParenClose = (from: number): boolean => {
    if (nextParenClose === -1) return false;
    if (nextParenClose < from) nextParenClose = src.indexOf(')', from);
    return nextParenClose !== -1;
  };
  const hasAngleClose = (from: number): boolean => {
    if (nextAngle === -1) return false;
    if (nextAngle < from) nextAngle = src.indexOf('>', from);
    return nextAngle !== -1;
  };
  // Scan a depth-balanced bracket span [ ... ] (nested brackets allowed),
  // starting just after the opening bracket at `labelStart`. Returns the index
  // of the matching close bracket, or -1 (stops at newline / end / window).
  const matchBracket = (labelStart: number): number => {
    let depth2 = 1;
    const stop = Math.min(src.length, labelStart + SCAN_WINDOW);
    for (let k = labelStart; k < stop; k++) {
      const c = src[k];
      if (c === '[') depth2++;
      else if (c === ']') { if (--depth2 === 0) return k; }
      else if (c === '\n') break;
    }
    return -1;
  };
  // Scan a depth-balanced paren span ( ... ), starting just after '('.
  const matchParen = (urlStart: number): number => {
    let depth2 = 1;
    const stop = Math.min(src.length, urlStart + SCAN_WINDOW);
    for (let k = urlStart; k < stop; k++) {
      const c = src[k];
      if (c === '(') depth2++;
      else if (c === ')') { if (--depth2 === 0) return k; }
      else if (c === '\n') break;
    }
    return -1;
  };
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
    // ![alt](url) image — must be checked before the [text](url) link branch.
    if (src[i] === '!' && src[i + 1] === '[' && hasBracketLink(i + 2)) {
      const close = matchBracket(i + 2);
      if (close !== -1 && src[close + 1] === '(' && hasParenClose(close + 2)) {
        const end = matchParen(close + 2);
        if (end > close) {
          const alt = src.slice(i + 2, close);
          // Optional "title" after the URL: (url "title")
          const rawUrl = src.slice(close + 2, end).trim();
          const um = /^(\S+)(?:\s+["'(](.*)["')])?$/.exec(rawUrl);
          const url = um ? um[1] : rawUrl;
          const title = um && um[2] ? um[2] : '';
          const safe = safeImageUrl(url);
          if (safe) {
            const titleAttr = title ? ` title="${escAttr(title)}"` : '';
            out += `<img src="${escAttr(safe)}" alt="${escAttr(alt)}"${titleAttr}>`;
          } else {
            // Unsafe/non-image-scheme src → show alt text only.
            out += esc(alt);
          }
          i = end + 1;
          continue;
        }
      }
    }
    // [text](url) link
    if (!noLinks && src[i] === '[' && hasBracketLink(i + 1)) {
      const close = matchBracket(i + 1);
      if (close !== -1 && src[close + 1] === '(' && hasParenClose(close + 2)) {
        const end = matchParen(close + 2);
        if (end > close) {
          const text = src.slice(i + 1, close);
          // Support an optional title: [text](url "title"). Pass only the bare
          // URL to safeUrl (otherwise the whitespace before the title would make
          // it reject the whole thing and drop the href).
          const rawUrl = src.slice(close + 2, end).trim();
          const um = /^(\S+)(?:\s+["'(]([\s\S]*)["')])?$/.exec(rawUrl);
          const url = um ? um[1] : rawUrl;
          const title = um && um[2] ? um[2] : '';
          out += anchor(text, url, depth, title);
          i = end + 1;
          continue;
        }
      }
    }
    // <https://autolink>
    if (!noLinks && src[i] === '<' && hasAngleClose(i + 1)) {
      // Search for the closing '>' only within a bounded window so repeated '<'
      // with a distant '>' can't be O(n^2); a real autolink is short anyway.
      const winEnd = Math.min(src.length, i + 1 + SCAN_WINDOW);
      let end = -1;
      for (let k = i + 1; k < winEnd; k++) {
        if (src[k] === '>') { end = k; break; }
        if (src[k] === '\n') break;
      }
      if (end > i) {
        const inner = src.slice(i + 1, end);
        const href = safeUrl(inner);
        if (href && /^[a-zA-Z][\w+.-]*:|^www\./.test(inner)) {
          out += `<a href="${escAttr(href)}">${esc(inner)}</a>`;
          i = end + 1;
          continue;
        }
      }
    }
    // **bold**
    if (src.startsWith('**', i)) {
      const j = src.indexOf('**', i + 2);
      if (j > i + 1) {
        const body = depth >= MAX_INLINE_DEPTH ? esc(src.slice(i + 2, j)) : inline(src.slice(i + 2, j), depth + 1, noLinks);
        out += `<strong>${body}</strong>`;
        i = j + 2;
        continue;
      }
    }
    // ~~strike~~
    if (src.startsWith('~~', i)) {
      const j = src.indexOf('~~', i + 2);
      if (j > i + 1) {
        const body = depth >= MAX_INLINE_DEPTH ? esc(src.slice(i + 2, j)) : inline(src.slice(i + 2, j), depth + 1, noLinks);
        out += `<s>${body}</s>`;
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
        const body = depth >= MAX_INLINE_DEPTH ? esc(src.slice(i + 1, j)) : inline(src.slice(i + 1, j), depth + 1, noLinks);
        out += `<em>${body}</em>`;
        i = j + 1;
        continue;
      }
    }
    // Bare URL autolink: http(s)://... or www.... Use a sticky match at `i` so
    // we never allocate src.slice(i) per character (that would be O(n^2) on a
    // long run of 'h'/'w'). The single-char prefilter keeps the common case cheap.
    if (!noLinks && (src[i] === 'h' || src[i] === 'H' || src[i] === 'w' || src[i] === 'W')) {
      BARE_URL_STICKY.lastIndex = i;
      const m = BARE_URL_STICKY.exec(src);
      if (m && m.index === i) {
        // Allow ')' inside the run so balanced parens survive (Wikipedia-style
        // https://e/foo_(bar)); trailing punctuation is trimmed afterward.
        const raw = m[0];
        // Count parens once (not per-iteration) so a long run of trailing ')'
        // can't make this O(n^2). Trim trailing punctuation via a moving end
        // index rather than repeated slicing.
        let opens = 0;
        let closes = 0;
        for (let k = 0; k < raw.length; k++) {
          if (raw[k] === '(') opens++;
          else if (raw[k] === ')') closes++;
        }
        let end = raw.length;
        for (;;) {
          const last = raw[end - 1];
          if (last === ')') {
            // Trailing ')' belongs to the link only if it has more '(' than ')'
            // (GFM); otherwise it closes surrounding prose like "(see …)".
            if (closes > opens) { closes--; end--; continue; }
            break;
          }
          if (last !== undefined && '.,;:!?]}\'"'.includes(last)) { end--; continue; }
          break;
        }
        const core = raw.slice(0, end);
        const suffix = raw.slice(end);
        const href = safeUrl(core);
        if (href) {
          out += `<a href="${escAttr(href)}">${esc(core)}</a>${esc(suffix)}`;
          i += core.length + suffix.length;
          continue;
        }
      }
    }
    out += esc(src[i]);
    i++;
  }
  return out;
}

const MENTION_TOKEN = /@\[([^\]]+)\]\(aad:([0-9a-fA-F-]{36})\)/g;
// Presence of any of these means the text is worth rendering as HTML.
const MD_TOKENS =
  /```|`[^`]+`|\*\*|~~|(?<!\*)\*(?!\*)|(?<!_)_(?!_)|\]\(|<(?:https?|mailto|tel):|\bhttps?:\/\/|\bwww\.|^\s{0,3}#{1,6}\s|^\s{0,3}>|^\s{0,3}([-*+]|\d+[.)])\s|^\s{0,3}([-*_])\1\1[-*_\s]*$|^\s{0,3}\|.*\|\s*$|^\s{0,3}\|?\s*:?-{2,}:?\s*\||^.+\n\s{0,3}(=+|-+)\s*$/im;

// Sticky (anchored-at-lastIndex) bare-URL matcher. Using `y` + `lastIndex = i`
// avoids `src.slice(i)` allocations in the hot inline() loop (which would be
// O(n^2) on a long run of 'h'/'w' characters once HTML mode is active).
// Excludes the same chars safeUrl() rejects (quotes, backtick, angle brackets,
// brackets/braces, C0 controls) so we never form a match that safeUrl would
// reject and then re-scan — which is quadratic on runs like "www.\x01www.\x01…".
// eslint-disable-next-line no-control-regex
const BARE_URL_STICKY = /(?:https?:\/\/|www\.)[^\s<>"'`\]}\u0000-\u001f]+/iy;

// Split a GFM table row "| a | b |" into trimmed cell strings (\| escapes a pipe).
function splitTableRow(row: string): string[] {
  let s = row.trim();
  if (s.startsWith('|')) s = s.slice(1);
  if (s.endsWith('|') && !s.endsWith('\\|')) s = s.slice(0, -1);
  const cells: string[] = [];
  let cur = '';
  for (let k = 0; k < s.length; k++) {
    if (s[k] === '\\' && s[k + 1] === '|') { cur += '|'; k++; continue; }
    if (s[k] === '|') { cells.push(cur); cur = ''; continue; }
    cur += s[k];
  }
  cells.push(cur);
  return cells.map((c) => c.trim());
}

// A GFM table delimiter row: |---|:--:|--:| etc.
// A GFM delimiter row: one or more ":?---:?" cells separated/bounded by pipes.
// Requires at least one literal "|" (so a bare "---" stays an HR) and supports
// single-column tables like "| --- |".
const TABLE_DELIM = /^\s{0,3}\|\s*:?-+:?\s*(\|\s*:?-+:?\s*)*\|?\s*$|^\s{0,3}:?-+:?\s*(\|\s*:?-+:?\s*)+\|?\s*$/;

// True if a line begins a block-level construct (used for lazy-continuation checks).
function isBlockStart(ln: string): boolean {
  return /^\s{0,3}(#{1,6}\s|>|([-*+]|\d+[.)])\s|([-*_])(\s*\3){2,}\s*$)/.test(ln);
}

// Render a run of "loose" (non-fenced) markdown lines into block-level HTML.
// `depth` guards against attacker-controlled blockquote nesting (e.g. ">>>>…"),
// which would otherwise recurse once per level and can exhaust the stack.
const MAX_BLOCK_DEPTH = 8;
function renderBlocks(src: string, depth = 0): string {
  const lines = src.split('\n');
  const out: string[] = [];
  let i = 0;

  while (i < lines.length) {
    const line = lines[i];

    // GFM table: "| h | h |" header, then a delimiter row, then body rows.
    if (line.includes('|') && i + 1 < lines.length && TABLE_DELIM.test(lines[i + 1])) {
      const header = splitTableRow(line);
      const aligns = splitTableRow(lines[i + 1]).map((c) => {
        const l = c.startsWith(':');
        const r = c.endsWith(':');
        return l && r ? 'center' : r ? 'right' : l ? 'left' : '';
      });
      i += 2;
      const bodyRows: string[][] = [];
      while (i < lines.length && lines[i].includes('|') && lines[i].trim() !== '') {
        bodyRows.push(splitTableRow(lines[i]));
        i++;
      }
      const alignAttr = (idx: number) =>
        aligns[idx] ? ` style="text-align:${aligns[idx]}"` : '';
      const thead = `<thead><tr>${header
        .map((c, idx) => `<th${alignAttr(idx)}>${inline(c)}</th>`)
        .join('')}</tr></thead>`;
      const tbody = `<tbody>${bodyRows
        .map(
          (r) =>
            `<tr>${header
              .map((_, idx) => `<td${alignAttr(idx)}>${inline(r[idx] ?? '')}</td>`)
              .join('')}</tr>`,
        )
        .join('')}</tbody>`;
      out.push(`<table>${thead}${tbody}</table>`);
      continue;
    }

    // Horizontal rule: ---, ***, ___ (3+), possibly spaced.
    if (/^\s{0,3}([-*_])(\s*\1){2,}\s*$/.test(line)) {
      out.push('<hr>');
      i++;
      continue;
    }

    // ATX heading: # .. ######  (trailing hashes trimmed procedurally to avoid
    // ambiguous-regex quadratic backtracking on long heading-like lines).
    const h = /^\s{0,3}(#{1,6})\s+(.*)$/.exec(line);
    if (h) {
      // Cap at h3: parseHtmlBody / BodySegment only model levels 1–3, so h4–h6
      // would lose their heading formatting on reload. Teams treats h1–h3 as the
      // meaningful chat heading sizes anyway.
      const level = Math.min(3, h[1].length);
      // Procedural right-trim of an optional closing "###" sequence: strip
      // trailing spaces, then trailing '#', then trailing spaces — O(n), no
      // backtracking (a regex like /\s+#+\s*$/ is quadratic on "   ###...x").
      let e = h[2].length;
      while (e > 0 && (h[2][e - 1] === ' ' || h[2][e - 1] === '\t')) e--;
      const beforeHashes = e;
      while (e > 0 && h[2][e - 1] === '#') e--;
      // Only treat the '#' run as a closing sequence if it was preceded by
      // whitespace (or is the whole line); otherwise keep it as content.
      if (e > 0 && h[2][e - 1] !== ' ' && h[2][e - 1] !== '\t') e = beforeHashes;
      else while (e > 0 && (h[2][e - 1] === ' ' || h[2][e - 1] === '\t')) e--;
      const body = h[2].slice(0, e);
      out.push(`<h${level}>${inline(body)}</h${level}>`);
      i++;
      continue;
    }

    // Setext heading: a text line underlined by === (h1) or --- (h2).
    if (
      line.trim() !== '' &&
      !isBlockStart(line) &&
      !line.includes('|') &&
      i + 1 < lines.length &&
      /^\s{0,3}(=+|-+)\s*$/.test(lines[i + 1])
    ) {
      const level = lines[i + 1].trim()[0] === '=' ? 1 : 2;
      out.push(`<h${level}>${inline(line.trim())}</h${level}>`);
      i += 2;
      continue;
    }

    // Blockquote: consecutive ">" lines, plus lazy-continuation text lines.
    if (/^\s{0,3}>/.test(line)) {
      const quoted: string[] = [];
      while (i < lines.length) {
        if (/^\s{0,3}>/.test(lines[i])) {
          quoted.push(lines[i].replace(/^\s{0,3}>\s?/, ''));
          i++;
        } else if (lines[i].trim() !== '' && !isBlockStart(lines[i])) {
          quoted.push(lines[i]); // lazy continuation
          i++;
        } else {
          break;
        }
      }
      const inner = quoted.join('\n');
      if (depth >= MAX_BLOCK_DEPTH) {
        // Nesting cap reached: stop recursing, emit the remaining quoted text
        // as inline HTML (still escaped) instead of deeper <blockquote> levels.
        out.push(`<blockquote>${inline(inner.replace(/\n/g, ' '))}</blockquote>`);
      } else {
        out.push(`<blockquote>${renderBlocks(inner, depth + 1)}</blockquote>`);
      }
      continue;
    }

    // Lists: unordered (-, *, +) or ordered (1. / 1)); supports GFM task items.
    // Emits valid nesting: a deeper list is placed *inside* the parent <li>,
    // e.g. `<ul><li>a<ul><li>b</li></ul></li></ul>`.
    const li = /^(\s{0,3})([-*+]|\d+[.)])\s+(.*)$/.exec(line);
    if (li) {
      const stack: Array<'ul' | 'ol'> = [];
      const indents: number[] = [];
      const buf: string[] = [];
      // Whether the current innermost <li> is still open (awaiting </li>).
      let liOpen = false;
      const renderItem = (raw: string): string => {
        const task = /^\[([ xX])\]\s+(.*)$/.exec(raw);
        if (task) {
          const checked = task[1].toLowerCase() === 'x';
          // Use unicode glyphs (☑/☐) rather than an <input> form control: Teams'
          // HTML sanitizer strips form controls and parseHtmlBody has no <input>
          // handling, so a checkbox wouldn't survive the send/reload round-trip.
          const box = checked ? '\u2611' : '\u2610'; // ☑ / ☐
          return `${box} ${inline(task[2])}`;
        }
        return inline(raw);
      };
      while (i < lines.length) {
        const mm = /^(\s*)([-*+]|\d+[.)])\s+(.*)$/.exec(lines[i]);
        if (!mm) break;
        const indent = mm[1].length;
        const ordered = /\d/.test(mm[2]);
        const kind: 'ul' | 'ol' = ordered ? 'ol' : 'ul';
        // Preserve a non-1 ordered-list start (e.g. "3." → <ol start="3">) so
        // Teams doesn't renumber the user's visible list from 1.
        const openTag = (): string => {
          if (kind === 'ol') {
            const n = parseInt(mm[2], 10);
            if (Number.isFinite(n) && n !== 1) return `<ol start="${n}">`;
            return '<ol>';
          }
          return '<ul>';
        };
        if (!stack.length || indent > indents[indents.length - 1]) {
          // Descend: nest the new list inside the currently-open <li> (do not
          // close that <li> — the sublist becomes part of it).
          stack.push(kind);
          indents.push(indent);
          buf.push(openTag());
        } else {
          // Same or shallower level: close the previous item first.
          if (liOpen) { buf.push('</li>'); liOpen = false; }
          while (
            stack.length > 1 &&
            indent < indents[indents.length - 1]
          ) {
            buf.push(`</${stack.pop()}>`);
            indents.pop();
            buf.push('</li>'); // close the parent <li> the sublist lived in
          }
          // Same indent but the marker type changed (e.g. "1." → "-"): close the
          // current list and open one of the new kind so ordered/unordered don't
          // get merged into a single <ol>/<ul>.
          if (stack.length && kind !== stack[stack.length - 1] && indent === indents[indents.length - 1]) {
            buf.push(`</${stack.pop()}>`);
            indents.pop();
            stack.push(kind);
            indents.push(indent);
            buf.push(openTag());
          }
        }
        buf.push(`<li>${renderItem(mm[3])}`);
        liOpen = true;
        i++;
      }
      // Unwind: close the open item, then each list level (each nested list is
      // inside a parent <li> that must also be closed).
      if (liOpen) buf.push('</li>');
      while (stack.length) {
        buf.push(`</${stack.pop()}>`);
        if (stack.length) buf.push('</li>');
      }
      out.push(buf.join(''));
      continue;
    }

    // Blank line → paragraph break.
    if (/^\s*$/.test(line)) {
      out.push('');
      i++;
      continue;
    }

    // Regular text: coalesce consecutive plain lines. A line ending in 2+ spaces
    // or a backslash is a hard break; every newline becomes <br> in chat anyway.
    const para: string[] = [];
    while (
      i < lines.length &&
      !/^\s*$/.test(lines[i]) &&
      !isBlockStart(lines[i]) &&
      !(lines[i].includes('|') && i + 1 < lines.length && TABLE_DELIM.test(lines[i + 1])) &&
      !(
        !lines[i].includes('|') &&
        i + 1 < lines.length &&
        /^\s{0,3}(=+|-+)\s*$/.test(lines[i + 1])
      )
    ) {
      para.push(inline(lines[i].replace(/( {2,}|\\)$/, '')));
      i++;
    }
    out.push(para.join('<br>'));
  }

  // Join blocks with a single <br>; a blank-line marker ('') between two blocks
  // becomes an extra <br> so paragraph breaks (e.g. "**T**\n\nBody") are kept.
  // Leading/trailing and collapsed multiple blank markers don't add stray breaks.
  let result = '';
  let pendingBlank = false;
  let wrote = false;
  for (const seg of out) {
    if (seg === '') { if (wrote) pendingBlank = true; continue; }
    if (wrote) result += pendingBlank ? '<br><br>' : '<br>';
    result += seg;
    wrote = true;
    pendingBlank = false;
  }
  return result;
}

export function mdToHtml(src: string): string | null {
  if (!MD_TOKENS.test(src)) return null;

  const parts: string[] = [];
  const flushText = (text: string) => {
    const rendered = text ? renderBlocks(text) : '';
    // Preserve a line break between prose and an adjacent code block when the
    // source had a newline at that boundary (renderBlocks trims edge blanks, so
    // otherwise "a\n\n```…```" would render as "a<codeblock…>" with no gap).
    const leadBreak = rendered && /^\s*\n/.test(text) ? '<br>' : '';
    const trailBreak = rendered && /\n\s*$/.test(text) ? '<br>' : '';
    if (rendered) parts.push(`${leadBreak}${rendered}${trailBreak}`);
  };

  // Procedural fence scan (linear). A regex like /```(...)([\s\S]*?)```/ is
  // quadratic on an unterminated fence with a long info-string run, so we scan
  // with indexOf instead. Supports variable-length fences (3+ backticks); the
  // closing fence must be at least as long as the opening one (CommonMark). An
  // opening fence with no valid closing fence is left as text.
  let last = 0;
  let searchFrom = 0;
  for (;;) {
    const open = src.indexOf('```', searchFrom);
    if (open === -1) break;
    // Count the full opening backtick run (>=3).
    let openLen = 0;
    while (src[open + openLen] === '`') openLen++;
    const p = open + openLen;
    // Only a short [A-Za-z0-9_+-]* prefix is treated as the language (bounded,
    // no backtracking). Any remaining info-string chars on the line become part
    // of the body — matching the original behavior — so we never scan/trim the
    // rest of the line (which was quadratic on "```!```!…").
    let langEnd = p;
    while (langEnd < src.length && langEnd - p < 32 && /[A-Za-z0-9_+.#-]/.test(src[langEnd])) langEnd++;
    const lang = src.slice(p, langEnd);
    const bodyStart = src[langEnd] === '\n' ? langEnd + 1 : langEnd;
    // Find a closing backtick run of length >= openLen.
    let close = -1;
    let scan = bodyStart;
    for (;;) {
      const cand = src.indexOf('`'.repeat(openLen), scan);
      if (cand === -1) break;
      let runLen = 0;
      while (src[cand + runLen] === '`') runLen++;
      if (runLen >= openLen) { close = cand; break; }
      scan = cand + runLen; // shorter run inside body → keep scanning
    }
    if (close === -1) break; // no valid closing fence → leave the rest as text
    flushText(src.slice(last, open));
    const body = src.slice(bodyStart, close).replace(/\n$/, '');
    parts.push(teamsCodeBlockHtml(body, lang || null));
    // Advance past the full closing run.
    let closeLen = 0;
    while (src[close + closeLen] === '`') closeLen++;
    last = close + closeLen;
    searchFrom = last;
  }
  flushText(src.slice(last));
  return parts.filter(Boolean).join('');
}

const TEAMS_LANGUAGE_ALIASES: Record<string, string> = {
  plain: 'plaintext', 'plain text': 'plaintext',
  js: 'javascript',
  ts: 'typescript', tsx: 'typescript',
  py: 'python',
  yml: 'yaml',
  svg: 'xml',
  sh: 'shell', zsh: 'shell',
  golang: 'go',
  cs: 'csharp', 'c#': 'csharp',
  'c++': 'cpp', cc: 'cpp',
  docker: 'dockerfile',
  bat: 'dos', batch: 'dos', cmd: 'dos',
  gql: 'graphql',
  kt: 'kotlin',
  kusto: 'kql',
  tex: 'latex',
  'objective-c': 'objectivec', objc: 'objectivec',
  matlab: 'octave',
  ps1: 'powershell',
  rb: 'ruby',
  'vb.net': 'vbnet', vb: 'vbnet',
  vbs: 'vbscript',
  md: 'markdown',
  text: 'plaintext', txt: 'plaintext',
};

const TEAMS_LANGUAGE_IDS = new Set([
  'plaintext', 'bash', 'c', 'cpp', 'csharp', 'css', 'dart', 'dockerfile', 'dos', 'go',
  'graphql', 'html', 'http', 'java', 'javascript', 'json', 'jsp', 'jsx', 'kotlin', 'kql',
  'latex', 'lisp', 'markdown', 'objectivec', 'octave', 'perl', 'php', 'powershell',
  'python', 'r', 'ruby', 'rust', 'scala', 'scss', 'shell', 'sql', 'swift', 'typescript',
  'vbnet', 'vbscript', 'verilog', 'vhdl', 'xml', 'yaml',
]);

const TEAMS_HIGHLIGHT_GRAMMAR: Record<string, string> = {
  html: 'xml',
  jsp: 'xml',
  jsx: 'javascript',
  kql: 'sql',
  octave: 'matlab',
};

let codeBlockSequence = 0;

function nextCodeBlockId(): string {
  const uuid = globalThis.crypto?.randomUUID?.();
  if (uuid) return `codeBlockEditor-${uuid}`;
  codeBlockSequence += 1;
  return `codeBlockEditor-${Date.now().toString(36)}-${codeBlockSequence.toString(36)}`;
}

/**
 * Emit the native Teams V2 code-block shape observed from its composer. Teams
 * stores a hidden CodeBlockEditor marker followed by a linked <pre>, and the
 * client submits pre-rendered highlight.js spans rather than a Graph-specific
 * <codeblock> element. This is what permits native languages such as YAML.
 */
export function teamsCodeBlockHtml(code: string, language: string | null): string {
  const requested = language?.trim().toLowerCase() || null;
  const aliased = requested ? (TEAMS_LANGUAGE_ALIASES[requested] ?? requested) : null;
  const teamsLanguage = aliased && TEAMS_LANGUAGE_IDS.has(aliased) ? aliased : null;
  const grammar = teamsLanguage ? (TEAMS_HIGHLIGHT_GRAMMAR[teamsLanguage] ?? teamsLanguage) : null;
  const rendered = highlight(code, grammar);
  // Explicit supported fences retain Teams' exact menu identifier even where
  // highlighting uses a compatible grammar (JSP→XML, KQL→SQL, Octave→MATLAB).
  // Unknown fences safely use auto-detection instead of emitting a class Teams
  // doesn't recognize.
  const detectedTeamsLanguage = rendered.language && TEAMS_LANGUAGE_IDS.has(rendered.language)
    ? rendered.language
    : 'plaintext';
  const effectiveLanguage = teamsLanguage ?? detectedTeamsLanguage;
  // Although a browser normally preserves literal newlines inside <pre>, the
  // Graph/Teams message sanitizer normalizes them as ordinary HTML whitespace.
  // Teams' message body therefore needs explicit <br> elements between lines.
  // parseHtmlBody converts these back to \n when a message is opened for editing.
  const highlightedHtml = rendered.html.replace(/\r?\n/g, '<br>');
  const itemId = nextCodeBlockId();
  return (
    `<p itemtype="http://schema.skype.com/CodeBlockEditor" id="x_${itemId}">&nbsp;</p>` +
    `<pre class="language-${escAttr(effectiveLanguage)} skipProofing" itemid="${itemId}" spellcheck="false">` +
    `<code>${highlightedHtml}</code></pre>`
  );
}

/**
 * Guaranteed-safe fallback for Graph tenants that reject the native Teams code
 * editor metadata. It keeps the block and its line breaks but removes language
 * and highlight classes, avoiding the all-or-nothing PATCH failure.
 */
export function downgradeTeamsCodeBlocks(html: string): string {
  const withoutMarkers = html.replace(
    /<p\b[^>]*\bitemtype="http:\/\/schema\.skype\.com\/CodeBlockEditor"[^>]*>(?:\s|&nbsp;)*<\/p>/gi,
    '',
  );
  return withoutMarkers.replace(
    /<pre\b[^>]*\bitemid="codeBlockEditor-[^"]+"[^>]*>\s*<code>([\s\S]*?)<\/code>\s*<\/pre>/gi,
    (_m, inner: string) => {
      const plainHighlightedHtml = inner.replace(/<\/?span\b[^>]*>/gi, '').replace(/\r?\n/g, '<br>');
      return `<codeblock><code>${plainHighlightedHtml}</code></codeblock>`;
    },
  );
}

export interface PendingImage {
  id: string;
  contentType: string;
  /** base64 (no data: prefix) */
  contentBytes: string;
  name?: string;
}

export interface MessageRef {
  contentType: 'messageReference' | 'forwardedMessageReference';
  id: string;
  messageId: string;
  messagePreview: string | null;
  messageSender: { user: { id?: string | null; displayName: string | null } } | null;
}

/** A file uploaded to a drive, referenced from a chat message as a file card. */
export interface FileReference {
  /** Fresh GUID that ties the <attachment> marker to the attachments[] entry. */
  id: string;
  /** Shareable URL (org-view link) the recipient can open. */
  contentUrl: string;
  name: string;
}

/** Append file `reference` attachments to an outgoing payload (mutates + returns).
 * Each file gets an <attachment id> marker in the body and an attachments[] entry;
 * a text body is upgraded to html so the markers render. */
export function withFileReferences<T extends { body: { contentType: 'text' | 'html'; content: string }; attachments?: unknown[] }>(
  payload: T,
  files: FileReference[],
): T {
  if (files.length === 0) return payload;
  if (payload.body.contentType === 'text') {
    payload.body = { contentType: 'html', content: esc(payload.body.content).replace(/\n/g, '<br>') };
  }
  const markers = files.map((f) => `<attachment id="${escAttr(f.id)}"></attachment>`).join('');
  payload.body.content = payload.body.content + markers;
  payload.attachments ??= [];
  for (const f of files) {
    (payload.attachments as unknown[]).push({ id: f.id, contentType: 'reference', contentUrl: f.contentUrl, name: f.name });
  }
  return payload;
}

/** Attach a reply/forward reference to an outgoing payload (mutates + returns). */
export function withMessageRef<T extends { body: { contentType: 'text' | 'html'; content: string }; attachments?: unknown[] }>(
  payload: T,
  ref: MessageRef,
): T {
  const attachId = ref.id;
  const content = JSON.stringify({
    messageId: ref.messageId,
    messagePreview: ref.messagePreview ?? '',
    messageSender: ref.messageSender ?? undefined,
  });
  // attachId can originate from tool input (e.g. replyToMessageId) — escape it
  // so it cannot break out of the attribute and inject markup.
  const marker = `<attachment id="${escAttr(String(attachId))}"></attachment>`;
  if (payload.body.contentType === 'text') {
    payload.body = { contentType: 'html', content: esc(payload.body.content).replace(/\n/g, '<br>') };
  }
  payload.body.content = marker + payload.body.content;
  (payload.attachments ??= []).push({ id: attachId, contentType: ref.contentType, content });
  return payload;
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
  const PH = MENTION_PH;
  // Strip any user-supplied copies of the private-use placeholder char up front
  // so user text can never forge or collide with our mention placeholders.
  const clean = text.split(PH).join("");
  const withPh = clean.replace(MENTION_TOKEN, (_m, name: string, id: string) => {
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
    return { body: { contentType: 'text', content: clean } };
  }

  let textHtml = html ?? esc(withPh).replace(/\n/g, '<br>');
  // Restore mention placeholders that survive into the rendered body as <at>
  // tags, renumbering them to a dense 0..n-1 sequence. A placeholder that landed
  // in an attribute (href/src/alt/title) was stripped by escAttr/safeUrl and so
  // never reaches here — Graph requires every mentions[] entry to have a matching
  // <at id> AND expects contiguous ids, so we renumber tags + objects together.
  const emittedMentions: OutgoingMention[] = [];
  textHtml = textHtml.replace(new RegExp(`${PH}(\\d+)${PH}`, 'g'), (_m, i) => {
    const mm = mentions[Number(i)];
    // Only restore placeholders we actually created. A fabricated placeholder
    // (e.g. the user typed the private-use char) has no matching mention — drop
    // it rather than dereferencing undefined.
    if (!mm) return '';
    const newId = emittedMentions.length;
    emittedMentions.push({ ...mm, id: newId });
    return `<at id="${newId}">${esc(mm.mentionText)}</at>`;
  });
  // Strip any orphaned placeholder characters that didn't form a full token.
  textHtml = textHtml.split(PH).join('');

  const imgHtml = images
    .map(
      (img) =>
        `<img src="${escAttr(`../hostedContents/${encodeURIComponent(String(img.id))}/$value`)}" style="max-width:100%">`,
    )
    .join('');
  const content = [textHtml, imgHtml].filter(Boolean).join(textHtml && imgHtml ? '<br>' : '');

  return {
    body: { contentType: 'html', content },
    ...(emittedMentions.length ? { mentions: emittedMentions } : {}),
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
