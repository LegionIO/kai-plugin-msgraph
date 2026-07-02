import React, { useCallback, useEffect, useRef, useState } from 'react';
import {
  $createParagraphNode,
  $createTextNode,
  $getRoot,
  $getSelection,
  $isParagraphNode,
  $isRangeSelection,
  COMMAND_PRIORITY_LOW,
  COMMAND_PRIORITY_HIGH,
  FORMAT_TEXT_COMMAND,
  $insertNodes,
  $isLineBreakNode,
  $isTextNode,
  KEY_ARROW_DOWN_COMMAND,
  KEY_ARROW_UP_COMMAND,
  KEY_ENTER_COMMAND,
  KEY_ESCAPE_COMMAND,
  PASTE_COMMAND,
  TextNode,
  LineBreakNode,
  type LexicalEditor,
  type TextFormatType,
} from 'lexical';
import { LexicalComposer, type InitialConfigType } from '@lexical/react/LexicalComposer';
import { RichTextPlugin } from '@lexical/react/LexicalRichTextPlugin';
import { ContentEditable } from '@lexical/react/LexicalContentEditable';
import { LexicalErrorBoundary } from '@lexical/react/LexicalErrorBoundary';
import { HistoryPlugin } from '@lexical/react/LexicalHistoryPlugin';
import { OnChangePlugin } from '@lexical/react/LexicalOnChangePlugin';
import { MarkdownShortcutPlugin } from '@lexical/react/LexicalMarkdownShortcutPlugin';
import { useLexicalComposerContext } from '@lexical/react/LexicalComposerContext';
import { mergeRegister } from '@lexical/utils';
import { $createCodeNode, $isCodeNode, CodeNode, CodeHighlightNode } from '@lexical/code';
import { LinkNode, AutoLinkNode } from '@lexical/link';
import { ListNode, ListItemNode } from '@lexical/list';
import { HeadingNode, QuoteNode, DRAG_DROP_PASTE } from '@lexical/rich-text';
import {
  BOLD_ITALIC_STAR,
  BOLD_STAR,
  INLINE_CODE,
  ITALIC_STAR,
  ITALIC_UNDERSCORE,
  STRIKETHROUGH,
} from '@lexical/markdown';
import { $setBlocksType } from '@lexical/selection';

import { MentionNode, ImageNode, $createMentionNode, $createImageNodeFromFile, pendingImageFiles } from './nodes.tsx';
import { serializeToTeams, type SerializedPayload } from './serialize.ts';
import { Avatar } from '../components/Avatar.tsx';
import type { MsgraphPluginState, Presence } from '../../shared/types.ts';

// ── Theme (class names Kai's Tailwind is likely to have; layout via inline style) ──

const theme = {
  paragraph: 'm-0',
  text: {
    bold: 'font-semibold',
    italic: 'italic',
    strikethrough: 'line-through',
    underline: 'underline',
    code: 'font-mono text-xs px-1 py-0.5 rounded bg-background/60 border border-border/60',
  },
  code: 'block font-mono text-xs rounded-lg border border-border bg-card px-2.5 py-2 my-1 whitespace-pre-wrap',
};

const MD_TRANSFORMERS = [
  BOLD_ITALIC_STAR,
  BOLD_STAR,
  ITALIC_STAR,
  ITALIC_UNDERSCORE,
  STRIKETHROUGH,
  INLINE_CODE,
];

export interface RichComposerProps {
  chatId: string;
  sending: boolean;
  onSend: (payload: SerializedPayload) => void;
  onSearchPeople: (query: string) => void;
  peopleSearch: MsgraphPluginState['peopleSearch'];
  photos: Record<string, string | null>;
  presence: Record<string, Presence>;
}

export function RichComposer(props: RichComposerProps) {
  const initialConfig: InitialConfigType = {
    namespace: 'msgraph-composer',
    theme,
    onError: (e) => console.error('[composer]', e),
    nodes: [
      MentionNode,
      ImageNode,
      CodeNode,
      CodeHighlightNode,
      LinkNode,
      AutoLinkNode,
      ListNode,
      ListItemNode,
      HeadingNode,
      QuoteNode,
    ],
  };

  return (
    <LexicalComposer initialConfig={initialConfig}>
      <ComposerInner {...props} />
    </LexicalComposer>
  );
}

function ComposerInner({
  chatId,
  sending,
  onSend,
  onSearchPeople,
  peopleSearch,
  photos,
  presence,
}: RichComposerProps) {
  const [editor] = useLexicalComposerContext();
  const [isEmpty, setIsEmpty] = useState(true);
  const [formats, setFormats] = useState<Set<TextFormatType>>(new Set());
  const attachRef = useRef<HTMLInputElement>(null);

  // Reset when switching chats.
  useEffect(() => {
    editor.update(() => {
      const root = $getRoot();
      root.clear();
      root.append($createParagraphNode());
    });
  }, [chatId, editor]);

  const onSendRef = useRef(onSend);
  useEffect(() => { onSendRef.current = onSend; }, [onSend]);

  const send = useCallback(() => {
    void (async () => {
      const { payload, isEmpty: empty } = await serializeToTeams(editor);
      if (empty) return;
      onSendRef.current(payload);
      editor.update(() => {
        const root = $getRoot();
        root.clear();
        root.append($createParagraphNode());
      });
      pendingImageFiles.clear();
    })();
  }, [editor]);

  const insertImages = useCallback(
    (files: File[]) => {
      const imgs = files.filter((f) => f.type.startsWith('image/'));
      if (!imgs.length) return;
      editor.update(() => {
        const nodes: LexicalNode[] = imgs.map((f) => $createImageNodeFromFile(f));
        nodes.push($createTextNode(' '));
        $insertNodes(nodes);
      });
    },
    [editor],
  );

  const insertCodeBlock = useCallback(() => {
    editor.update(() => {
      const sel = $getSelection();
      if (!$isRangeSelection(sel)) return;
      const anchorTop = sel.anchor.getNode().getTopLevelElement();
      const focusTop = sel.focus.getNode().getTopLevelElement();
      // Range within a single non-code paragraph → split and wrap just the selection.
      if (
        !sel.isCollapsed() &&
        anchorTop &&
        anchorTop.is(focusTop) &&
        $isParagraphNode(anchorTop) &&
        !$isCodeNode(anchorTop)
      ) {
        const text = sel.getTextContent();
        sel.removeText();
        const code = $createCodeNode();
        code.append($createTextNode(text));
        // Split the paragraph at the (now-collapsed) caret and insert the code block between.
        const caret = $getSelection();
        if ($isRangeSelection(caret)) {
          const split = caret.insertParagraph();
          const before = (split ?? anchorTop).getPreviousSibling() ?? anchorTop;
          before.insertAfter(code);
          code.selectEnd();
        }
        return;
      }
      // Collapsed, multi-block, or already a code node → convert touched blocks.
      $setBlocksType(sel, () => $createCodeNode());
    });
  }, [editor]);

  // Enter-to-send (Shift+Enter = newline). When mention popover is open, Enter selects.
  const mentionOpenRef = useRef(false);
  useEffect(() => {
    return mergeRegister(
      editor.registerCommand(
        KEY_ENTER_COMMAND,
        (event) => {
          if (mentionOpenRef.current) return false;
          if (event?.shiftKey) return false;
          // Inside a code block: let Enter insert a newline.
          const sel = $getSelection();
          if ($isRangeSelection(sel)) {
            const top = sel.anchor.getNode().getTopLevelElement();
            if (top && $isCodeNode(top)) return false;
            // Line is exactly ``` or ```lang → convert to code block instead of sending.
            // (Fence conversion is handled by FenceTransformPlugin on the closing fence.)
          }
          event?.preventDefault();
          send();
          return true;
        },
        COMMAND_PRIORITY_LOW,
      ),
      editor.registerCommand(
        PASTE_COMMAND,
        (event) => {
          const dt =
            (event as ClipboardEvent | null)?.clipboardData ??
            (event as InputEvent | null)?.dataTransfer ??
            null;
          const types = Array.from(dt?.types ?? []);
          console.log('[msgraph composer] PASTE_COMMAND', { evt: event?.constructor?.name, types, files: dt?.files?.length, items: dt?.items?.length });
          const files: File[] = [];
          for (const f of Array.from(dt?.files ?? [])) if (f.type.startsWith('image/')) files.push(f);
          if (!files.length) {
            for (const it of Array.from(dt?.items ?? [])) {
              if (it.kind === 'file' && it.type.startsWith('image/')) {
                const f = it.getAsFile();
                if (f) files.push(f);
              }
            }
          }
          if (!files.length) return false;
          (event as Event | null)?.preventDefault?.();
          const nodes: LexicalNode[] = files.map((f) => $createImageNodeFromFile(f));
          nodes.push($createTextNode(' '));
          $insertNodes(nodes);
          console.log('[msgraph composer] inserted', files.length, 'image(s) via PASTE_COMMAND');
          return true;
        },
        COMMAND_PRIORITY_HIGH,
      ),
      editor.registerCommand(
        DRAG_DROP_PASTE,
        (files) => {
          const imgs = (files as File[]).filter((f) => f.type.startsWith('image/'));
          console.log('[msgraph composer] DRAG_DROP_PASTE', { total: (files as File[]).length, images: imgs.length });
          if (!imgs.length) return false;
          const nodes: LexicalNode[] = imgs.map((f) => $createImageNodeFromFile(f));
          nodes.push($createTextNode(' '));
          $insertNodes(nodes);
          return true;
        },
        COMMAND_PRIORITY_HIGH,
      ),
    );
  }, [editor, send, insertImages]);

  const fmtActive = (f: TextFormatType) => formats.has(f);
  const toggleFmt = (f: TextFormatType) => editor.dispatchCommand(FORMAT_TEXT_COMMAND, f);

  return (
    <div className="border-t border-border/50 p-3 shrink-0">
      <div className="rounded-lg border border-border bg-muted focus-within:border-primary transition-colors">
        <div className="relative">
          <RichTextPlugin
            contentEditable={
              <ContentEditable
                className="text-sm text-foreground outline-none px-3 py-2"
                style={{ minHeight: 36, maxHeight: 200, overflowY: 'auto' }}
                aria-placeholder="Message…"
                placeholder={
                  <div
                    style={{ position: 'absolute', top: 8, left: 12, pointerEvents: 'none' }}
                    className="text-sm text-muted-foreground"
                  >
                    Message…
                  </div>
                }
              />
            }
            ErrorBoundary={LexicalErrorBoundary}
          />
          <HistoryPlugin />
          <MarkdownShortcutPlugin transformers={MD_TRANSFORMERS} />
          <OnChangePlugin
            onChange={(state, ed) => {
              state.read(() => {
                setIsEmpty(!$getRoot().getTextContent().trim() && !hasNonTextContent(ed));
                const sel = $getSelection();
                if ($isRangeSelection(sel)) {
                  const next = new Set<TextFormatType>();
                  (['bold', 'italic', 'strikethrough', 'underline', 'code'] as TextFormatType[]).forEach((f) => {
                    if (sel.hasFormat(f)) next.add(f);
                  });
                  setFormats(next);
                }
              });
            }}
          />
          <CodeBlockEscapePlugin />
          <FenceTransformPlugin />
          <MentionTypeahead
            editor={editor}
            onSearch={onSearchPeople}
            peopleSearch={peopleSearch}
            photos={photos}
            presence={presence}
            openRef={mentionOpenRef}
          />
        </div>
        <div className="flex items-center justify-between gap-2 border-t border-border/60 px-1.5 py-1">
          <div className="flex items-center gap-0.5">
            <ToolbarBtn active={fmtActive('bold')} onClick={() => toggleFmt('bold')} title="Bold (⌘B)">
              <b>B</b>
            </ToolbarBtn>
            <ToolbarBtn active={fmtActive('italic')} onClick={() => toggleFmt('italic')} title="Italic (⌘I)">
              <i>I</i>
            </ToolbarBtn>
            <ToolbarBtn active={fmtActive('underline')} onClick={() => toggleFmt('underline')} title="Underline (⌘U)">
              <u>U</u>
            </ToolbarBtn>
            <ToolbarBtn active={fmtActive('strikethrough')} onClick={() => toggleFmt('strikethrough')} title="Strikethrough">
              <s>S</s>
            </ToolbarBtn>
            <ToolbarBtn active={fmtActive('code')} onClick={() => toggleFmt('code')} title="Inline code">
              <span className="font-mono">{'</>'}</span>
            </ToolbarBtn>
            <div style={{ width: 1, height: 16 }} className="bg-border mx-1" />
            <ToolbarBtn onClick={insertCodeBlock} title="Code block">
              <svg className="w-3.5 h-3.5" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={2} strokeLinecap="round" strokeLinejoin="round">
                <rect x="3" y="4" width="18" height="16" rx="2" /><path d="m9 10-2 2 2 2M15 10l2 2-2 2" />
              </svg>
            </ToolbarBtn>
            <ToolbarBtn
              onClick={() =>
                editor.update(() => {
                  const sel = $getSelection();
                  if ($isRangeSelection(sel)) sel.insertText('@');
                })
              }
              title="Mention someone"
            >
              @
            </ToolbarBtn>
            <input
              ref={attachRef}
              type="file"
              accept="image/*"
              multiple
              hidden
              onChange={(e) => {
                editor.focus();
                const fs = Array.from(e.target.files ?? []);
                insertImages(fs);
                e.target.value = '';
              }}
            />
            <ToolbarBtn onClick={() => attachRef.current?.click()} title="Insert image">
              <svg className="w-3.5 h-3.5" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={2} strokeLinecap="round" strokeLinejoin="round">
                <rect x="3" y="3" width="18" height="18" rx="2" /><circle cx="9" cy="9" r="2" /><path d="m21 15-5-5L5 21" />
              </svg>
            </ToolbarBtn>
          </div>
          <button
            type="button"
            disabled={isEmpty || sending}
            onClick={send}
            className="px-3 py-1.5 text-xs font-medium bg-primary text-primary-foreground rounded-lg hover:bg-primary/90 disabled:opacity-50 transition-colors"
          >
            {sending ? '…' : 'Send'}
          </button>
        </div>
      </div>
    </div>
  );
}

function ToolbarBtn({
  active,
  onClick,
  title,
  children,
}: {
  active?: boolean;
  onClick: () => void;
  title: string;
  children: React.ReactNode;
}) {
  return (
    <button
      type="button"
      title={title}
      onMouseDown={(e) => e.preventDefault()}
      onClick={onClick}
      style={{ width: 26, height: 26 }}
      className={`flex items-center justify-center rounded text-xs transition-colors ${
        active ? 'bg-primary/20 text-primary' : 'text-muted-foreground hover:text-foreground hover:bg-muted'
      }`}
    >
      {children}
    </button>
  );
}

function hasNonTextContent(editor: LexicalEditor): boolean {
  return editor.getEditorState().read(() => {
    let found = false;
    $getRoot().getChildren().forEach((b) => {
      if (found) return;
      if (!('getChildren' in b)) return;
      (b as { getChildren: () => LexicalNode[] }).getChildren().forEach((c: LexicalNode) => {
        if (c.getType() === 'mention' || c.getType() === 'inline-image') found = true;
      });
    });
    return found;
  });
}

import type { LexicalNode } from 'lexical';

// ── Code-block boundary escape (arrow past first/last line inserts a sibling paragraph) ──

function CodeBlockEscapePlugin() {
  const [editor] = useLexicalComposerContext();
  useEffect(() => {
    const escape = (dir: 'up' | 'down') => (ev: KeyboardEvent | null | undefined) => {
      const sel = $getSelection();
      if (!$isRangeSelection(sel) || !sel.isCollapsed()) return false;
      const anchor = sel.anchor.getNode();
      const block = anchor.getTopLevelElement();
      if (!block || !$isCodeNode(block)) return false;
      const atStart = dir === 'up' && block.getFirstDescendant()?.getKey() === anchor.getKey() && sel.anchor.offset === 0;
      const last = block.getLastDescendant();
      const atEnd =
        dir === 'down' &&
        last?.getKey() === anchor.getKey() &&
        sel.anchor.offset === anchor.getTextContentSize();
      if (!atStart && !atEnd) return false;
      const sib = dir === 'down' ? block.getNextSibling() : block.getPreviousSibling();
      ev?.preventDefault();
      if (sib) {
        (dir === 'down' ? sib.selectStart() : sib.selectEnd());
      } else {
        const p = $createParagraphNode();
        if (dir === 'down') block.insertAfter(p);
        else block.insertBefore(p);
        p.select();
      }
      return true;
    };
    return mergeRegister(
      editor.registerCommand(KEY_ARROW_DOWN_COMMAND, escape('down'), COMMAND_PRIORITY_LOW),
      editor.registerCommand(KEY_ARROW_UP_COMMAND, escape('up'), COMMAND_PRIORITY_LOW),
    );
  }, [editor]);
  return null;
}

// ── ```-fence auto-convert. Line-aware: works whether the fence is its own
//    paragraph or a line inside a Shift+Enter'd paragraph. Triggers when the
//    fence line is followed by Space or a LineBreakNode. ──

const OPEN_FENCE_RE = /^(`{3,})([\w+-]*)\s*$/;
const CLOSE_FENCE_RE = /^(`{3,})\s*$/;

function fenceInfo(n: LexicalNode | null): { ticks: number; lang: string | null } | null {
  if (!n || !$isTextNode(n) || n.getType() !== 'text') return null;
  const m = OPEN_FENCE_RE.exec(n.getTextContent());
  if (!m) return null;
  const prev = n.getPreviousSibling();
  if (prev && !$isLineBreakNode(prev)) return null;
  return { ticks: m[1].length, lang: m[2] || null };
}

/** Convert only when `node` is a closing fence with a matching opener earlier in the same paragraph. */
function convertFenceLine(node: TextNode): boolean {
  const close = CLOSE_FENCE_RE.exec(node.getTextContent());
  if (!close) return false;
  const ticks = close[1].length;

  const prev = node.getPreviousSibling();
  const next = node.getNextSibling();
  if (prev && !$isLineBreakNode(prev)) return false;
  const terminated = /\s$/.test(node.getTextContent()) || next === null || $isLineBreakNode(next);
  if (!terminated) return false;

  const para = node.getParent();
  if (!para || !$isParagraphNode(para)) return false;

  // Walk back to find an opener with the SAME tick count.
  let opener: TextNode | null = null;
  let openerLang: string | undefined;
  const between: LexicalNode[] = [];
  for (let s: LexicalNode | null = prev; s; s = s.getPreviousSibling()) {
    const fi = fenceInfo(s);
    if (fi && fi.ticks === ticks) {
      opener = s as TextNode;
      openerLang = fi.lang ?? undefined;
      break;
    }
    between.unshift(s);
  }
  if (!opener) return false;

  const code = $createCodeNode(openerLang);
  const content = between
    .map((n) => ($isLineBreakNode(n) ? '\n' : n.getTextContent()))
    .join('')
    .replace(/^\n+|\n+$/g, '');
  if (content) code.append($createTextNode(content));

  const tailStart = $isLineBreakNode(next) ? next.getNextSibling() : next;
  const tail: LexicalNode[] = [];
  for (let s = tailStart; s; s = s.getNextSibling()) tail.push(s);

  const openerPrev = opener.getPreviousSibling();
  if ($isLineBreakNode(openerPrev)) openerPrev.remove();
  opener.remove();
  for (const n of between) n.remove();
  if ($isLineBreakNode(next)) next.remove();
  node.remove();

  if (para.getChildrenSize() === 0) para.replace(code);
  else para.insertAfter(code);
  if (tail.length) {
    const afterPara = $createParagraphNode();
    for (const t of tail) afterPara.append(t);
    code.insertAfter(afterPara);
  }
  code.selectEnd();
  return true;
}

function FenceTransformPlugin() {
  const [editor] = useLexicalComposerContext();
  useEffect(() => {
    return mergeRegister(
      // Space (or any trailing whitespace) after ```lang on its own line.
      editor.registerNodeTransform(TextNode, (n) => {
        if (n.getType() !== 'text') return; // skip MentionNode etc.
        convertFenceLine(n);
      }),
      // Shift+Enter after ```lang: the LineBreakNode appears; check the text node before it.
      editor.registerNodeTransform(LineBreakNode, (br) => {
        const p = br.getPreviousSibling();
        if ($isTextNode(p) && p.getType() === 'text') convertFenceLine(p);
      }),
    );
  }, [editor]);
  return null;
}

// ── @-mention typeahead (portal-free) ──

function MentionTypeahead({
  editor,
  onSearch,
  peopleSearch,
  photos,
  presence,
  openRef,
}: {
  editor: LexicalEditor;
  onSearch: (q: string) => void;
  peopleSearch: MsgraphPluginState['peopleSearch'];
  photos: Record<string, string | null>;
  presence: Record<string, Presence>;
  openRef: React.MutableRefObject<boolean>;
}) {
  const [match, setMatch] = useState<{ query: string } | null>(null);
  const [highlight, setHighlight] = useState(0);
  const abandonedPrefix = useRef<string | null>(null);
  const debounce = useRef<ReturnType<typeof setTimeout> | null>(null);

  const onSearchRef = useRef(onSearch);
  useEffect(() => { onSearchRef.current = onSearch; }, [onSearch]);

  useEffect(() => {
    return editor.registerUpdateListener(({ editorState }) => {
      editorState.read(() => {
        const sel = $getSelection();
        if (!$isRangeSelection(sel) || !sel.isCollapsed()) {
          setMatch((cur) => (cur ? null : cur));
          return;
        }
        const node = sel.anchor.getNode();
        const offset = sel.anchor.offset;
        const text = node.getTextContent().slice(0, offset);
        // Allow spaces/commas so "@Doe, Jane A" works. Anchor at the nearest @ not preceded by a word char.
        const m = /(?:^|[^\p{L}\p{N}@])@([^\n@]{0,60})$/u.exec(text);
        let q = m ? m[1].replace(/\s+$/, '') : null;
        if (q !== null) {
          const ap = abandonedPrefix.current;
          if (ap) {
            const ql = q.toLowerCase();
            if (ql.length > ap.length && ql.startsWith(ap)) {
              q = null; // extended past a zero-result prefix — stay closed
            } else if (ql === ap) {
              q = null; // exactly the dead query — stay closed (no flicker)
            } else {
              abandonedPrefix.current = null; // cursor moved/edited to a different (or shorter) prefix — resume
            }
          }
        } else {
          abandonedPrefix.current = null;
        }
        setMatch((cur) => {
          if (q === null) return cur === null ? cur : null;
          return cur?.query === q ? cur : { query: q };
        });
      });
    });
  }, [editor]);

  useEffect(() => {
    openRef.current = !!match;
    if (debounce.current) clearTimeout(debounce.current);
    if (!match) return;
    debounce.current = setTimeout(() => onSearchRef.current(match.query), 180);
    return () => {
      if (debounce.current) clearTimeout(debounce.current);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [match?.query]);

  const resultsCurrent = !!match && peopleSearch?.query === match.query;
  const suggestions = resultsCurrent ? peopleSearch?.results ?? [] : [];
  const searching = !!match && (!resultsCurrent || peopleSearch?.loading);

  const suggestionsRef = useRef(suggestions);
  suggestionsRef.current = suggestions;
  const highlightRef = useRef(highlight);
  highlightRef.current = highlight;
  useEffect(() => setHighlight(0), [match?.query, suggestions.length]);

  // When a query with whitespace comes back empty, remember it so further typing doesn't keep searching.
  useEffect(() => {
    if (match && resultsCurrent && !peopleSearch?.loading && suggestions.length === 0 && /\s/.test(match.query)) {
      abandonedPrefix.current = match.query.toLowerCase();
      setMatch(null);
    }
  }, [match, resultsCurrent, peopleSearch?.loading, suggestions.length]);

  const pick = (p: { id: string; displayName: string }) => {
    editor.update(() => {
      const sel = $getSelection();
      if (!$isRangeSelection(sel) || !sel.isCollapsed()) return;
      const node = sel.anchor.getNode();
      const offset = sel.anchor.offset;
      const text = node.getTextContent();
      const before = text.slice(0, offset);
      const m = /(?:^|[^\p{L}\p{N}@])@([^\n@]{0,60})$/u.exec(before);
      if (!m) return;
      const delStart = offset - m[1].length - 1;
      sel.anchor.set(node.getKey(), delStart, 'text');
      sel.focus.set(node.getKey(), offset, 'text');
      sel.removeText();
      sel.insertNodes([$createMentionNode(p.id, p.displayName), $createTextNode(' ')]);
    });
    setMatch(null);
  };

  useEffect(() => {
    if (!match) return;
    const move = (delta: number) => (ev: KeyboardEvent | null | undefined) => {
      const n = suggestionsRef.current.length;
      if (n === 0) return false;
      ev?.preventDefault();
      setHighlight((h) => (((h + delta) % n) + n) % n);
      return true;
    };
    return mergeRegister(
      editor.registerCommand(KEY_ARROW_DOWN_COMMAND, move(1), COMMAND_PRIORITY_HIGH),
      editor.registerCommand(KEY_ARROW_UP_COMMAND, move(-1), COMMAND_PRIORITY_HIGH),
      editor.registerCommand(
        KEY_ENTER_COMMAND,
        (event) => {
          const list = suggestionsRef.current;
          const idx = Math.min(highlightRef.current, list.length - 1);
          if (list[idx]) {
            event?.preventDefault();
            pick(list[idx]);
            return true;
          }
          return false;
        },
        COMMAND_PRIORITY_HIGH,
      ),
      editor.registerCommand(
        KEY_ESCAPE_COMMAND,
        () => {
          setMatch(null);
          return true;
        },
        COMMAND_PRIORITY_HIGH,
      ),
    );
  }, [editor, match]);

  if (!match) return null;
  return (
    <div
      style={{ position: 'absolute', bottom: 'calc(100% + 4px)', left: 8, width: 260, maxHeight: 220 }}
      className="z-30 overflow-y-auto rounded-lg border border-border bg-card shadow-xl"
    >
      {searching && suggestions.length === 0 && (
        <div className="px-3 py-2 text-[11px] text-muted-foreground animate-pulse">Searching…</div>
      )}
      {!searching && suggestions.length === 0 && (
        <div className="px-3 py-2 text-[11px] text-muted-foreground">
          {match.query ? 'No matches' : 'Type a name…'}
        </div>
      )}
      {suggestions.map((p, i) => (
        <button
          key={p.id}
          type="button"
          ref={i === highlight ? (el) => el?.scrollIntoView({ block: 'nearest' }) : undefined}
          onMouseEnter={() => setHighlight(i)}
          onMouseDown={(e) => {
            e.preventDefault();
            pick(p);
          }}
          className={`w-full flex items-center gap-2 px-2 py-1.5 text-left ${
            i === highlight ? 'bg-primary/15' : 'hover:bg-muted/60'
          }`}
        >
          <Avatar id={p.id} name={p.displayName} photo={photos[p.id]} presence={presence[p.id]} size={7} />
          <div className="min-w-0">
            <div className="text-xs text-foreground truncate">{p.displayName}</div>
            <div className="text-[10px] text-muted-foreground truncate">{p.email}</div>
          </div>
        </button>
      ))}
    </div>
  );
}
