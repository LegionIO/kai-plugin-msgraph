import React, { useCallback, useEffect, useImperativeHandle, useRef, useState, forwardRef } from 'react';
import {
  $createParagraphNode,
  $createTextNode,
  $getRoot,
  $getSelection,
  $insertNodes,
  $isParagraphNode,
  $isRangeSelection,
  COMMAND_PRIORITY_HIGH,
  FORMAT_TEXT_COMMAND,
  PASTE_COMMAND,
  type LexicalEditor,
  type LexicalNode,
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
import { $setBlocksType } from '@lexical/selection';

import { ImageNode, $createImageNodeFromFile, pendingImageFiles } from './nodes.tsx';
import { serializeToMail, type SerializedMail } from './serialize.ts';
import { ensureSyntaxThemeInjected } from './syntax-theme.ts';
import {
  theme,
  MD_TRANSFORMERS,
  CodeHighlightPlugin,
  CodeBlockEscapePlugin,
  FenceTransformPlugin,
  ToolbarBtn,
  hasNonTextContent,
} from './RichComposer.tsx';

export interface MailEditorHandle {
  serialize: () => Promise<SerializedMail>;
  clear: () => void;
  focus: () => void;
}

export const MailEditor = forwardRef<MailEditorHandle, { placeholder?: string; minHeight?: number }>(
  function MailEditor({ placeholder = 'Write your message…', minHeight = 180 }, ref) {
    useEffect(ensureSyntaxThemeInjected, []);
    const initialConfig: InitialConfigType = {
      namespace: 'msgraph-mail-composer',
      theme,
      onError: (e) => console.error('[mail-composer]', e),
      nodes: [ImageNode, CodeNode, CodeHighlightNode, LinkNode, AutoLinkNode, ListNode, ListItemNode, HeadingNode, QuoteNode],
    };
    return (
      <LexicalComposer initialConfig={initialConfig}>
        <MailEditorInner ref={ref} placeholder={placeholder} minHeight={minHeight} />
      </LexicalComposer>
    );
  },
);

const MailEditorInner = forwardRef<MailEditorHandle, { placeholder: string; minHeight: number }>(
  function MailEditorInner({ placeholder, minHeight }, ref) {
    const [editor] = useLexicalComposerContext();
    const [formats, setFormats] = useState<Set<TextFormatType>>(new Set());
    const attachRef = useRef<HTMLInputElement>(null);

    useImperativeHandle(ref, () => ({
      serialize: () => serializeToMail(editor),
      clear: () => {
        editor.update(() => {
          const root = $getRoot();
          root.clear();
          root.append($createParagraphNode());
        });
        pendingImageFiles.clear();
      },
      focus: () => editor.focus(),
    }), [editor]);

    const insertImages = useCallback((files: File[]) => {
      const imgs = files.filter((f) => f.type.startsWith('image/'));
      if (!imgs.length) return;
      editor.update(() => {
        const nodes: LexicalNode[] = imgs.map((f) => $createImageNodeFromFile(f));
        nodes.push($createTextNode(' '));
        $insertNodes(nodes);
      });
    }, [editor]);

    const insertCodeBlock = useCallback(() => {
      editor.update(() => {
        const sel = $getSelection();
        if (!$isRangeSelection(sel)) return;
        const anchorTop = sel.anchor.getNode().getTopLevelElement();
        const focusTop = sel.focus.getNode().getTopLevelElement();
        if (!sel.isCollapsed() && anchorTop && anchorTop.is(focusTop) && $isParagraphNode(anchorTop) && !$isCodeNode(anchorTop)) {
          const text = sel.getTextContent();
          sel.removeText();
          const code = $createCodeNode();
          code.append($createTextNode(text));
          const caret = $getSelection();
          if ($isRangeSelection(caret)) {
            const split = caret.insertParagraph();
            const before = (split ?? anchorTop).getPreviousSibling() ?? anchorTop;
            before.insertAfter(code);
            code.selectEnd();
          }
          return;
        }
        $setBlocksType(sel, () => $createCodeNode());
      });
    }, [editor]);

    useEffect(() => {
      return mergeRegister(
        editor.registerCommand(
          PASTE_COMMAND,
          (event) => {
            const dt = (event as ClipboardEvent | null)?.clipboardData ?? (event as InputEvent | null)?.dataTransfer ?? null;
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
            return true;
          },
          COMMAND_PRIORITY_HIGH,
        ),
        editor.registerCommand(
          DRAG_DROP_PASTE,
          (files) => {
            const imgs = (files as File[]).filter((f) => f.type.startsWith('image/'));
            if (!imgs.length) return false;
            const nodes: LexicalNode[] = imgs.map((f) => $createImageNodeFromFile(f));
            nodes.push($createTextNode(' '));
            $insertNodes(nodes);
            return true;
          },
          COMMAND_PRIORITY_HIGH,
        ),
      );
    }, [editor]);

    const fmtActive = (f: TextFormatType) => formats.has(f);
    const toggleFmt = (f: TextFormatType) => editor.dispatchCommand(FORMAT_TEXT_COMMAND, f);

    return (
      <div className="rounded-lg border border-border bg-background focus-within:border-primary transition-colors">
        <div className="flex items-center gap-0.5 border-b border-border/60 px-1.5 py-1">
          <ToolbarBtn active={fmtActive('bold')} onClick={() => toggleFmt('bold')} title="Bold (⌘B)"><b>B</b></ToolbarBtn>
          <ToolbarBtn active={fmtActive('italic')} onClick={() => toggleFmt('italic')} title="Italic (⌘I)"><i>I</i></ToolbarBtn>
          <ToolbarBtn active={fmtActive('underline')} onClick={() => toggleFmt('underline')} title="Underline (⌘U)"><u>U</u></ToolbarBtn>
          <ToolbarBtn active={fmtActive('strikethrough')} onClick={() => toggleFmt('strikethrough')} title="Strikethrough"><s>S</s></ToolbarBtn>
          <ToolbarBtn active={fmtActive('code')} onClick={() => toggleFmt('code')} title="Inline code"><span className="font-mono">{'</>'}</span></ToolbarBtn>
          <div style={{ width: 1, height: 16 }} className="bg-border mx-1" />
          <ToolbarBtn onClick={insertCodeBlock} title="Code block">
            <svg className="w-3.5 h-3.5" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={2} strokeLinecap="round" strokeLinejoin="round">
              <rect x="3" y="4" width="18" height="16" rx="2" /><path d="m9 10-2 2 2 2M15 10l2 2-2 2" />
            </svg>
          </ToolbarBtn>
          <input
            ref={attachRef}
            type="file"
            accept="image/*"
            multiple
            hidden
            onChange={(e) => {
              editor.focus();
              insertImages(Array.from(e.target.files ?? []));
              e.target.value = '';
            }}
          />
          <ToolbarBtn onClick={() => attachRef.current?.click()} title="Insert image">
            <svg className="w-3.5 h-3.5" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={2} strokeLinecap="round" strokeLinejoin="round">
              <rect x="3" y="3" width="18" height="18" rx="2" /><circle cx="9" cy="9" r="2" /><path d="m21 15-5-5L5 21" />
            </svg>
          </ToolbarBtn>
        </div>
        <div className="relative">
          <RichTextPlugin
            contentEditable={
              <ContentEditable
                className="text-sm text-foreground outline-none px-3 py-2.5"
                style={{ minHeight, maxHeight: 380, overflowY: 'auto' }}
                aria-placeholder={placeholder}
                placeholder={
                  <div style={{ position: 'absolute', top: 10, left: 12, pointerEvents: 'none' }} className="text-sm text-muted-foreground">
                    {placeholder}
                  </div>
                }
              />
            }
            ErrorBoundary={LexicalErrorBoundary}
          />
          <HistoryPlugin />
          <MarkdownShortcutPlugin transformers={MD_TRANSFORMERS} />
          <OnChangePlugin
            onChange={(state) => {
              state.read(() => {
                void hasNonTextContent;
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
          <CodeHighlightPlugin />
          <CodeBlockEscapePlugin />
          <FenceTransformPlugin />
        </div>
      </div>
    );
  },
);

export function editorFormatState(_ed: LexicalEditor) { void _ed; }
