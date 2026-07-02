import {
  $getRoot,
  $isElementNode,
  $isLineBreakNode,
  $isParagraphNode,
  $isTextNode,
  type LexicalEditor,
  type LexicalNode,
} from 'lexical';
import { $isCodeNode } from '@lexical/code';
import { $isLinkNode } from '@lexical/link';
import { $isListNode, $isListItemNode } from '@lexical/list';
import { $isMentionNode, $isImageNode, pendingImageFiles } from './nodes.tsx';
import type { OutgoingMention } from '../../shared/markdown.ts';
import { fileToBase64 } from '../components/NewChatDialog.tsx';

const esc = (s: string) =>
  s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

interface Ctx {
  mentions: OutgoingMention[];
  images: Array<{ tempId: string; contentType: string }>;
}

function textNodeHtml(node: LexicalNode): string {
  if (!$isTextNode(node)) return '';
  let out = esc(node.getTextContent());
  if (node.hasFormat('code')) return `<code>${out}</code>`;
  if (node.hasFormat('bold')) out = `<strong>${out}</strong>`;
  if (node.hasFormat('italic')) out = `<em>${out}</em>`;
  if (node.hasFormat('strikethrough')) out = `<s>${out}</s>`;
  if (node.hasFormat('underline')) out = `<u>${out}</u>`;
  return out;
}

function inlineHtml(node: LexicalNode, ctx: Ctx): string {
  if ($isMentionNode(node)) {
    const idx = ctx.mentions.length;
    ctx.mentions.push({
      id: idx,
      mentionText: node.getDisplayName(),
      mentioned: { user: { id: node.getUserId(), displayName: node.getDisplayName(), userIdentityType: 'aadUser' } },
    });
    return `<at id="${idx}">${esc(node.getDisplayName())}</at>`;
  }
  if ($isTextNode(node)) return textNodeHtml(node);
  if ($isLineBreakNode(node)) return '<br>';
  if ($isImageNode(node)) {
    const existing = node.getExistingUrl();
    if (existing) {
      return `<img src="${esc(existing)}" style="max-width:100%">`;
    }
    const tempId = node.getTempId();
    ctx.images.push({ tempId, contentType: node.getContentType() });
    return `<img src="../hostedContents/${tempId}/$value" style="max-width:100%">`;
  }
  if ($isLinkNode(node)) {
    const inner = node.getChildren().map((c) => inlineHtml(c, ctx)).join('');
    return `<a href="${esc(node.getURL())}">${inner}</a>`;
  }
  if ($isElementNode(node)) {
    return node.getChildren().map((c) => inlineHtml(c, ctx)).join('');
  }
  return '';
}

function blockHtml(node: LexicalNode, ctx: Ctx): string {
  if ($isCodeNode(node)) {
    const lang = node.getLanguage();
    const cls = lang ? ` class="language-${esc(lang)}"` : '';
    const text = node.getTextContent();
    return `<codeblock${cls}><code>${esc(text).replace(/\n/g, '<br>')}</code></codeblock>`;
  }
  if ($isListNode(node)) {
    const tag = node.getListType() === 'number' ? 'ol' : 'ul';
    const items = node
      .getChildren()
      .map((li) => ($isListItemNode(li) ? `<li>${li.getChildren().map((c) => inlineHtml(c, ctx)).join('')}</li>` : ''))
      .join('');
    return `<${tag}>${items}</${tag}>`;
  }
  if ($isParagraphNode(node) || $isElementNode(node)) {
    const inner = ($isElementNode(node) ? node.getChildren() : []).map((c) => inlineHtml(c, ctx)).join('');
    return `<p>${inner || '&nbsp;'}</p>`;
  }
  return '';
}

export interface SerializedPayload {
  body: { contentType: 'text' | 'html'; content: string };
  mentions?: OutgoingMention[];
  hostedContents?: Array<{ '@microsoft.graph.temporaryId': string; contentBytes: string; contentType: string }>;
}

export async function serializeToTeams(
  editor: LexicalEditor,
): Promise<{ payload: SerializedPayload; isEmpty: boolean }> {
  const { html, plain, ctx, simple } = editor.getEditorState().read(() => {
    const c: Ctx = { mentions: [], images: [] };
    const root = $getRoot();
    const blocks = root.getChildren();
    const plainText = root.getTextContent();
    const isSimple = blocks.length <= 1 && ctxProbeSimple(blocks[0]);
    const h = blocks.map((b) => blockHtml(b, c)).join('');
    return { html: h, plain: plainText, ctx: c, simple: isSimple };
  });

  if (simple && plain.trim()) {
    return { payload: { body: { contentType: 'text', content: plain } }, isEmpty: false };
  }

  const isEmpty = !plain.trim() && ctx.images.length === 0;
  const payload: SerializedPayload = { body: { contentType: 'html', content: html } };
  if (ctx.mentions.length) payload.mentions = ctx.mentions;
  if (ctx.images.length) {
    payload.hostedContents = await Promise.all(
      ctx.images.map(async (img) => {
        const file = pendingImageFiles.get(img.tempId);
        const contentBytes = file ? await fileToBase64(file) : '';
        return {
          '@microsoft.graph.temporaryId': img.tempId,
          contentBytes,
          contentType: img.contentType,
        };
      }),
    );
  }
  return { payload, isEmpty };
}

function ctxProbeSimple(node: LexicalNode | undefined): boolean {
  if (!node || !$isParagraphNode(node)) return false;
  for (const c of node.getChildren()) {
    if ($isLineBreakNode(c)) continue;
    if ($isMentionNode(c)) return false;
    if (!$isTextNode(c)) return false;
    if (c.getFormat() !== 0) return false;
  }
  return true;
}
