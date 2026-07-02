import {
  $createLineBreakNode,
  $createParagraphNode,
  $createTextNode,
  type LexicalNode,
  type ParagraphNode,
} from 'lexical';
import { $createCodeNode } from '@lexical/code';
import { $createLinkNode } from '@lexical/link';
import { $createMentionNode, $createImageNodeFromExisting } from './nodes.tsx';
import type { BodySegment } from '../../shared/types.ts';

/**
 * Build top-level Lexical block nodes from parsed BodySegments so an existing
 * message can be loaded into the composer for editing. Images use the
 * hostedContents cache for display but keep their original Graph URL for
 * round-trip serialization.
 */
export function $segmentsToBlocks(
  segments: BodySegment[],
  hostedContents: Record<string, string | null>,
): LexicalNode[] {
  const blocks: LexicalNode[] = [];
  let para: ParagraphNode = $createParagraphNode();
  blocks.push(para);

  const pushInline = (n: LexicalNode) => para.append(n);
  const newPara = () => {
    para = $createParagraphNode();
    blocks.push(para);
  };

  for (const seg of segments) {
    switch (seg.type) {
      case 'text':
        pushInline($createTextNode(seg.text));
        break;
      case 'br':
        pushInline($createLineBreakNode());
        break;
      case 'mention':
        if (seg.userId) pushInline($createMentionNode(seg.userId, seg.displayName));
        else pushInline($createTextNode(seg.displayName));
        break;
      case 'code': {
        const t = $createTextNode(seg.code);
        t.setFormat('code');
        pushInline(t);
        break;
      }
      case 'link': {
        const link = $createLinkNode(seg.href);
        link.append($createTextNode(seg.text || seg.href));
        pushInline(link);
        break;
      }
      case 'image': {
        const display = hostedContents[seg.url] ?? seg.url;
        pushInline($createImageNodeFromExisting(seg.url, display));
        break;
      }
      case 'hr':
        newPara();
        break;
      case 'heading': {
        // Headings collapse to a bold line in the composer (Teams headings are rare in chat).
        newPara();
        const t = $createTextNode(seg.segments.map(flatText).join(''));
        t.setFormat('bold');
        pushInline(t);
        newPara();
        break;
      }
      case 'blockquote': {
        newPara();
        for (const inner of $segmentsToBlocks(seg.segments, hostedContents)) blocks.push(inner);
        para = blocks[blocks.length - 1] as ParagraphNode;
        newPara();
        break;
      }
      case 'codeblock': {
        const code = $createCodeNode(seg.lang ?? undefined);
        code.append($createTextNode(seg.code));
        blocks.push(code);
        newPara();
        break;
      }
      case 'table': {
        // No table node in the composer — flatten to tab-separated lines so content isn't lost.
        newPara();
        const t = $createTextNode(
          [
            ...(seg.header ? [seg.header.join('\t')] : []),
            ...seg.rows.map((r) => r.join('\t')),
          ].join('\n'),
        );
        t.setFormat('code');
        pushInline(t);
        newPara();
        break;
      }
    }
  }

  return blocks.filter(
    (b, i) => !('getChildrenSize' in b) || (b as ParagraphNode).getChildrenSize() > 0 || i === 0,
  );
}

function flatText(s: BodySegment): string {
  switch (s.type) {
    case 'text': return s.text;
    case 'br': return '\n';
    case 'mention': return s.displayName;
    case 'code': return s.code;
    case 'link': return s.text;
    default: return '';
  }
}
