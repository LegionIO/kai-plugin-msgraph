import React from 'react';
import {
  DecoratorNode,
  TextNode,
  type EditorConfig,
  type LexicalNode,
  type NodeKey,
  type SerializedLexicalNode,
  type SerializedTextNode,
  type Spread,
} from 'lexical';

// ── MentionNode (TextNode subclass so caret/arrow/backspace work natively) ──

export type SerializedMentionNode = Spread<{ userId: string }, SerializedTextNode>;

export class MentionNode extends TextNode {
  __userId: string;

  static getType(): string {
    return 'mention';
  }
  static clone(node: MentionNode): MentionNode {
    return new MentionNode(node.__userId, node.__text, node.__key);
  }
  static importJSON(j: SerializedMentionNode): MentionNode {
    const n = new MentionNode(j.userId, j.text);
    n.setFormat(j.format);
    n.setDetail(j.detail);
    n.setStyle(j.style);
    return n;
  }

  constructor(userId: string, displayName: string, key?: NodeKey) {
    super(displayName, key);
    this.__userId = userId;
  }

  exportJSON(): SerializedMentionNode {
    return { ...super.exportJSON(), type: 'mention', version: 1, userId: this.__userId };
  }

  createDOM(config: EditorConfig): HTMLElement {
    const dom = super.createDOM(config);
    dom.style.color = '#2563eb';
    dom.style.background = 'rgba(37,99,235,0.10)';
    dom.style.borderRadius = '3px';
    dom.style.padding = '0 2px';
    dom.style.fontWeight = '500';
    dom.spellcheck = false;
    return dom;
  }

  isTextEntity(): true {
    return true;
  }
  canInsertTextBefore(): boolean {
    return false;
  }
  canInsertTextAfter(): boolean {
    return false;
  }

  getUserId(): string {
    return this.__userId;
  }
  getDisplayName(): string {
    return this.getTextContent();
  }
}

export function $createMentionNode(userId: string, displayName: string): MentionNode {
  const n = new MentionNode(userId, displayName);
  n.setMode('token');
  return n;
}
export function $isMentionNode(node: LexicalNode | null | undefined): node is MentionNode {
  return node instanceof MentionNode;
}

// ── ImageNode — stores a blob URL for synchronous insert; the backing File is
//    kept in a side-map keyed by tempId and converted to base64 at send time. ──

export const pendingImageFiles = new Map<string, File>();

export type SerializedImageNode = Spread<
  { tempId: string; contentType: string; src: string; name: string | null; existingUrl: string | null },
  SerializedLexicalNode
>;

export class ImageNode extends DecoratorNode<React.ReactNode> {
  __tempId: string;
  __contentType: string;
  __src: string;
  __name: string | null;
  /** When set, this is an already-hosted Graph image; serialize back to this URL verbatim (no new hostedContents). */
  __existingUrl: string | null;

  static getType(): string {
    return 'inline-image';
  }
  static clone(n: ImageNode): ImageNode {
    return new ImageNode(n.__tempId, n.__contentType, n.__src, n.__name, n.__existingUrl, n.__key);
  }
  static importJSON(j: SerializedImageNode): ImageNode {
    return new ImageNode(j.tempId, j.contentType, j.src, j.name, j.existingUrl ?? null);
  }

  constructor(
    tempId: string,
    contentType: string,
    src: string,
    name: string | null,
    existingUrl: string | null = null,
    key?: NodeKey,
  ) {
    super(key);
    this.__tempId = tempId;
    this.__contentType = contentType;
    this.__src = src;
    this.__name = name;
    this.__existingUrl = existingUrl;
  }

  exportJSON(): SerializedImageNode {
    return {
      type: 'inline-image',
      version: 1,
      tempId: this.__tempId,
      contentType: this.__contentType,
      src: this.__src,
      name: this.__name,
      existingUrl: this.__existingUrl,
    };
  }

  createDOM(): HTMLElement {
    // Render the image as real DOM so it lives inside contentEditable at the
    // caret position. decorate() is a no-op because Kai's runtime doesn't
    // expose ReactDOM.createPortal, which Lexical uses to mount decorator JSX.
    const wrap = document.createElement('span');
    wrap.style.display = 'inline-block';
    wrap.style.verticalAlign = 'text-bottom';
    wrap.contentEditable = 'false';
    const img = document.createElement('img');
    img.src = this.__src;
    img.alt = this.__name ?? '';
    img.draggable = false;
    img.style.maxWidth = '100%';
    img.style.maxHeight = '200px';
    img.style.borderRadius = '6px';
    img.style.display = 'block';
    wrap.appendChild(img);
    return wrap;
  }
  updateDOM(): false {
    return false;
  }
  isInline(): true {
    return true;
  }
  isKeyboardSelectable(): boolean {
    return true;
  }

  getTempId(): string {
    return this.__tempId;
  }
  getContentType(): string {
    return this.__contentType;
  }
  getSrc(): string {
    return this.__src;
  }
  getExistingUrl(): string | null {
    return this.__existingUrl;
  }

  decorate(): null {
    return null;
  }
}

export function $createImageNodeFromFile(file: File): ImageNode {
  const tempId = crypto.randomUUID?.() ?? String(Date.now() + Math.random());
  const src = URL.createObjectURL(file);
  pendingImageFiles.set(tempId, file);
  return new ImageNode(tempId, file.type || 'image/png', src, file.name || null, null);
}
/** For already-hosted Graph images being edited: displaySrc is the fetched data-URL, existingUrl is the Graph URL to emit. */
export function $createImageNodeFromExisting(existingUrl: string, displaySrc: string): ImageNode {
  return new ImageNode('', 'image/*', displaySrc, null, existingUrl);
}
export function $isImageNode(node: LexicalNode | null | undefined): node is ImageNode {
  return node instanceof ImageNode;
}
