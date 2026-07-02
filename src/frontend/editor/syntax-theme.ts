/**
 * Shared token colors for both Lexical's in-editor CodeHighlightNode classes
 * (`.msg-tok-*`) and highlight.js output (`.hljs-*`) in received messages.
 * Injected once as a <style> tag since Kai's Tailwind bundle won't include
 * arbitrary token classes.
 */

const light = {
  comment: '#6b7280',
  keyword: '#7c3aed',
  string: '#047857',
  number: '#b45309',
  function: '#2563eb',
  class: '#0d9488',
  attr: '#0369a1',
  tag: '#be185d',
  punct: '#64748b',
  builtin: '#9333ea',
  operator: '#475569',
  property: '#0e7490',
  variable: '#c2410c',
};

const dark = {
  comment: '#9ca3af',
  keyword: '#c4b5fd',
  string: '#6ee7b7',
  number: '#fbbf24',
  function: '#93c5fd',
  class: '#5eead4',
  attr: '#7dd3fc',
  tag: '#f9a8d4',
  punct: '#94a3b8',
  builtin: '#d8b4fe',
  operator: '#cbd5e1',
  property: '#67e8f9',
  variable: '#fdba74',
};

// Map both Lexical prism-token names and hljs class names to our palette keys.
const map: Record<string, keyof typeof light> = {
  // lexical/prism
  comment: 'comment', prolog: 'comment', cdata: 'comment', doctype: 'comment',
  keyword: 'keyword', selector: 'keyword', important: 'keyword', atrule: 'keyword',
  string: 'string', char: 'string', regex: 'string', inserted: 'string',
  number: 'number', boolean: 'number', constant: 'number', symbol: 'number',
  function: 'function', method: 'function',
  'class-name': 'class', class: 'class', namespace: 'class',
  attr: 'attr', 'attr-name': 'attr', 'attr-value': 'string',
  tag: 'tag', deleted: 'tag',
  punctuation: 'punct',
  builtin: 'builtin',
  operator: 'operator', entity: 'operator', url: 'operator',
  property: 'property',
  variable: 'variable', 'template-string': 'string',
  // hljs
  'hljs-comment': 'comment', 'hljs-quote': 'comment',
  'hljs-keyword': 'keyword', 'hljs-selector-tag': 'keyword', 'hljs-meta-keyword': 'keyword',
  'hljs-string': 'string', 'hljs-regexp': 'string', 'hljs-addition': 'string',
  'hljs-number': 'number', 'hljs-literal': 'number',
  'hljs-title': 'function', 'hljs-function': 'function',
  'hljs-class': 'class', 'hljs-type': 'class',
  'hljs-attr': 'attr', 'hljs-attribute': 'attr', 'hljs-selector-attr': 'attr',
  'hljs-name': 'tag', 'hljs-tag': 'tag', 'hljs-deletion': 'tag',
  'hljs-punctuation': 'punct',
  'hljs-built_in': 'builtin', 'hljs-builtin-name': 'builtin',
  'hljs-operator': 'operator', 'hljs-link': 'operator',
  'hljs-property': 'property', 'hljs-params': 'property',
  'hljs-variable': 'variable', 'hljs-template-variable': 'variable',
  'hljs-subst': 'variable',
};

// Lexical theme.codeHighlight expects tokenName → className.
export const lexicalCodeHighlightTheme: Record<string, string> = Object.fromEntries(
  Object.keys(map)
    .filter((k) => !k.startsWith('hljs-'))
    .map((k) => [k, `msg-tok-${map[k]}`]),
);

function rules(scheme: typeof light, selPrefix: string): string {
  const groups = new Map<keyof typeof light, Set<string>>();
  for (const [name, key] of Object.entries(map)) {
    const sel = name.startsWith('hljs-') ? `.${name}` : `.msg-tok-${key}`;
    if (!groups.has(key)) groups.set(key, new Set());
    groups.get(key)!.add(`${selPrefix}${sel}`);
  }
  return [...groups.entries()]
    .map(([key, sels]) => `${[...sels].join(',')}{color:${scheme[key]}}`)
    .join('\n');
}

const CSS = `
${rules(light, '')}
${rules(dark, '.dark ')}
.msg-editor-code[data-language]:not([data-language=""])::before {
  content: attr(data-language);
  display: block;
  font: 500 9px/1 ui-sans-serif, system-ui, sans-serif;
  letter-spacing: .06em;
  text-transform: uppercase;
  opacity: .55;
  margin-bottom: 6px;
}
`;

let injected = false;
export function ensureSyntaxThemeInjected(): void {
  if (injected || typeof document === 'undefined') return;
  const el = document.createElement('style');
  el.setAttribute('data-msgraph-syntax', '');
  el.textContent = CSS;
  document.head.appendChild(el);
  injected = true;
}
