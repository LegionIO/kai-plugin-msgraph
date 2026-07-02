import hljs from 'highlight.js/lib/core';
import javascript from 'highlight.js/lib/languages/javascript';
import typescript from 'highlight.js/lib/languages/typescript';
import python from 'highlight.js/lib/languages/python';
import json from 'highlight.js/lib/languages/json';
import yaml from 'highlight.js/lib/languages/yaml';
import xml from 'highlight.js/lib/languages/xml';
import css from 'highlight.js/lib/languages/css';
import bash from 'highlight.js/lib/languages/bash';
import sql from 'highlight.js/lib/languages/sql';
import go from 'highlight.js/lib/languages/go';
import java from 'highlight.js/lib/languages/java';
import csharp from 'highlight.js/lib/languages/csharp';
import markdown from 'highlight.js/lib/languages/markdown';
import diff from 'highlight.js/lib/languages/diff';
import plaintext from 'highlight.js/lib/languages/plaintext';

const langs: Record<string, { fn: any; aliases?: string[] }> = {
  javascript: { fn: javascript, aliases: ['js', 'jsx'] },
  typescript: { fn: typescript, aliases: ['ts', 'tsx'] },
  python: { fn: python, aliases: ['py'] },
  json: { fn: json },
  yaml: { fn: yaml, aliases: ['yml'] },
  xml: { fn: xml, aliases: ['html', 'svg'] },
  css: { fn: css },
  bash: { fn: bash, aliases: ['sh', 'shell', 'zsh'] },
  sql: { fn: sql },
  go: { fn: go, aliases: ['golang'] },
  java: { fn: java },
  csharp: { fn: csharp, aliases: ['cs', 'c#'] },
  markdown: { fn: markdown, aliases: ['md'] },
  diff: { fn: diff },
  plaintext: { fn: plaintext, aliases: ['text', 'txt'] },
};

for (const [name, { fn, aliases }] of Object.entries(langs)) {
  hljs.registerLanguage(name, fn);
  for (const a of aliases ?? []) hljs.registerAliases(a, { languageName: name });
}

const AUTO_SUBSET = ['typescript', 'javascript', 'python', 'json', 'yaml', 'xml', 'bash', 'sql', 'go', 'diff'];

const DISPLAY_NAME: Record<string, string> = {
  javascript: 'JavaScript',
  typescript: 'TypeScript',
  csharp: 'C#',
  cpp: 'C++',
  yaml: 'YAML',
  json: 'JSON',
  sql: 'SQL',
  xml: 'XML',
  css: 'CSS',
  bash: 'Bash',
  go: 'Go',
  python: 'Python',
  java: 'Java',
  markdown: 'Markdown',
  diff: 'Diff',
  plaintext: 'Text',
};

function friendlyName(id: string | null, code: string): string | null {
  if (!id) return null;
  if (id === 'xml') {
    // hljs canonicalizes html/svg to 'xml'. Prefer the more specific label when the content matches.
    if (/<svg\b/i.test(code)) return 'SVG';
    if (/<!doctype\s+html|<html\b|<head\b|<body\b|<div\b|<span\b|<script\b|<a\b|<p\b/i.test(code)) return 'HTML';
    return 'XML';
  }
  return DISPLAY_NAME[id] ?? id.charAt(0).toUpperCase() + id.slice(1);
}

export function highlight(code: string, lang: string | null): {
  html: string;
  language: string | null;
  displayName: string | null;
} {
  try {
    if (lang && hljs.getLanguage(lang)) {
      const r = hljs.highlight(code, { language: lang, ignoreIllegals: true });
      const id = r.language ?? lang;
      return { html: r.value, language: id, displayName: friendlyName(id, code) };
    }
    const r = hljs.highlightAuto(code, AUTO_SUBSET);
    return { html: r.value, language: r.language ?? null, displayName: friendlyName(r.language ?? null, code) };
  } catch {
    return { html: escapeHtml(code), language: lang, displayName: friendlyName(lang, code) };
  }
}

function escapeHtml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}
