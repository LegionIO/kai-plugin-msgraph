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
import c from 'highlight.js/lib/languages/c';
import cpp from 'highlight.js/lib/languages/cpp';
import dart from 'highlight.js/lib/languages/dart';
import dockerfile from 'highlight.js/lib/languages/dockerfile';
import dos from 'highlight.js/lib/languages/dos';
import graphql from 'highlight.js/lib/languages/graphql';
import http from 'highlight.js/lib/languages/http';
import kotlin from 'highlight.js/lib/languages/kotlin';
import latex from 'highlight.js/lib/languages/latex';
import lisp from 'highlight.js/lib/languages/lisp';
import objectivec from 'highlight.js/lib/languages/objectivec';
import matlab from 'highlight.js/lib/languages/matlab';
import perl from 'highlight.js/lib/languages/perl';
import php from 'highlight.js/lib/languages/php';
import powershell from 'highlight.js/lib/languages/powershell';
import r from 'highlight.js/lib/languages/r';
import ruby from 'highlight.js/lib/languages/ruby';
import rust from 'highlight.js/lib/languages/rust';
import scala from 'highlight.js/lib/languages/scala';
import scss from 'highlight.js/lib/languages/scss';
import shell from 'highlight.js/lib/languages/shell';
import swift from 'highlight.js/lib/languages/swift';
import vbnet from 'highlight.js/lib/languages/vbnet';
import vbscript from 'highlight.js/lib/languages/vbscript';
import verilog from 'highlight.js/lib/languages/verilog';
import vhdl from 'highlight.js/lib/languages/vhdl';

const langs: Record<string, { fn: any; aliases?: string[] }> = {
  javascript: { fn: javascript, aliases: ['js', 'jsx'] },
  typescript: { fn: typescript, aliases: ['ts', 'tsx'] },
  python: { fn: python, aliases: ['py'] },
  json: { fn: json },
  yaml: { fn: yaml, aliases: ['yml'] },
  xml: { fn: xml, aliases: ['html', 'svg', 'jsp'] },
  css: { fn: css },
  bash: { fn: bash, aliases: ['sh', 'shell', 'zsh'] },
  sql: { fn: sql, aliases: ['kql'] },
  go: { fn: go, aliases: ['golang'] },
  java: { fn: java },
  csharp: { fn: csharp, aliases: ['cs', 'c#'] },
  markdown: { fn: markdown, aliases: ['md'] },
  diff: { fn: diff },
  plaintext: { fn: plaintext, aliases: ['text', 'txt'] },
  c: { fn: c },
  cpp: { fn: cpp, aliases: ['c++', 'cc'] },
  dart: { fn: dart },
  dockerfile: { fn: dockerfile, aliases: ['docker'] },
  dos: { fn: dos, aliases: ['bat', 'batch', 'cmd'] },
  graphql: { fn: graphql, aliases: ['gql'] },
  http: { fn: http },
  kotlin: { fn: kotlin, aliases: ['kt'] },
  latex: { fn: latex, aliases: ['tex'] },
  lisp: { fn: lisp },
  objectivec: { fn: objectivec, aliases: ['objective-c', 'objc'] },
  matlab: { fn: matlab, aliases: ['octave'] },
  perl: { fn: perl },
  php: { fn: php },
  powershell: { fn: powershell, aliases: ['ps1'] },
  r: { fn: r },
  ruby: { fn: ruby, aliases: ['rb'] },
  rust: { fn: rust },
  scala: { fn: scala },
  scss: { fn: scss },
  shell: { fn: shell },
  swift: { fn: swift },
  vbnet: { fn: vbnet, aliases: ['vb'] },
  vbscript: { fn: vbscript, aliases: ['vbs'] },
  verilog: { fn: verilog },
  vhdl: { fn: vhdl },
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
  c: 'C',
  dart: 'Dart',
  dockerfile: 'DockerFile',
  dos: 'DOS',
  graphql: 'GraphQL',
  html: 'HTML',
  http: 'HTTP',
  jsp: 'JSP',
  jsx: 'JSX',
  kotlin: 'Kotlin',
  kql: 'KQL',
  latex: 'LaTeX',
  lisp: 'Lisp',
  objectivec: 'Objective-C',
  octave: 'Octave',
  perl: 'Perl',
  php: 'PHP',
  powershell: 'PowerShell',
  r: 'R',
  ruby: 'Ruby',
  rust: 'Rust',
  scala: 'Scala',
  scss: 'SCSS',
  shell: 'Shell',
  swift: 'Swift',
  vbnet: 'VB.NET',
  vbscript: 'VBScript',
  verilog: 'Verilog',
  vhdl: 'VHDL',
};

function friendlyName(id: string | null, code: string): string | null {
  if (!id) return null;
  if (id === 'xml') {
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
    if (lang) return { html: escapeHtml(code), language: lang, displayName: friendlyName(lang, code) };
    const r = hljs.highlightAuto(code, AUTO_SUBSET);
    return { html: r.value, language: r.language ?? null, displayName: friendlyName(r.language ?? null, code) };
  } catch {
    return { html: escapeHtml(code), language: lang, displayName: friendlyName(lang, code) };
  }
}

function escapeHtml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}
