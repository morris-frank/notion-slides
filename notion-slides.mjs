#!/usr/bin/env node
/**
 * Notion-style markdown → PowerPoint (single-file CLI).
 * Usage: notion-slides input.md --theme light --out output.pptx
 */

import fs from 'fs';
import path from 'path';
import crypto from 'crypto';
import { createRequire } from 'module';
import { fileURLToPath } from 'url';
import MarkdownIt from 'markdown-it';
import PptxGenJS from 'pptxgenjs';

const require = createRequire(import.meta.url);
const { init: initMathJax } = require('mathjax-full/es5/node-main.js');

// ---------------------------------------------------------------------------
// CLI
// ---------------------------------------------------------------------------

function parseArgs(argv) {
  const out = {
    input: null,
    theme: null,
    output: null,
    cacheDir: path.join(process.cwd(), '.cache'),
    offline: false,
    debugLayout: false,
  };
  const rest = [];
  for (let i = 2; i < argv.length; i++) {
    const a = argv[i];
    if (a === '--theme') out.theme = argv[++i];
    else if (a === '--out') out.output = argv[++i];
    else if (a === '--cache') out.cacheDir = path.resolve(argv[++i]);
    else if (a === '--offline') out.offline = true;
    else if (a === '--debug-layout') out.debugLayout = true;
    else if (!a.startsWith('-')) rest.push(a);
    else throw new Error(`Unknown flag: ${a}`);
  }
  if (rest.length < 1) throw new Error('Usage: node notion-slides.mjs input.md --theme light|dark|theme.json --out output.pptx');
  out.input = path.resolve(rest[0]);
  if (!out.theme) throw new Error('Missing --theme (light, dark, or path to theme.json)');
  if (!out.output) throw new Error('Missing --out output.pptx');
  // Bare "light"/"dark" pick the bundled theme next to this script.
  const bundled = new URL(`theme.${out.theme}.json`, import.meta.url);
  out.theme = /^(light|dark)$/.test(out.theme) ? fileURLToPath(bundled) : path.resolve(out.theme);
  out.output = path.resolve(out.output);
  return out;
}

// ---------------------------------------------------------------------------
// Theme
// ---------------------------------------------------------------------------

function loadTheme(themePath) {
  const raw = fs.readFileSync(themePath, 'utf8');
  const t = JSON.parse(raw);
  const req = ['meta', 'slide', 'fonts', 'typeScale', 'layout', 'text', 'blocks'];
  for (const k of req) {
    if (!t[k]) throw new Error(`Theme missing required section: ${k}`);
  }
  if (!t.slide.layout) t.slide.layout = 'LAYOUT_WIDE';
  t.emojiMap = t.emojiMap || {};
  t.layout.footerHeight = t.layout.footerHeight ?? 0.34;
  t.headerFooter = {
    footerCenter: '',
    footerFontPt: 9,
    footerColor: '#28322e',
    footerMiddleFraction: 0.6,
    footerRemainderLeftFraction: 0.62,
    logoMaxHeight: 0.38,
    logoSlotWidth: 1.05,
    logoTitleGap: 0.1,
    ...(t.headerFooter || {}),
  };
  return t;
}

// ---------------------------------------------------------------------------
// Markdown preprocessing
// ---------------------------------------------------------------------------

function sha256(s) {
  return crypto.createHash('sha256').update(s).digest('hex');
}

function transformOutsideTripleBackticks(src, fn) {
  const re = /```[\s\S]*?```/g;
  let out = '';
  let last = 0;
  let m;
  while ((m = re.exec(src))) {
    out += fn(src.slice(last, m.index));
    out += m[0];
    last = m.index + m[0].length;
  }
  out += fn(src.slice(last));
  return out;
}

function rewriteDisplayMath(src) {
  return transformOutsideTripleBackticks(src, (chunk) =>
    chunk.replace(/\$\$([\s\S]*?)\$\$/g, (_, tex) => `\n\n\`\`\`math\n${tex.trim()}\n\`\`\`\n\n`),
  );
}

function extractAsides(src) {
  const asides = [];
  const re = /<aside>([\s\S]*?)<\/aside>/gi;
  const replaced = src.replace(re, (_match, innerRaw) => {
    const inner = innerRaw.trim();
    if (/<aside/i.test(inner)) throw new Error('Nested <aside> is not supported');
    const idx = asides.length;
    asides.push(inner);
    return `\n\n<div data-aside="${idx}"></div>\n\n`;
  });
  return { text: replaced, asides };
}

// ---------------------------------------------------------------------------
// AST types (JSDoc for readability)
// ---------------------------------------------------------------------------

/** @typedef {{ kind:'paragraph', inlines:any[] }} ParagraphBlock */
/** @typedef {{ kind:'bullets', items:any[][] }} BulletListBlock */
/** @typedef {{ kind:'code-block', text:string, language?:string }} CodeBlock */
/** @typedef {{ kind:'math-block', tex:string }} MathBlock */
/** @typedef {{ kind:'callout', icon?:{src:string, widthPx?:number}, blocks:any[] }} CalloutBlock */
/** @typedef {{ kind:'subheading', inlines:any[] }} SubheadingBlock */

// ---------------------------------------------------------------------------
// Parse slide body → blocks
// ---------------------------------------------------------------------------

function parseAsideInner(innerMd, md, themeDir, emojiMap) {
  const imgM = innerMd.match(/<img\s+[^>]*src=["']([^"']+)["'][^>]*(?:width=["']([^"']+)["'])?[^>]*\/?>/i);
  let icon;
  let rest = innerMd;
  if (imgM) {
    const src = imgM[1];
    const wpx = imgM[2] ? parsePx(imgM[2]) : undefined;
    icon = { src, widthPx: wpx };
    rest = innerMd.replace(imgM[0], '').trim();
  } else {
    const first = innerMd.trimStart();
    const cp = first.codePointAt(0);
    if (cp) {
      const ch = String.fromCodePoint(cp);
      if (emojiMap[ch]) {
        const ref = emojiMap[ch];
        icon = { src: /^https?:/i.test(ref) ? ref : path.resolve(themeDir, ref) };
        rest = [...first].slice(1).join('').trimStart();
      }
    }
  }
  if (/<aside/i.test(rest)) throw new Error('Nested <aside> inside callout is not supported');
  const blocks = parseBlocksFromMarkdown(rest, md, { themeDir, emojiMap, asides: [] });
  return { kind: 'callout', icon, blocks };
}

function parsePx(s) {
  const m = String(s).match(/^(\d+(?:\.\d+)?)\s*px$/i);
  return m ? Number(m[1]) : undefined;
}

function parseBlocksFromMarkdown(body, md, ctx) {
  const tokens = md.parse(body, {});
  return walkBlockTokens(tokens, 0, md, ctx).blocks;
}

function walkBlockTokens(tokens, i, md, ctx) {
  const blocks = [];
  while (i < tokens.length) {
    const t = tokens[i];
    if (t.type === 'heading_open' && /^h[1-6]$/.test(t.tag)) {
      const inlineTok = tokens[i + 1];
      i += 3;
      blocks.push({
        kind: 'subheading',
        inlines: walkInlines(inlineTok.children || [], md, ctx),
      });
      continue;
    }
    if (t.type === 'paragraph_open') {
      const inlineTok = tokens[i + 1];
      i += 3;
      blocks.push({ kind: 'paragraph', inlines: walkInlines(inlineTok.children || [], md, ctx) });
      continue;
    }
    if (t.type === 'bullet_list_open') {
      const { items, next } = parseBulletList(tokens, i, md, ctx);
      blocks.push({ kind: 'bullets', items });
      i = next;
      continue;
    }
    if (t.type === 'fence') {
      const lang = (t.info || '').trim();
      if (lang === 'math') blocks.push({ kind: 'math-block', tex: t.content.trim() });
      else blocks.push({ kind: 'code-block', text: t.content.replace(/\n$/, ''), language: lang || undefined });
      i++;
      continue;
    }
    if (t.type === 'html_block' && t.content.includes('data-aside=')) {
      const m = t.content.match(/data-aside="(\d+)"/);
      if (!m) {
        console.warn('[warn] Unrecognized aside placeholder, skipping');
        i++;
        continue;
      }
      const idx = Number(m[1]);
      const inner = ctx.asides[idx];
      if (inner == null) throw new Error(`Aside index ${idx} out of range`);
      blocks.push(parseAsideInner(inner, md, ctx.themeDir, ctx.emojiMap));
      i++;
      continue;
    }
    if (t.type === 'html_block') {
      console.warn('[warn] Dropping unsupported html_block');
      i++;
      continue;
    }
    if (t.type === 'table_open') {
      console.warn('[warn] Tables are not supported in v1, skipping');
      while (tokens[i].type !== 'table_close') i++;
      i++;
      continue;
    }
    if (t.type === 'blockquote_open') {
      throw new Error('Blockquotes are not supported in v1');
    }
    if (t.type === 'hr') {
      i++;
      continue;
    }
    if (t.hidden) {
      i++;
      continue;
    }
    console.warn(`[warn] Unsupported token ${t.type}, skipping`);
    if (t.nesting === 1) {
      let depth = 1;
      i++;
      while (i < tokens.length && depth > 0) {
        depth += tokens[i].nesting;
        i++;
      }
    } else {
      i++;
    }
  }
  return { blocks, i };
}

function parseBulletList(tokens, startIdx, md, ctx) {
  const items = [];
  let i = startIdx + 1;
  while (tokens[i]?.type !== 'bullet_list_close') {
    if (tokens[i]?.type === 'list_item_open') {
      i++;
      if (tokens[i]?.type === 'paragraph_open') {
        const inl = tokens[i + 1];
        items.push(walkInlines(inl.children || [], md, ctx));
        i += 3;
      } else {
        items.push([]);
      }
      if (tokens[i]?.type === 'list_item_close') i++;
    } else i++;
  }
  return { items, next: i + 1 };
}

function splitInlineMath(text) {
  const parts = text.split('$');
  if (parts.length < 3) return [{ kind: 'text', text }];
  const out = [];
  for (let k = 0; k < parts.length; k++) {
    if (k % 2 === 0) {
      if (parts[k]) out.push({ kind: 'text', text: parts[k] });
    } else {
      if (parts[k].includes('\n')) return [{ kind: 'text', text }];
      const tex = parts[k].trim();
      if (tex) out.push({ kind: 'math-inline', tex });
    }
  }
  return out;
}

function applyTextStyle(parts, active) {
  return parts.map((p) => {
    if (p.kind === 'text' && active.strong) return { kind: 'strong', text: p.text };
    if (p.kind === 'text' && active.em) return { kind: 'em', text: p.text };
    return p;
  });
}

function walkInlines(children, md, ctx, active = { strong: false, em: false }) {
  const res = [];
  if (!children) return res;
  let i = 0;
  while (i < children.length) {
    const c = children[i];
    if (c.type === 'text') {
      res.push(...applyTextStyle(splitInlineMath(c.content), active));
      i++;
    } else if (c.type === 'strong_open') {
      const inner = [];
      i++;
      while (i < children.length && children[i].type !== 'strong_close') {
        inner.push(children[i]);
        i++;
      }
      const sub = walkInlines(inner, md, ctx, { ...active, strong: true });
      for (const x of sub) {
        if (x.kind === 'text') res.push({ kind: 'strong', text: x.text });
        else if (x.kind === 'math-inline') res.push(x);
        else if (x.kind === 'code') res.push(x);
        else if (x.kind === 'link') res.push(x);
        else if (x.kind === 'strong' || x.kind === 'em') res.push(x);
      }
      if (children[i]?.type === 'strong_close') i++;
    } else if (c.type === 'em_open') {
      const inner = [];
      i++;
      while (i < children.length && children[i].type !== 'em_close') {
        inner.push(children[i]);
        i++;
      }
      const sub = walkInlines(inner, md, ctx, { ...active, em: true });
      for (const x of sub) {
        if (x.kind === 'text') res.push({ kind: 'em', text: x.text });
        else res.push(x);
      }
      if (children[i]?.type === 'em_close') i++;
    } else if (c.type === 'code_inline') {
      res.push({ kind: 'code', text: c.content });
      i++;
    } else if (c.type === 'softbreak' || c.type === 'hardbreak') {
      res.push({ kind: 'text', text: ' ' });
      i++;
    } else if (c.type === 'link_open') {
      const href = c.attrs?.find((a) => a[0] === 'href')?.[1] || '';
      const inner = [];
      i++;
      while (i < children.length && children[i].type !== 'link_close') {
        inner.push(children[i]);
        i++;
      }
      const sub = walkInlines(inner, md, ctx, active);
      const label = sub
        .filter((x) => x.kind === 'text' || x.kind === 'strong' || x.kind === 'em')
        .map((x) => x.text || '')
        .join('');
      res.push({ kind: 'link', text: label || href, url: href });
      if (children[i]?.type === 'link_close') i++;
    } else i++;
  }
  return res;
}

function mergeAdjacentText(inlines) {
  const out = [];
  for (const x of inlines) {
    if (out.length && x.kind === 'text' && out[out.length - 1].kind === 'text') {
      out[out.length - 1].text += x.text;
    } else out.push({ ...x });
  }
  return out;
}

// ---------------------------------------------------------------------------
// Split slides + deck parse
// ---------------------------------------------------------------------------

function readMarkdown(file) {
  return fs.readFileSync(file, 'utf8');
}

function stripBom(s) {
  return String(s).replace(/^[\uFEFF\u200E\u200F\u202A\u202B\u202C]+/, '').trim();
}

/**
 * First `#` / `##` / `###` at line start (outside `<aside>` and outside ``` fences)
 * becomes the slide title; that line is removed. Lines before it stay in the body.
 */
function peelSlideTitle(chunk) {
  const lines = String(chunk).replace(/\r\n/g, '\n').split('\n');
  let asideDepth = 0;
  let inFence = false;
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const opens = (line.match(/<aside\b/gi) || []).length;
    const closes = (line.match(/<\/aside>/gi) || []).length;
    asideDepth = Math.max(0, asideDepth + opens - closes);
    if (/^\s*```/.test(line)) {
      inFence = !inFence;
      continue;
    }
    if (asideDepth !== 0 || inFence) continue;
    if (!line.trim()) continue;
    const m = line.match(/^(#{1,3})\s+(.+)$/);
    if (m) {
      const title = stripBom(m[2]).trim();
      const body = [...lines.slice(0, i), ...lines.slice(i + 1)].join('\n').trim();
      return { title, body };
    }
    break;
  }
  return { title: '', body: chunk.trim() };
}

/**
 * Deck title slide: only the first `# …` at the top of the file.
 * Everything after that line is split exclusively on `---` (horizontal rules);
 * each non-empty chunk is one slide. Another `#` inside a chunk is allowed as that
 * slide’s title line (see peelSlideTitle).
 */
function splitSlides(raw) {
  const normalized = raw.replace(/\r\n/g, '\n');
  const firstH1 = normalized.match(/^[\s\uFEFF\u200E\u200F]*#\s+([^\n]+)/);
  if (!firstH1) throw new Error('Missing deck title: document must start with `# ...`');
  const deckTitle = stripBom(firstH1[1]).trim();
  const afterH1 = normalized.slice(firstH1.index + firstH1[0].length).replace(/^\n+/, '');
  const rawChunks = afterH1.length ? afterH1.split(/\n---\n/) : [];

  const slides = [];
  for (let i = 0; i < rawChunks.length; i++) {
    const chunk = rawChunks[i].trim();
    if (!chunk) continue;
    slides.push(peelSlideTitle(chunk));
  }
  return { title: deckTitle, slides };
}

function parseSlideBody(body, md, themeDir, emojiMap) {
  const ex = extractAsides(body);
  const text = rewriteDisplayMath(ex.text);
  const ctx = { themeDir, emojiMap, asides: ex.asides };
  return parseBlocksFromMarkdown(text, md, ctx);
}

// ---------------------------------------------------------------------------
// Assets + Math
// ---------------------------------------------------------------------------

let _mj = null;
async function ensureMathJax() {
  if (!_mj) _mj = await initMathJax({ loader: { load: ['input/tex', 'output/svg'] } });
  return _mj;
}

function svgDimsFromMathJax(svgXml, bodyPt, maxW, display) {
  let wEx = 2;
  let hEx = 1;
  const wm = svgXml.match(/width="([\d.]+)ex"/);
  const hm = svgXml.match(/height="([\d.]+)ex"/);
  if (wm) wEx = parseFloat(wm[1]);
  if (hm) hEx = parseFloat(hm[1]);
  const exIn = (bodyPt / 72) * 0.5;
  let w = wEx * exIn;
  let h = hEx * exIn;
  if (display && w > maxW) {
    const r = maxW / w;
    w *= r;
    h *= r;
  }
  if (!display) {
    const cap = (bodyPt / 72) * 1.25;
    if (h > cap) {
      const r = cap / h;
      w *= r;
      h *= r;
    }
  }
  return { wIn: Math.max(0.06, w), hIn: Math.max(0.06, h) };
}

async function renderMathToSvgFile(tex, display, { cacheDir, offline }) {
  const key = `${display ? 'D' : 'I'}:${tex}`;
  const hash = sha256(key);
  const dir = path.join(cacheDir, 'math');
  fs.mkdirSync(dir, { recursive: true });
  const outPath = path.join(dir, `sha256_${hash}.svg`);
  if (fs.existsSync(outPath)) {
    return { path: outPath };
  }
  if (offline) throw new Error(`Offline mode: missing math cache for:\n${tex}`);
  let MJ;
  try {
    MJ = await ensureMathJax();
  } catch (e) {
    throw new Error(`MathJax failed to initialize: ${e.message}`);
  }
  let node;
  try {
    node = MJ.tex2svg(tex, { display });
  } catch (e) {
    throw new Error(`Malformed LaTeX (MathJax): ${tex}\n${e.message}`);
  }
  const html = MJ.startup.adaptor.outerHTML(node);
  const m = html.match(/<svg[\s\S]*?<\/svg>/);
  const svg = m ? m[0] : html;
  fs.writeFileSync(outPath, svg, 'utf8');
  return { path: outPath };
}

function isBlockedFetchUrl(href) {
  try {
    const u = new URL(href);
    if (u.protocol !== 'http:' && u.protocol !== 'https:') return true;
    const h = u.hostname.toLowerCase();
    if (h === 'localhost' || h === '0.0.0.0' || h.endsWith('.localhost')) return true;
    const ip = h.match(/^(\d{1,3})\.(\d{1,3})\.(\d{1,3})\.(\d{1,3})$/);
    if (ip) {
      const a = Number(ip[1]);
      const b = Number(ip[2]);
      if (a === 10) return true;
      if (a === 127) return true;
      if (a === 0) return true;
      if (a === 169 && b === 254) return true;
      if (a === 192 && b === 168) return true;
      if (a === 172 && b >= 16 && b <= 31) return true;
    }
    return false;
  } catch {
    return true;
  }
}

/** Notion exports often reference /icons/*.svg; fetch from notion.so if missing locally. */
function tryNotionHostedIconUrl(src) {
  let p = String(src).trim().replace(/\\/g, '/');
  if (p.startsWith('./')) p = p.slice(2);
  if (!/\.(svg|png|jpe?g|webp|gif)$/i.test(p)) return null;
  if (p.startsWith('/icons/')) return `https://www.notion.so${p}`;
  if (/^icons\//i.test(p)) return `https://www.notion.so/${p}`;
  return null;
}

async function resolveAssetSrc(src, { cacheDir, offline, themeDir }) {
  if (/^https?:\/\//i.test(src)) {
    if (isBlockedFetchUrl(src)) throw new Error(`Blocked URL (local/private): ${src}`);
    const hash = sha256(src);
    const dir = path.join(cacheDir, 'assets');
    fs.mkdirSync(dir, { recursive: true });
    const tryExt = (ext) => path.join(dir, `sha256_${hash}.${ext}`);
    for (const ext of ['svg', 'png', 'jpg', 'jpeg', 'webp', 'gif']) {
      const p = tryExt(ext);
      if (fs.existsSync(p)) return p;
    }
    if (offline) throw new Error(`Offline: missing cached asset for ${src}`);
    const res = await fetch(src);
    if (!res.ok) throw new Error(`Failed to download ${src}: ${res.status}`);
    const buf = Buffer.from(await res.arrayBuffer());
    let ext = 'bin';
    const ct = (res.headers.get('content-type') || '').toLowerCase();
    if (ct.includes('svg')) ext = 'svg';
    else if (ct.includes('png')) ext = 'png';
    else if (ct.includes('jpeg') || ct.includes('jpg')) ext = 'jpg';
    else if (src.toLowerCase().includes('.svg')) ext = 'svg';
    else if (src.toLowerCase().includes('.png')) ext = 'png';
    else if (buf.length > 4 && buf[0] === 0x89 && buf[1] === 0x50) ext = 'png';
    const dest = tryExt(ext);
    fs.writeFileSync(dest, buf);
    return dest;
  }
  const local = path.isAbsolute(src) ? src : path.resolve(themeDir, src);
  if (!fs.existsSync(local)) {
    const notionUrl = tryNotionHostedIconUrl(src);
    if (notionUrl) return resolveAssetSrc(notionUrl, { cacheDir, offline, themeDir });
    throw new Error(`Local asset not found: ${local}`);
  }
  return local;
}

// ---------------------------------------------------------------------------
// Layout estimation (inches)
// ---------------------------------------------------------------------------

const LH = { title: 1.1, body: 1.2, code: 1.15, callout: 1.2, sub: 1.15 };

function charWIn(fontPt) {
  return (fontPt / 72) * 0.52;
}

function monoWIn(fontPt) {
  return (fontPt / 72) * 0.6;
}

function estimateWrappedLines(text, fontPt, widthIn) {
  if (!text.trim()) return 1;
  const cw = charWIn(fontPt);
  const words = text.split(/\s+/).filter(Boolean);
  let lineUsed = 0;
  let lines = 1;
  for (const w of words) {
    const need = w.length * cw + (lineUsed > 0 ? cw : 0);
    if (lineUsed + need > widthIn && lineUsed > 0) {
      lines++;
      lineUsed = w.length * cw;
    } else lineUsed += need;
  }
  return lines;
}

function inlinePlainText(inlines) {
  return inlines
    .map((x) => {
      if (x.kind === 'text' || x.kind === 'strong' || x.kind === 'em') return x.text || '';
      if (x.kind === 'code') return x.text;
      if (x.kind === 'link') return x.text;
      if (x.kind === 'math-inline') return 'M'; // placeholder width
      return '';
    })
    .join('');
}

async function estimateBlockHeight(block, theme, contentW, cacheCtx, mathPrep) {
  const ts = theme.typeScale;
  const ly = theme.layout;
  const bodyPt = ts.bodyPt;
  const bulletPt = ts.bulletPt;
  const subPt = ts.subheadingPt;
  const codePt = ts.codeBlockPt;
  const icPt = ts.inlineCodePt;

  if (block.kind === 'paragraph') {
    const merged = mergeAdjacentText(block.inlines);
    const hasMath = merged.some((x) => x.kind === 'math-inline');
    if (!hasMath) {
      const t = inlinePlainText(merged);
      const lines = estimateWrappedLines(t, bodyPt, contentW);
      return (bodyPt / 72) * LH.body * lines;
    }
    const lineH = (bodyPt / 72) * LH.body;
    let x = 0;
    let line = 1;
    let rowH = lineH;
    for (const part of merged) {
      if (part.kind === 'math-inline') {
        const m = await mathPrep(part.tex, false);
        const w = m.wIn;
        const h = m.hIn;
        if (x + w > contentW && x > 0) {
          line++;
          x = 0;
          rowH = lineH;
        }
        x += w + 0.05;
        rowH = Math.max(rowH, h);
      } else {
        const txt =
          part.kind === 'code'
            ? part.text
            : part.kind === 'text' || part.kind === 'strong' || part.kind === 'em'
              ? part.text
              : part.kind === 'link'
                ? part.text
                : '';
        const pieceW = part.kind === 'code' ? monoWIn(icPt) * txt.length : charWIn(bodyPt) * txt.length;
        if (x + pieceW > contentW && x > 0) {
          line++;
          x = 0;
          rowH = lineH;
        }
        x += pieceW;
        rowH = Math.max(rowH, lineH);
      }
    }
    return line * rowH + (line - 1) * 0.02;
  }
  if (block.kind === 'subheading') {
    const t = inlinePlainText(mergeAdjacentText(block.inlines));
    const lines = estimateWrappedLines(t, subPt, contentW);
    return (subPt / 72) * LH.sub * lines;
  }
  if (block.kind === 'bullets') {
    const indent = 0.35;
    const w = contentW - indent;
    let h = 0;
    for (const item of block.items) {
      const merged = mergeAdjacentText(item);
      const t = inlinePlainText(merged);
      const lines = estimateWrappedLines(t, bulletPt, w);
      h += (bulletPt / 72) * LH.body * lines + ly.bulletGap;
    }
    return h;
  }
  if (block.kind === 'code-block') {
    const lines = block.text.split('\n').length;
    const lineH = (codePt / 72) * LH.code;
    const pad = ly.codePadding * 2;
    return pad + lineH * lines + 0.05;
  }
  if (block.kind === 'math-block') {
    const m = await mathPrep(block.tex, true);
    return m.hIn + ly.mathBlockGap;
  }
  if (block.kind === 'callout') {
    const pad = ly.calloutPadding * 2;
    const iconW = block.icon ? ly.calloutIconSize + ly.calloutIconGap : 0;
    const innerW = contentW - pad * 2 - iconW;
    let innerH = 0;
    for (const b of block.blocks) {
      innerH += (await estimateBlockHeight(b, theme, innerW, cacheCtx, mathPrep)) + ly.calloutGap * 0.2;
    }
    const iconBox = block.icon ? ly.calloutIconSize : 0;
    return pad + Math.max(innerH, iconBox) + 0.06;
  }
  return 0;
}

async function estimateSlideBlocksHeight(blocks, theme, contentW, cacheCtx, mathPrep) {
  const ly = theme.layout;
  let total = 0;
  for (const b of blocks) {
    total += (await estimateBlockHeight(b, theme, contentW, cacheCtx, mathPrep)) + ly.paragraphGap;
  }
  return total;
}

// ---------------------------------------------------------------------------
// Render
// ---------------------------------------------------------------------------

function hexNoHash(c) {
  return String(c).replace(/^#/, '');
}

function firstExistingInDir(themeDir, basenames) {
  for (const name of basenames) {
    const p = path.join(themeDir, name);
    if (fs.existsSync(p)) return p;
  }
  return null;
}

function slideBackgroundIsDark(theme) {
  const bg = theme?.slide?.background;
  if (!bg) return false;
  const hex = String(bg).replace(/^#/, '');
  if (!/^[0-9a-f]{6}$/i.test(hex)) return false;
  const r = parseInt(hex.slice(0, 2), 16);
  const g = parseInt(hex.slice(2, 4), 16);
  const b = parseInt(hex.slice(4, 6), 16);
  const lum = (0.2126 * r + 0.7152 * g + 0.0722 * b) / 255;
  return lum < 0.45;
}

/** Dark theme: meta name, filename, or dark slide background (lum). */
function isDarkThemeStyle(theme, themePath) {
  const name = String(theme?.meta?.name || '').toLowerCase();
  if (name.includes('dark')) return true;
  const stem = path.basename(themePath || '', path.extname(themePath || '')).toLowerCase();
  if (stem.includes('dark')) return true;
  return slideBackgroundIsDark(theme);
}

/**
 * logo.light.png / logo.dark.png (or .svg) by theme; then logo.png / logo.svg;
 * then the opposite mode as last resort.
 */
function resolveLogoPath(themeDir, theme, themePath) {
  if (theme?.headerFooter?.logo === false) return null;
  const dark = isDarkThemeStyle(theme, themePath);
  const mode = dark ? 'dark' : 'light';
  const other = dark ? 'light' : 'dark';
  return (
    firstExistingInDir(themeDir, [`logo.${mode}.png`, `logo.${mode}.svg`]) ||
    firstExistingInDir(themeDir, ['logo.png', 'logo.svg']) ||
    firstExistingInDir(themeDir, [`logo.${other}.png`, `logo.${other}.svg`]) ||
    null
  );
}

function readPngDimensions(filePath) {
  const buf = fs.readFileSync(filePath);
  if (buf.length < 24) return null;
  if (buf[0] !== 0x89 || buf[1] !== 0x50 || buf[2] !== 0x4e || buf[3] !== 0x47) return null;
  const w = buf.readUInt32BE(16);
  const h = buf.readUInt32BE(20);
  if (!w || !h) return null;
  return { w, h };
}

function readSvgDimensions(filePath) {
  const s = fs.readFileSync(filePath, 'utf8');
  const vb = s.match(/viewBox\s*=\s*["']\s*[\d.\s-]+\s+[\d.\s-]+\s+([\d.]+)\s+([\d.]+)\s*["']/i);
  if (vb) {
    const w = parseFloat(vb[1]);
    const h = parseFloat(vb[2]);
    if (w > 0 && h > 0) return { w, h };
  }
  const wm = s.match(/\bwidth\s*=\s*["']([\d.]+)(?:px)?["']/i);
  const hm = s.match(/\bheight\s*=\s*["']([\d.]+)(?:px)?["']/i);
  if (wm && hm) {
    const w = parseFloat(wm[1]);
    const h = parseFloat(hm[1]);
    if (w > 0 && h > 0) return { w, h };
  }
  return null;
}

function getLogoIntrinsicSize(logoPath) {
  const ext = path.extname(logoPath).toLowerCase();
  if (ext === '.png') return readPngDimensions(logoPath);
  if (ext === '.svg') return readSvgDimensions(logoPath);
  return null;
}

/**
 * Height = logoMaxHeight; width from aspect ratio (no letterboxing).
 * titleReserve: space to leave beside slide title so it does not run under the logo.
 */
function computeLogoLayout(logoPath, hf) {
  if (!logoPath) return null;
  const maxH = hf.logoMaxHeight;
  const gap = hf.logoTitleGap ?? 0.1;
  const dims = getLogoIntrinsicSize(logoPath);
  let logoW;
  let logoH = maxH;
  if (dims && dims.h > 0) {
    logoW = maxH * (dims.w / dims.h);
  } else {
    logoW = hf.logoSlotWidth;
  }
  return { path: logoPath, w: logoW, h: logoH, titleReserve: logoW + gap };
}

function addSlideChrome(slide, theme, { deckTitle, slideNumber, logoLayout }) {
  const SLIDE_W = 13.333;
  const SLIDE_H = 7.5;
  const ly = theme.layout;
  const hf = theme.headerFooter;
  const footerH = ly.footerHeight;
  const footY = SLIDE_H - ly.marginBottom - footerH;
  const fw = SLIDE_W - ly.marginLeft - ly.marginRight;
  const midFrac = hf.footerMiddleFraction ?? 0.6;
  const wMid = fw * midFrac;
  const rem = fw - wMid;
  const leftShare = hf.footerRemainderLeftFraction ?? 0.62;
  const wLeft = rem * leftShare;
  const wRight = rem - wLeft;
  const fx0 = ly.marginLeft;
  const pt = hf.footerFontPt;
  const color = hexNoHash(hf.footerColor);
  const face = theme.fonts.body;

  slide.addText(deckTitle && deckTitle.trim() ? deckTitle : '\u200b', {
    x: fx0,
    y: footY,
    w: wLeft,
    h: footerH - 0.02,
    fontSize: pt,
    fontFace: face,
    color,
    align: 'left',
    valign: 'middle',
    fit: 'shrink',
  });
  slide.addText(hf.footerCenter && hf.footerCenter.trim() ? hf.footerCenter : '\u200b', {
    x: fx0 + wLeft,
    y: footY,
    w: wMid,
    h: footerH - 0.02,
    fontSize: pt,
    fontFace: face,
    color,
    align: 'center',
    valign: 'middle',
    fit: 'shrink',
  });
  slide.addText(String(slideNumber), {
    x: fx0 + wLeft + wMid,
    y: footY,
    w: wRight,
    h: footerH - 0.02,
    fontSize: pt,
    fontFace: face,
    color,
    align: 'right',
    valign: 'middle',
    fit: 'shrink',
  });

  if (logoLayout) {
    slide.addImage({
      path: logoLayout.path,
      x: SLIDE_W - ly.marginRight - logoLayout.w,
      y: ly.marginTop,
      w: logoLayout.w,
      h: logoLayout.h,
    });
  }
}

/** Slide stub so renderBlock can measure heights without drawing (callout sizing). */
const NOOP_SLIDE = {
  addText() {
    return this;
  },
  addImage() {
    return this;
  },
  addShape() {
    return this;
  },
  background: {},
};

async function renderDeck(deckAst, theme, pptx, opts) {
  const SLIDE_W = 13.333;
  const SLIDE_H = 7.5;
  const ly = theme.layout;
  const ts = theme.typeScale;
  const cacheDir = opts.cacheDir;
  const themeDir = path.dirname(opts.themePath);
  fs.mkdirSync(cacheDir, { recursive: true });

  const contentW = SLIDE_W - ly.marginLeft - ly.marginRight;
  const footerH = ly.footerHeight;
  const contentBottom = SLIDE_H - ly.marginBottom - footerH;
  const titleReserve = opts.logoLayout?.titleReserve ?? 0;
  const titleW = contentW - titleReserve;

  const mathPathCache = new Map();
  const mathPrep = async (tex, display) => {
    const k = `${display ? 'D' : 'I'}|${tex}`;
    let svgPath = mathPathCache.get(k);
    if (!svgPath) {
      const r = await renderMathToSvgFile(tex, display, { cacheDir, offline: opts.offline });
      svgPath = r.path;
      mathPathCache.set(k, svgPath);
    }
    const xml = fs.readFileSync(svgPath, 'utf8');
    return { path: svgPath, ...svgDimsFromMathJax(xml, ts.bodyPt, contentW * 0.92, display) };
  };

  for (let si = 0; si < deckAst.slides.length; si++) {
    const slideSpec = deckAst.slides[si];
    const hasTitle = Boolean(slideSpec.title && slideSpec.title.trim());
    const contentTop0 = ly.marginTop + (hasTitle ? ly.titleHeight + ly.contentGapAfterTitle : 0);
    const contentHAvail = contentBottom - contentTop0;

    const slide = pptx.addSlide();
    slide.background = { color: hexNoHash(theme.slide.background) };

    if (hasTitle) {
      slide.addText(slideSpec.title, {
        x: ly.marginLeft,
        y: ly.marginTop,
        w: titleW,
        h: ly.titleHeight,
        fontSize: ts.titlePt,
        fontFace: theme.fonts.title,
        color: hexNoHash(theme.text.titleColor),
        bold: false,
        align: 'left',
        valign: 'top',
      });
    }

    let y = contentTop0;
    if (opts.debugLayout) {
      slide.addShape(pptx.ShapeType.rect, {
        x: ly.marginLeft,
        y: contentTop0,
        w: contentW,
        h: contentHAvail,
        fill: { color: '0088CC', transparency: 92 },
        line: { color: '0088CC', width: 0.5 },
      });
      const est = await estimateSlideBlocksHeight(slideSpec.blocks, theme, contentW, {}, mathPrep);
      console.error(
        `[debug-layout] slide ${si + 1}: content box ${contentHAvail.toFixed(3)} in tall, estimated blocks ${est.toFixed(3)} in (overflow allowed)`,
      );
    }

    for (const block of slideSpec.blocks) {
      y += await renderBlock(slide, block, {
        pptx,
        theme,
        x0: ly.marginLeft,
        y0: y,
        w: contentW,
        cacheDir,
        offline: opts.offline,
        themeDir,
        mathPrep,
      });
      y += ly.paragraphGap;
    }

    addSlideChrome(slide, theme, {
      deckTitle: deckAst.title,
      slideNumber: si + 2,
      logoLayout: opts.logoLayout,
    });
  }
}

async function renderParagraphMixed(slide, inlines, ctx) {
  const { theme, x0, y0, w, mathPrep } = ctx;
  const ts = theme.typeScale;
  const bodyPt = ts.bodyPt;
  const icPt = ts.inlineCodePt;
  const lineH = (bodyPt / 72) * LH.body;
  const merged = mergeAdjacentText(inlines);
  const hasMath = merged.some((x) => x.kind === 'math-inline');
  if (!hasMath) {
    const runs = [];
    for (const x of merged) {
      if (x.kind === 'text' && x.text)
        runs.push({
          text: x.text,
          options: {
            fontSize: bodyPt,
            fontFace: theme.fonts.body,
            color: hexNoHash(theme.text.bodyColor),
          },
        });
      else if (x.kind === 'strong')
        runs.push({
          text: x.text,
          options: {
            fontSize: bodyPt,
            fontFace: theme.fonts.body,
            color: hexNoHash(theme.text.bodyColor),
            bold: true,
          },
        });
      else if (x.kind === 'em')
        runs.push({
          text: x.text,
          options: {
            fontSize: bodyPt,
            fontFace: theme.fonts.body,
            color: hexNoHash(theme.text.bodyColor),
            italic: true,
          },
        });
      else if (x.kind === 'code')
        runs.push({
          text: x.text,
          options: {
            fontSize: icPt,
            fontFace: theme.fonts.mono,
            color: hexNoHash(theme.text.inlineCodeColor),
            highlight: hexNoHash(theme.blocks.inlineCode.fill),
          },
        });
      else if (x.kind === 'link')
        runs.push({
          text: x.text,
          options: {
            fontSize: bodyPt,
            fontFace: theme.fonts.body,
            color: hexNoHash(theme.text.linkColor),
            hyperlink: { url: x.url },
          },
        });
    }
    const t = merged.map((x) => inlinePlainText([x])).join('');
    const lines = estimateWrappedLines(t, bodyPt, w);
    const h = lineH * lines + 0.02;
    slide.addText(runs.length ? runs : [{ text: ' ', options: {} }], {
      x: x0,
      y: y0,
      w,
      h,
      valign: 'top',
      align: 'left',
    });
    return h;
  }

  let y = y0;
  let lineItems = [];
  let x = x0;
  let rowH = lineH;

  async function flushLine() {
    if (!lineItems.length) return;
    for (const it of lineItems) {
      if (it.kind === 'text') {
        slide.addText(it.text, {
          x: it.x,
          y,
          w: it.w,
          h: rowH,
          fontSize: it.fontSize,
          fontFace: it.fontFace,
          color: it.color,
          bold: it.bold,
          italic: it.italic,
          highlight: it.highlight,
          hyperlink: it.hyperlink,
          valign: 'top',
          align: 'left',
        });
      } else if (it.kind === 'img') {
        slide.addImage({ path: it.path, x: it.x, y: y + (rowH - it.h) * 0.15, w: it.w, h: it.h });
      }
    }
    y += rowH;
    lineItems = [];
    x = x0;
    rowH = lineH;
  }

  for (const part of merged) {
    if (part.kind === 'math-inline') {
      const m = await mathPrep(part.tex, false);
      let imW = m.wIn;
      let imH = m.hIn;
      if (x + imW > x0 + w && x > x0) {
        await flushLine();
      }
      lineItems.push({ kind: 'img', path: m.path, x, w: imW, h: imH });
      x += imW + 0.06;
      rowH = Math.max(rowH, imH);
    } else {
      let txt = '';
      let fontSize = bodyPt;
      let bold = false;
      let italic = false;
      let fontFace = theme.fonts.body;
      let color = hexNoHash(theme.text.bodyColor);
      let highlight;
      let hyperlink;
      if (part.kind === 'text') txt = part.text;
      else if (part.kind === 'strong') {
        txt = part.text;
        bold = true;
      } else if (part.kind === 'em') {
        txt = part.text;
        italic = true;
      } else if (part.kind === 'code') {
        txt = part.text;
        fontSize = icPt;
        fontFace = theme.fonts.mono;
        color = hexNoHash(theme.text.inlineCodeColor);
        highlight = hexNoHash(theme.blocks.inlineCode.fill);
      } else if (part.kind === 'link') {
        txt = part.text;
        color = hexNoHash(theme.text.linkColor);
        hyperlink = { url: part.url };
      }
      if (!txt) continue;
      const cw = part.kind === 'code' ? monoWIn(icPt) : charWIn(bodyPt);
      const words = txt.split(/(\s+)/);
      for (const word of words) {
        if (!word) continue;
        const pieceW = [...word].length * cw;
        if (x + pieceW > x0 + w && x > x0) {
          await flushLine();
        }
        lineItems.push({
          kind: 'text',
          text: word,
          x,
          w: Math.min(pieceW + 0.01, w),
          fontSize,
          fontFace,
          color,
          bold,
          italic,
          highlight,
          hyperlink,
        });
        x += pieceW;
        rowH = Math.max(rowH, lineH);
      }
    }
  }
  await flushLine();
  return y - y0;
}

async function renderBlock(slide, block, ctx) {
  const { theme, x0, y0, w, cacheDir, offline, themeDir, mathPrep, pptx } = ctx;
  const ts = theme.typeScale;
  const ly = theme.layout;

  if (block.kind === 'paragraph') {
    return await renderParagraphMixed(slide, block.inlines, { ...ctx, y0 });
  }
  if (block.kind === 'subheading') {
    const subPt = ts.subheadingPt;
    const t = inlinePlainText(mergeAdjacentText(block.inlines));
    const lines = estimateWrappedLines(t, subPt, w);
    const h = (subPt / 72) * LH.sub * lines;
    const runs = [];
    for (const x of mergeAdjacentText(block.inlines)) {
      if (x.kind === 'text' && x.text)
        runs.push({
          text: x.text,
          options: { fontSize: subPt, fontFace: theme.fonts.body, color: hexNoHash(theme.text.mutedColor) },
        });
      else if (x.kind === 'strong')
        runs.push({
          text: x.text,
          options: {
            fontSize: subPt,
            fontFace: theme.fonts.body,
            color: hexNoHash(theme.text.mutedColor),
            bold: true,
          },
        });
      else if (x.kind === 'em')
        runs.push({
          text: x.text,
          options: {
            fontSize: subPt,
            fontFace: theme.fonts.body,
            color: hexNoHash(theme.text.mutedColor),
            italic: true,
          },
        });
      else if (x.kind === 'code')
        runs.push({
          text: x.text,
          options: {
            fontSize: ts.inlineCodePt,
            fontFace: theme.fonts.mono,
            color: hexNoHash(theme.text.inlineCodeColor),
            highlight: hexNoHash(theme.blocks.inlineCode.fill),
          },
        });
    }
    slide.addText(runs.length ? runs : [{ text: ' ', options: {} }], {
      x: x0,
      y: y0,
      w,
      h,
      valign: 'top',
      align: 'left',
      bold: true,
    });
    return h;
  }
  if (block.kind === 'bullets') {
    const bulletPt = ts.bulletPt;
    let y = y0;
    for (const item of block.items) {
      const merged = mergeAdjacentText(item);
      const hasMath = merged.some((x) => x.kind === 'math-inline');
      const t = inlinePlainText(merged);
      const lines = estimateWrappedLines(t, bulletPt, w - 0.35);
      const h = hasMath
        ? (await renderParagraphMixed(slide, merged, { ...ctx, y0: y, w: w - 0.35, x0: x0 + 0.35 })) + ly.bulletGap
        : (bulletPt / 72) * LH.body * lines + ly.bulletGap;
      if (!hasMath) {
        const runs = [];
        for (const x of merged) {
          if (x.kind === 'text' && x.text)
            runs.push({
              text: x.text,
              options: { fontSize: bulletPt, fontFace: theme.fonts.body, color: hexNoHash(theme.text.bodyColor) },
            });
          else if (x.kind === 'strong')
            runs.push({
              text: x.text,
              options: {
                fontSize: bulletPt,
                fontFace: theme.fonts.body,
                color: hexNoHash(theme.text.bodyColor),
                bold: true,
              },
            });
          else if (x.kind === 'em')
            runs.push({
              text: x.text,
              options: {
                fontSize: bulletPt,
                fontFace: theme.fonts.body,
                color: hexNoHash(theme.text.bodyColor),
                italic: true,
              },
            });
          else if (x.kind === 'code')
            runs.push({
              text: x.text,
              options: {
                fontSize: ts.inlineCodePt,
                fontFace: theme.fonts.mono,
                color: hexNoHash(theme.text.inlineCodeColor),
                highlight: hexNoHash(theme.blocks.inlineCode.fill),
              },
            });
        }
        slide.addText(runs.length ? runs : [{ text: ' ', options: {} }], {
          x: x0 + 0.35,
          y,
          w: w - 0.35,
          h,
          valign: 'top',
          bullet: true,
          align: 'left',
        });
      } else {
        slide.addShape(pptx.ShapeType.ellipse, {
          x: x0 + 0.08,
          y: y + (bulletPt / 72) * 0.35,
          w: 0.08,
          h: 0.08,
          fill: { color: hexNoHash(theme.text.bodyColor) },
          line: { width: 0 },
        });
      }
      y += h;
    }
    return y - y0;
  }
  if (block.kind === 'code-block') {
    const codePt = ts.codeBlockPt;
    const lines = block.text.split('\n').length;
    const lineH = (codePt / 72) * LH.code;
    const pad = ly.codePadding;
    const innerH = lineH * lines;
    const boxH = innerH + pad * 2;
    slide.addShape(pptx.ShapeType.roundRect, {
      x: x0,
      y: y0,
      w,
      h: boxH,
      fill: { color: hexNoHash(theme.blocks.code.fill) },
      line: { color: hexNoHash(theme.blocks.code.border), pt: theme.blocks.code.borderWidth },
      rectRadius: theme.blocks.code.radius,
    });
    slide.addText(block.text, {
      x: x0 + pad,
      y: y0 + pad,
      w: w - pad * 2,
      h: innerH + 0.05,
      fontSize: codePt,
      fontFace: theme.fonts.mono,
      color: hexNoHash(theme.text.bodyColor),
      valign: 'top',
      align: 'left',
    });
    return boxH;
  }
  if (block.kind === 'math-block') {
    const m = await mathPrep(block.tex, true);
    const imW = Math.min(m.wIn, w * 0.92);
    const ratio = imW / m.wIn;
    const imH = m.hIn * ratio;
    const ix = x0 + (w - imW) / 2;
    slide.addImage({ path: m.path, x: ix, y: y0, w: imW, h: imH });
    return imH + ly.mathBlockGap;
  }
  if (block.kind === 'callout') {
    const pad = ly.calloutPadding;
    const iconSize = ly.calloutIconSize;
    const iconGap = ly.calloutIconGap;
    let iconPath;
    if (block.icon) {
      iconPath = await resolveAssetSrc(block.icon.src, { cacheDir, offline, themeDir });
    }
    const innerW = w - pad * 2 - (iconPath ? iconSize + iconGap : 0);
    const innerCtxBase = { ...ctx, x0: x0 + pad + (iconPath ? iconSize + iconGap : 0), w: innerW };

    let innerSum = 0;
    for (const b of block.blocks) {
      innerSum += (await renderBlock(NOOP_SLIDE, b, { ...innerCtxBase, y0: 0 })) + ly.calloutGap * 0.15;
    }
    if (block.blocks.length) innerSum -= ly.calloutGap * 0.15;

    const boxH = Math.max(innerSum + pad * 2, (iconPath ? iconSize : 0) + pad * 2);
    slide.addShape(pptx.ShapeType.roundRect, {
      x: x0,
      y: y0,
      w,
      h: boxH,
      fill: { color: hexNoHash(theme.blocks.callout.fill) },
      line: { color: hexNoHash(theme.blocks.callout.border), pt: theme.blocks.callout.borderWidth },
      rectRadius: theme.blocks.callout.radius,
    });
    let ix = x0 + pad;
    if (iconPath) {
      slide.addImage({ path: iconPath, x: ix, y: y0 + pad, w: iconSize, h: iconSize, sizing: { type: 'contain', w: iconSize, h: iconSize } });
      ix += iconSize + iconGap;
    }
    let cy = y0 + pad;
    for (const b of block.blocks) {
      cy += await renderBlock(slide, b, { ...ctx, x0: ix, y0: cy, w: innerW });
      cy += ly.calloutGap * 0.15;
    }
    return boxH;
  }
  return 0;
}

function renderTitleSlide(pptx, title, theme, opts) {
  const SLIDE_W = 13.333;
  const SLIDE_H = 7.5;
  const ly = theme.layout;
  const ts = theme.typeScale;
  const slide = pptx.addSlide();
  slide.background = { color: hexNoHash(theme.slide.background) };
  const align = ly.titleSlideAlign === 'center' ? 'center' : 'left';
  const titleReserve = opts.logoLayout?.titleReserve ?? 0;
  const x = align === 'center' ? SLIDE_W * 0.1 : ly.marginLeft;
  const w = (align === 'center' ? SLIDE_W * 0.8 : SLIDE_W - ly.marginLeft - ly.marginRight) - titleReserve;
  slide.addText(title, {
    x,
    y: SLIDE_H * 0.4,
    w,
    h: 1.2,
    fontSize: ts.titlePt,
    fontFace: theme.fonts.title,
    color: hexNoHash(theme.text.titleColor),
    align,
    valign: 'middle',
  });
  addSlideChrome(slide, theme, {
    deckTitle: title,
    slideNumber: 1,
    logoLayout: opts.logoLayout,
  });
}

async function writePptx(deckAst, theme, outPath, opts) {
  const pptx = new PptxGenJS();
  pptx.layout = theme.slide.layout || 'LAYOUT_WIDE';
  pptx.author = deckAst.title;
  pptx.title = deckAst.title;

  const themeDir = path.dirname(opts.themePath);
  const logoPath = resolveLogoPath(themeDir, theme, opts.themePath);
  const logoLayout = computeLogoLayout(logoPath, theme.headerFooter);
  const deckOpts = { ...opts, logoPath, logoLayout };

  renderTitleSlide(pptx, deckAst.title, theme, deckOpts);
  await renderDeck(deckAst, theme, pptx, deckOpts);

  await pptx.writeFile({ fileName: outPath });
}

// ---------------------------------------------------------------------------
// main
// ---------------------------------------------------------------------------

async function main() {
  const args = parseArgs(process.argv);
  const theme = loadTheme(args.theme);
  const raw = readMarkdown(args.input);
  const { title, slides: slideChunks } = splitSlides(raw);
  const md = new MarkdownIt({ html: true, linkify: true, typographer: false });
  const themeDir = path.dirname(args.theme);

  const deckAst = {
    title,
    slides: slideChunks.map((s) => ({
      title: s.title,
      blocks: parseSlideBody(s.body, md, themeDir, theme.emojiMap),
    })),
  };

  await writePptx(deckAst, theme, args.output, {
    cacheDir: args.cacheDir,
    offline: args.offline,
    debugLayout: args.debugLayout,
    themePath: args.theme,
  });
  console.log(`Wrote ${args.output}`);
}

main().catch((e) => {
  console.error(e.message || e);
  process.exit(1);
});
