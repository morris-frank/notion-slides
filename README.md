<img src="brand/icon/icon-notion-slides-on-obsidian-512.png" align="left" width="128" hspace="16" alt="notion-slides icon">

<h3>notion-slides</h3>

<p>
  <sub>PASTE FROM NOTION, GET A DECK</sub>
  <br>
  <strong>Convert Notion-style Markdown into a themed <code>.pptx</code> deck: titles, bullets, code, math, callouts and images.</strong>
  <br>
  <br>
  <a href="https://www.npmjs.com/package/notion-slides"><img src="https://img.shields.io/npm/v/notion-slides?style=flat-square&amp;color=CB3837&amp;logo=npm&amp;labelColor=2D2825" alt="npm version"></a>
  <img src="https://img.shields.io/badge/node-%E2%89%A518-D78A7A?style=flat-square&amp;labelColor=2D2825" alt="Node 18+">
  <img src="https://img.shields.io/badge/output-.pptx-D78A7A?style=flat-square&amp;labelColor=2D2825" alt="pptx output">
  <a href="theme.light.json"><img src="https://img.shields.io/badge/themes-light%20%2B%20dark-D78A7A?style=flat-square&amp;labelColor=2D2825" alt="Light and dark themes"></a>
</p>

<br clear="left">

Ships with unbranded light and dark themes (`theme.light.json`, `theme.dark.json`).

**Requirements:** [Node.js](https://nodejs.org/) 18+ (for `fetch`).

## Install

```bash
npm install -g notion-slides
```

Or run without installing: `npx notion-slides deck.md --theme light --out deck.pptx`.

## Usage (quick)

```bash
notion-slides input.md --theme light --out output.pptx
```

`--theme` takes `light`, `dark` (bundled, unbranded) or a path to your own theme JSON; a custom theme picks up `logo.light.png` / `logo.dark.png` (or `logo.png`) from its own directory unless the theme sets `"headerFooter": { "logo": false }`. Other flags: `--out`, `--cache <dir>`, `--offline`, `--debug-layout`.

## Working with Notion

- **Markdown export:** You can feed Notion’s exported Markdown into this tool, but **callouts and other rich blocks are usually lost** in that export.
- **Richer content:** In Notion, open the page, **select all blocks** (or the content you need), **copy**, and **paste into a `.md` file**. That tends to preserve more structure (including callout-style HTML the converter understands). Pasting usually **does not bring over the page title the same way as a single top-level heading**, and the **first headline level can differ** from what you want for the deck.
- **Deck title:** Add the **document / deck title yourself** as the **first line of the file** using a single top-level heading, for example:

  ```markdown
  # My deck title

  …rest of pasted content…
  ```

  The first `# …` line becomes the title slide only; body slides are split on `---` between them.

## Repo layout

| Path | Role |
|------|------|
| `notion-slides.mjs` | CLI |
| `theme.light.json` / `theme.dark.json` | Layout, colors, header/footer copy |
| `examples/` | Sample Markdown |

