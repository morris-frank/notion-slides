<img src="brand/icon/icon-notion-slides-on-obsidian-512.png" align="left" width="128" hspace="16" alt="notion-slides icon">

<h3>notion-slides</h3>

<p>
  <sub>PASTE FROM NOTION, GET A DECK</sub>
  <br>
  <strong>Convert Notion-style Markdown into a themed <code>.pptx</code> deck: titles, bullets, code, math, callouts and images.</strong>
  <br>
  <br>
  <img src="https://img.shields.io/badge/node-%E2%89%A518-8EDE3D?style=flat-square&amp;labelColor=16211B" alt="Node 18+">
  <img src="https://img.shields.io/badge/output-.pptx-8EDE3D?style=flat-square&amp;labelColor=16211B" alt="pptx output">
  <a href="theme.light.json"><img src="https://img.shields.io/badge/themes-light%20%2B%20dark-1AB172?style=flat-square&amp;labelColor=16211B" alt="Light and dark themes"></a>
</p>

<br clear="left">

Ships with Soilytix light/dark themes (`theme.light.json`, `theme.dark.json`) and matching `logo.light.png` / `logo.dark.png` in the repo root.

**Requirements:** [Node.js](https://nodejs.org/) 18+ (for `fetch`).

## One-shot install (curl + PATH)

This downloads the latest `main` tree, installs npm dependencies, and links the CLI into `~/.local/bin` (adjust `INSTALL_DIR` / `BIN_DIR` if you prefer).

```bash
INSTALL_DIR="${INSTALL_DIR:-$HOME/.local/share/notion-slides}" \
BIN_DIR="${BIN_DIR:-$HOME/.local/bin}" \
REPO_TGZ="https://github.com/morris-frank/notion-slides/archive/refs/heads/main.tar.gz" \
&& TMP="$(mktemp -d)" && trap 'rm -rf "$TMP"' EXIT \
&& curl -fsSL "$REPO_TGZ" | tar -xz -C "$TMP" \
&& rm -rf "$INSTALL_DIR" \
&& mv "$TMP/notion-slides-main" "$INSTALL_DIR" \
&& (cd "$INSTALL_DIR" && npm ci --omit=dev) \
&& mkdir -p "$BIN_DIR" \
&& ln -sf "$INSTALL_DIR/notion-slides.mjs" "$BIN_DIR/notion-slides" \
&& chmod +x "$INSTALL_DIR/notion-slides.mjs" \
&& echo "Installed to $INSTALL_DIR — ensure PATH includes $BIN_DIR, e.g. export PATH=\"$BIN_DIR:\$PATH\""
```

Themes and logos live next to the CLI in `INSTALL_DIR`; pass **`--theme`** as an absolute path, or run from that directory:

```bash
export PATH="$HOME/.local/bin:$PATH"
notion-slides deck.md \
  --theme "$HOME/.local/share/notion-slides/theme.light.json" \
  --out deck.pptx
```

For the dark theme, use `theme.dark.json` instead.

## Usage (quick)

```bash
node notion-slides.mjs input.md --theme theme.light.json --out output.pptx
```

Flags include `--theme`, `--out`, `--cache <dir>`, `--offline`, `--debug-layout`.

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
| `logo.light.png` / `logo.dark.png` | Header logo (theme picks by light/dark) |
| `examples/` | Sample Markdown |

