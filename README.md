# powerpoint-by-anthropic-mac-sidecar

A sidecar MCP server that fixes four broken handles in Anthropic's first-party PowerPoint MCP connector for Claude Desktop on macOS.

## Why this exists

Claude Desktop on macOS ships with a first-party **"PowerPoint By Anthropic"** MCP connector (part of the Cowork integration suite). It drives Microsoft PowerPoint for Mac through AppleScript. As of mid-2026, four of its handles fail with AppleScript syntax error `-2741` on every call:

- `add_slide`
- `add_slide_from_template` (effectively — no working way to clone an existing slide)
- `insert_image`
- `get_slide_content`

Upstream bug reports:

- [anthropics/claude-code#20473](https://github.com/anthropics/claude-code/issues/20473) — *PowerPoint By Anthropic MCP Connector*
- [anthropics/claude-code#26385](https://github.com/anthropics/claude-code/issues/26385) — *Cowork: Word & PowerPoint MCP connectors have AppleScript syntax errors*

Both have been open for months without a fix. The bundled server lives inside Claude Desktop's embedded Bun.js filesystem (`/$bunfs/root/claude`), so end users cannot patch it.

## Root cause

The four broken handles are not random typos — they share one structural mistake: Anthropic's AppleScript templates look like Microsoft PowerPoint **VBA on Windows** transliterated into AppleScript, but **PowerPoint for Mac's AppleScript dictionary is shaped differently**. Specifically:

| Broken upstream | Why it fails | Correct dictionary form |
|---|---|---|
| `make new slide … with properties {slide layout:slide layout blank}` | `slide layout` is not an enum constant — it's a reference to a layout object on a slide master. | Create the slide first, then `set slide layout of newSlide to slide layout N of slide master of activePres`. |
| `make new picture with properties {file name:…, left position:…}` | `picture` is not a creatable class on `shapes`; `file name` / `left position` are not property names. | Use the `add picture` command with named (not record-style) parameters. |
| `if has text frame shp then …` | Missing `of` separator — `has text frame` is a boolean property of a shape, not a free function. | `if (has text frame of shp) then …` |
| (no clone primitive at all) | — | Use `duplicate slide N of activePres` and then move with `move newSlide to …`. |

This sidecar re-implements those four handles using the dictionary-correct syntax, runs them through `osascript`, and exposes them as MCP tools so Claude can call them directly.

## What this is **not**

- Not a fork or patch of the upstream connector — that connector is closed-source and bundled into Claude Desktop. This server runs **next to** it under different tool names so the two coexist without collisions.
- Not a cross-platform PowerPoint MCP. It is macOS-only by design — every operation goes through AppleScript against the live Microsoft PowerPoint for Mac process. If you want a portable XML-injection server, use [python-pptx](https://python-pptx.readthedocs.io/) or one of the [python-pptx-based MCP servers](https://github.com/GongRzhe/Office-PowerPoint-MCP-Server).
- Not a full PowerPoint controller. The upstream connector's working handles (`create_presentation`, `open_presentation`, `save_presentation`, `close_presentation`, `delete_slide`, `set_slide_title`, `add_text_to_slide`) are left alone — use the upstream connector for those.

## Why AppleScript and not python-pptx

When you edit a `.pptx` file with `python-pptx`, you are writing OOXML by hand. The file is closed and re-saved; PowerPoint's runtime theme inheritance is not involved. In practice this loses or distorts slide-level formatting that depends on master/layout cascades, theme color references, or live shape geometry — the sort of issues that show up as "fonts look wrong" or "the corporate template doesn't apply".

AppleScript drives the **running** PowerPoint application. Theme inheritance, layout cascades, and master overrides are evaluated by PowerPoint itself, exactly as if a human were clicking through the UI. For workflows that rely on a specific corporate template, this is the only reliable path on macOS today.

## Tools

All tools are prefixed with `sidecar_` to avoid clashing with the upstream connector.

### `sidecar_add_slide(layout_index: int = 7, position: int | None = None)`
Append (or insert) a new slide and assign a layout from the active presentation's slide master. `layout_index` is 1-based into the slide master's layouts collection. The default `7` is "Blank" in PowerPoint for Mac's stock theme; templates may number layouts differently.

### `sidecar_add_slide_from_template(source_slide_index: int, position: int | None = None)`
Duplicate an existing slide. The new slide inherits every style detail from the source slide (theme placeholders, fonts, custom layout overrides). This is the recommended path when you have a template deck and want to add a new slide that looks identical to slide N.

### `sidecar_insert_image(slide_index: int, image_path: str, left: float = 0, top: float = 0, width: float = 0, height: float = 0)`
Insert an image into a slide. Coordinates are in points (PowerPoint's native unit). Pass `0` for `left`/`top`/`width`/`height` to fall back to safe defaults (50, 50, 400, 300).

### `sidecar_get_slide_content(slide_index: int) -> {"text": str, "shapes": [{"name": str, "text": str}]}`
Read all text from shapes on a slide that have a text frame. Returns both a joined `text` string and a structured per-shape list.

## Installation

Requires:
- macOS with Microsoft PowerPoint for Mac installed and running
- Python 3.10+
- Claude Desktop (or any MCP client that supports stdio)

```bash
git clone <this-repo>
cd powerpoint-by-anthropic-mac-sidecar
python3 -m venv .venv
source .venv/bin/activate
pip install -e .
```

The console script `pptx-sidecar` is now on your `PATH` (inside the venv).

## Wiring into Claude Desktop

Edit `~/Library/Application Support/Claude/claude_desktop_config.json` and add an entry to `mcpServers`:

```json
{
  "mcpServers": {
    "powerpoint-by-anthropic-mac-sidecar": {
      "command": "/absolute/path/to/.venv/bin/pptx-sidecar"
    }
  }
}
```

Restart Claude Desktop. The four `sidecar_*` tools will appear alongside the upstream PowerPoint connector's tools.

## Smoke test

With PowerPoint open and a presentation active:

```bash
osascript -e 'tell application "Microsoft PowerPoint" to return name of active presentation'
```

If that prints the file name, AppleScript can talk to PowerPoint and the sidecar will work. If it prints nothing or errors, fix that first — every sidecar tool depends on the same channel.

## License

MIT. See [LICENSE](LICENSE).
