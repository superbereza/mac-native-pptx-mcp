# mac-native-pptx-mcp

A full-featured PowerPoint MCP server for macOS — drives Microsoft PowerPoint for Mac through AppleScript with the dictionary-correct syntax that Anthropic's first-party connector got wrong.

Originally born as a sidecar to fix four broken handles in the first-party "PowerPoint By Anthropic" connector ([#20473](https://github.com/anthropics/claude-code/issues/20473), [#26385](https://github.com/anthropics/claude-code/issues/26385)). Now stands alone: 21 tools covering the full lifecycle (create / open / save / close / export PDF), slide editing (add / clone / delete / move / set layout), shape editing (insert image, set text, replace text, set geometry, delete), and visual feedback (single thumbnail + grid overview). You don't need the upstream Anthropic connector at all.

Live-verified against PowerPoint for Mac 16.x against a real corporate-template deck.

## Why this exists

Claude Desktop on macOS ships with a first-party **"PowerPoint By Anthropic"** MCP connector (part of the Cowork integration suite). It drives Microsoft PowerPoint for Mac through AppleScript. As of mid-2026, several of its handles fail with AppleScript error `-2741` (syntax) or `-50` (parameter) on every call:

- `add_slide`
- `insert_image`
- `get_slide_content`
- `set_slide_title` / `add_text_to_slide` (work only when a TITLE placeholder is already on the slide)

Upstream bug reports:

- [anthropics/claude-code#20473](https://github.com/anthropics/claude-code/issues/20473) — *PowerPoint By Anthropic MCP Connector*
- [anthropics/claude-code#26385](https://github.com/anthropics/claude-code/issues/26385) — *Cowork: Word & PowerPoint MCP connectors have AppleScript syntax errors*

Both have been open for months without a fix. The bundled server lives inside Claude Desktop's embedded Bun.js filesystem (`/$bunfs/root/claude`), so end users cannot patch it.

## Root cause and live-verified findings

The upstream errors are not random typos. Anthropic's AppleScript templates look like Microsoft PowerPoint **VBA on Windows** transliterated into AppleScript, but **PowerPoint for Mac's AppleScript dictionary is shaped differently**. We iterated each fragment directly via `osascript` against a live 62-slide deck and discovered:

| Pattern that fails | Why | Working form |
|---|---|---|
| `repeat with X in shapes of slide` | Hangs PowerPoint indefinitely — lazy-eval of `X` keeps the process pegged. | `repeat with i from 1 to N` + `set X to shape i of slide`. |
| `left of shape` | Wrong property name (it's the Windows-VBA spelling, doesn't exist on Mac). | `left position of shape`. |
| `index of placeholder format of shp`, `type of placeholder format of shp` | `index` and `type` are reserved AppleScript keywords, so the parser reads them as boolean-property tokens. | Pipe-quote: `(|placeholder index| of placeholder format of shp)`. |
| `set row to ...` | `row` is a reserved class name (table row). | Use any other identifier (`set lineStr to ...`). |
| `make new slide at end of slides of p` | "Can't make class slide" — `slides` collection on presentation does not accept `make`. | `make new slide at end of p` (no `of slides`). |
| `set slide layout of slide to slide layout N of slide master` | "Invalid key form -10002" — `slide layout N of master` is not a valid reference; that property doesn't exist on master. | Set the slide property `layout` to an enum constant (`set layout of slide to slide layout blank`) for built-in enums; or `set custom layout of slide N to (custom layout of slide M)` to clone another slide's layout reference. |
| `set layout of slide to slide layout X` (built-in enum) | Works, but only flips the metadata flag — shapes are NOT restructured to match the layout. Not exposed as a tool here to avoid misleading callers. | For visual layout change, use `pptx_set_slide_layout_from_template` or recreate the slide. |
| `duplicate slide N of activePres` (all variants tried) | Returns Parameter error -50; Mac AppleScript does not implement `duplicate` for slides. | `tell active presentation` + `copy object slide N` + `paste object` (rides PowerPoint's clipboard). True visual clone, all freeform shapes preserved. |
| `add picture file name ...` | The command does not exist in the Mac dictionary — `picture` is a read-only class. | `make new shape` (rectangle of target geometry) + `user picture <shape> picture file "<path>"` — sets the rectangle's fill to the image. Visually identical to a "real" picture shape. |
| `replace tr what ... replacement ...` | No native find/replace verb on text range — syntax error. | Read content of text range, do `str.replace` in Python, write content back. Loses styled runs. |
| `if has text frame shp then` | Missing `of`; parser chokes. | `if (has text frame of shp) then`. |
| `AppleScript's text item delimiters` inside `tell application` | Returns -2763 ("no result returned") unreliably. | Concatenate with sentinel substrings (`<<F>>`, `<<NL>>`) and split in Python. |
| Writing PDF/PNG to `/tmp` or `~/Documents` | PowerPoint is sandboxed — triggers TCC prompts or silently fails. | Write to `~/Library/Containers/com.microsoft.Powerpoint/Data/tmp/`. |

This server implements every handle using the working forms above and exposes them as MCP tools so Claude can call them directly.

## What this is **not**

- Not a fork or patch of Anthropic's connector — that connector is closed-source and bundled into Claude Desktop. The bug-fix backstory above is just the reason this project exists; you don't need to keep Anthropic's connector enabled to use this server.
- Not a cross-platform PowerPoint MCP. macOS-only by design — every operation goes through AppleScript against the live Microsoft PowerPoint for Mac process. If you want a portable XML-injection server, use [python-pptx](https://python-pptx.readthedocs.io/) or one of the [python-pptx-based MCP servers](https://github.com/GongRzhe/Office-PowerPoint-MCP-Server).

## Why AppleScript and not python-pptx

When you edit a `.pptx` file with `python-pptx`, you are writing OOXML by hand. The file is closed and re-saved; PowerPoint's runtime theme inheritance is not involved. In practice this loses or distorts slide-level formatting that depends on master/layout cascades, theme color references, or live shape geometry — the sort of issues that show up as "fonts look wrong" or "the corporate template doesn't apply".

AppleScript drives the **running** PowerPoint application. Theme inheritance, layout cascades, and master overrides are evaluated by PowerPoint itself, exactly as if a human were clicking through the UI. For workflows that rely on a specific corporate template, this is the only reliable path on macOS today.

## Tools

All 21 tools are prefixed with `pptx_` so they're easy to spot in tool inventories and grep over transcripts. (Earlier releases used a `sidecar_` prefix, when the project was a sidecar to Anthropic's connector. The tag `sidecar-final` marks that pre-rename state — check it out if you need the old API.)

### Presentation lifecycle

#### `pptx_create_presentation(save_to_path: str | None = None)`
Create a new blank presentation. If `save_to_path` is given, save immediately; otherwise leave it unsaved in memory.

#### `pptx_open_presentation(pptx_path: str)`
Open a pptx file. If it's already open, brings that existing instance to front (no re-open).

#### `pptx_save_presentation(save_as_path: str | None = None)`
Save the active presentation. With `save_as_path`, behaves as Save As. Without, saves to the current path.

#### `pptx_close_presentation(save_changes: bool = False, presentation_name: str | None = None)`
Close a presentation. Defaults to closing the active one without saving — pass `presentation_name` to target a specific file by name, `save_changes=True` to preserve unsaved changes.

#### `pptx_export_pdf(pdf_path: str, presentation_name: str | None = None)`
Export to PDF. PowerPoint is sandboxed — writing inside `~/Library/Containers/com.microsoft.Powerpoint/Data/` runs silently; outside paths may trigger one-time TCC prompts.

### Slide creation

#### `pptx_add_slide(layout: str = "blank", position: int | None = None)`
Add a new slide using one of PowerPoint's **built-in** layout enums: `blank`, `title`, `title only`, `text`, `chart`, `comparison`, `section header`, ... (~30 enums total). Pass the name without the `slide layout ` prefix. For custom layouts from a corporate template, use one of the `*_from_template` tools instead — built-in enums don't map to template layouts.

#### `pptx_add_slide_from_template(source_slide_index: int, position: int | None = None)`
Full **visual clone** of an existing slide via `copy object` + `paste object`. Inherits everything: layout placeholders, fonts, theme colors, AND every freeform decorative shape the author drew on top. Default use case when you want a new slide that looks identical to an existing one and then edit text/images on top.

#### `pptx_add_blank_slide_from_template(source_slide_index: int, position: int | None = None)`
**Layout-only clone** — an empty new slide that inherits the source slide's `custom layout` reference. Only the layout's declared placeholder slots materialize on the new slide; the source's freeform decorations are NOT copied. Use when you want a clean layout shell to fill from scratch.

### Visual feedback

#### `pptx_get_slide_thumbnail(slide_index: int, dpi: int = 100)`
Render one slide as a PNG and return it inline as an MCP `ImageContent` block — visible to the model. The main bridge for visual feedback: without it the assistant is blind to formatting errors. PDF-exports the active presentation to a sandboxed temp dir, then `pdftoppm` extracts the requested page.

#### `pptx_get_deck_overview(start_slide: int = 1, per_page: int = 12, columns: int = 4, dpi: int = 60)`
Render up to `per_page` slides as a single composite PNG laid out in a `columns`-wide grid, each thumbnail labeled `#N`. Useful for whole-deck overview at a glance (e.g. a 60-slide deck → 5 calls with `start_slide=1,13,25,37,49`).

### Content I/O

#### `pptx_get_slide_content(slide_index: int)`
Read all text from shapes on a slide that have a text frame. Returns `{"text": str, "shapes": [{name, text}]}`.

#### `pptx_list_shapes(slide_index: int)`
Inventory every shape on a slide: `name`, geometry (`left`, `top`, `width`, `height` in points), current `text`, and placeholder info (`placeholder_idx`, `placeholder_type`) if the shape happens to be a layout placeholder. Diagnostic step before `set_text_in_shape_by_name` — tells you which name maps to which slot. Works on freeform shapes too (most corporate templates).

#### `pptx_set_text_in_shape_by_name(slide_index: int, shape_name: str, text: str)`
Write text into a shape addressed by its `name` (stable identifier in PowerPoint's AppleScript bridge). Caveat: replaces all text and preserves only the formatting of the first run — multi-run styled text collapses to a single style.

#### `pptx_replace_text_in_shape_by_name(slide_index: int, shape_name: str, old: str, new: str)`
Substring replacement. PowerPoint Mac has no native find/replace verb on text ranges, so this reads the current content, `str.replace`s in Python, and writes back. Same run-collapsing caveat as `set_text_in_shape_by_name`.

#### `pptx_insert_image(slide_index: int, image_path: str, left: float = 50, top: float = 50, width: float = 0, height: float = 0)`
Insert an image. Mechanism: `make new shape` (rectangle of given geometry) + `user picture <shape> picture file <path>` to set the rectangle's fill to the image. Visually identical to a "real" picture shape; in OOXML it's a rectangle with `<a:blipFill>`. Sizing policy: if `width`/`height` are 0, probes the image's pixel dimensions via `sips` and fits within an 800×600-point cap while preserving aspect ratio. First call may show a one-time TCC permission prompt for image access.

### Shape and slide manipulation

#### `pptx_delete_shape_by_name(slide_index: int, shape_name: str)`
Delete the first shape with the matching name. Pairs with `set_slide_layout_from_template` (which is additive) — call this to trim stale placeholders left over from a prior layout.

#### `pptx_set_shape_geometry(slide_index: int, shape_name: str, left, top, width, height)`
Reposition or resize a shape. Any of `left`/`top`/`width`/`height` left as `None` keeps the current value — useful when you only want to shift X or only resize width.

#### `pptx_move_slide(from_index: int, to_index: int)`
Reorder slides natively via AppleScript. Cleaner than python-pptx `_sldIdLst` manipulation — no zip-duplicate hazard, no need to re-pack the file.

#### `pptx_set_slide_layout_from_template(slide_index: int, source_slide_index: int)`
Apply the `custom layout` of another slide to a target slide. Adds the new layout's placeholders on top of existing shapes (additive — does NOT remove old shapes). For a clean visual result, follow up with `pptx_delete_shape_by_name` to trim stale shapes.

## Installation

Requires:
- macOS with Microsoft PowerPoint for Mac installed and running
- Python 3.10+
- `pdftoppm` (ships with Homebrew `poppler`)
- Claude Desktop (or any MCP client that supports stdio)

```bash
git clone https://github.com/superbereza/mac-native-pptx-mcp.git
cd mac-native-pptx-mcp
python3 -m venv .venv
source .venv/bin/activate
pip install -e .
```

The console script `pptx-mcp` is now on your `PATH` (inside the venv).

## Wiring into Claude Desktop

Edit `~/Library/Application Support/Claude/claude_desktop_config.json` and add an entry to `mcpServers`:

```json
{
  "mcpServers": {
    "mac-native-pptx-mcp": {
      "command": "/absolute/path/to/.venv/bin/pptx-mcp"
    }
  }
}
```

(The console script is still named `pptx-mcp` for backward compatibility with existing Claude Desktop configs — the binary's name doesn't have to match the project name.)

Restart Claude Desktop. The `pptx_*` tools will appear alongside the upstream PowerPoint connector's tools.

## Smoke test

With PowerPoint open and a presentation active:

```bash
osascript -e 'tell application "Microsoft PowerPoint" to return name of active presentation'
```

If that prints the file name, AppleScript can talk to PowerPoint and the sidecar will work. If it prints nothing or errors, fix that first — every sidecar tool depends on the same channel.

## Debugging

Logs are written by Claude Desktop to `~/Library/Logs/Claude/mcp-server-powerpoint-by-anthropic-mac-sidecar.log`. Every JSON-RPC request and response plus any stderr from the Python process (including AppleScript fragments at DEBUG level) lands there.

```bash
tail -f ~/Library/Logs/Claude/mcp-server-powerpoint-by-anthropic-mac-sidecar.log
```

When iterating on new AppleScript fragments locally, run them directly via `osascript -e '...'` against a live PowerPoint presentation — faster feedback than reloading Claude Desktop between every change.

## License

MIT. See [LICENSE](LICENSE).
