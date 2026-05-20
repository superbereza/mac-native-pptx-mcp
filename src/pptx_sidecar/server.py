"""Sidecar MCP server — drop-in replacement for the broken handles in Anthropic's PowerPoint connector.

The bug: Anthropic's bundled PowerPoint MCP (Claude Desktop, Cowork mode) ships AppleScript
templates that don't match PowerPoint for Mac's AppleScript dictionary. `add_slide`,
`insert_image`, and `get_slide_content` fail with syntax error -2741 on every call. See
upstream issues: https://github.com/anthropics/claude-code/issues/20473 and #26385.

This server speaks correct AppleScript against PowerPoint for Mac and exposes the same
operations under different tool names so it can run as a sidecar next to the broken
upstream connector without name collisions. It also adds tools that the upstream connector
never exposed at all: a thumbnail renderer for visual feedback, a placeholder lister for
diagnostics, and addressing by `placeholder_format.idx` (the OOXML attribute) so that
multi-section layouts can be filled without relying on shape ordering.
"""

from __future__ import annotations

import base64
import glob
import logging
import os
import subprocess
import tempfile
from typing import Any

from mcp.server.fastmcp import FastMCP, Image

logger = logging.getLogger("pptx_sidecar")

mcp = FastMCP("powerpoint-by-anthropic-mac-sidecar")


# --- AppleScript runner ---------------------------------------------------

def _escape_applescript_string(s: str) -> str:
    """Escape a Python string so it can be embedded inside an AppleScript string literal."""
    return s.replace("\\", "\\\\").replace('"', '\\"')


def _run_osascript(script: str, timeout: int = 60) -> str:
    """Run an AppleScript via osascript and return stdout.

    Raises RuntimeError with stderr on failure so the MCP client sees a clear error.
    """
    logger.debug("Running AppleScript:\n%s", script)
    try:
        result = subprocess.run(
            ["osascript", "-e", script],
            capture_output=True,
            text=True,
            timeout=timeout,
            check=False,
        )
    except subprocess.TimeoutExpired as e:
        raise RuntimeError(f"osascript timed out after {timeout}s") from e
    if result.returncode != 0:
        raise RuntimeError(
            f"osascript failed (exit {result.returncode}): {result.stderr.strip() or result.stdout.strip()}"
        )
    return result.stdout.strip()


# --- Tools: the 4 broken upstream handles, fixed --------------------------

@mcp.tool()
def sidecar_add_slide(layout_index: int = 7, position: int | None = None) -> dict[str, Any]:
    """Add a new slide to the active presentation using a layout from the slide master.

    Use this instead of upstream `add_slide` — that one is broken (#20473).

    Upstream bug: Anthropic's `add_slide` passes `slide layout:slide layout blank` as a
    property record, but `slide layout` is not an enum constant — it's a reference to a
    layout object on a slide master. PowerPoint's parser rejects it with syntax error -2741.

    This handle creates the slide first, then assigns the layout via
    `set slide layout of newSlide to slide layout N of slide master of activePres`.

    Args:
        layout_index: 1-based index of the layout in the slide master.
        position: 1-based index to insert at. If omitted, the slide is appended at the end.

    Returns:
        dict with `slide_index` of the new slide.
    """
    if position is None:
        insertion = "make new slide at end of slides of activePres"
    else:
        insertion = f"make new slide at before slide {int(position)} of activePres"

    script = f'''
tell application "Microsoft PowerPoint"
    set activePres to active presentation
    set newSlide to {insertion}
    set slide layout of newSlide to slide layout {int(layout_index)} of slide master of activePres
    return slide index of newSlide
end tell
'''
    out = _run_osascript(script)
    return {"slide_index": int(out) if out.isdigit() else out}


@mcp.tool()
def sidecar_add_slide_from_template(source_slide_index: int, position: int | None = None) -> dict[str, Any]:
    """Duplicate an existing slide and (optionally) move the copy to a target position.

    Recommended path for projects with a heavy corporate template — the duplicate inherits
    every style detail from the source (theme placeholders, fonts, layout overrides),
    sidestepping layout-index guesswork entirely.

    Args:
        source_slide_index: 1-based index of the slide to duplicate.
        position: 1-based target index for the duplicate. If omitted, left immediately
            after the source (PowerPoint's default behavior).

    Returns:
        dict with `slide_index` of the new slide.
    """
    # PowerPoint's `duplicate` requires a `to <location>` parameter — without it the
    # command rejects with -50 Parameter error. Default: place the copy right after
    # the source (PowerPoint's intuitive "duplicate" UX).
    if position is None:
        loc = f"to after slide {int(source_slide_index)} of activePres"
    else:
        loc = f"to before slide {int(position)} of activePres"

    script = f'''
tell application "Microsoft PowerPoint"
    set activePres to active presentation
    set newSlide to (duplicate slide {int(source_slide_index)} of activePres {loc})
    return slide index of newSlide
end tell
'''
    out = _run_osascript(script)
    return {"slide_index": int(out) if out.isdigit() else out}


@mcp.tool()
def sidecar_insert_image(
    slide_index: int,
    image_path: str,
    left: float = 0,
    top: float = 0,
    width: float = 0,
    height: float = 0,
) -> dict[str, Any]:
    """Insert an image into a slide using PowerPoint for Mac's `add picture` command.

    Use this instead of upstream `insert_image` — that one is broken (#20473).

    Upstream bug: Anthropic's `insert_image` does `make new picture with properties
    {file name:..., left position:...}`. `picture` is not a creatable class on shapes,
    and `file name`/`left position` aren't valid property keys.

    Correct dictionary form: `add picture` command with named (not record-style) parameters.

    Args:
        slide_index: 1-based index of the target slide.
        image_path: Absolute POSIX path to the image file.
        left, top: Position in points (0 → safe default).
        width, height: Size in points (0 → safe default).

    Returns:
        dict with `name` of the inserted shape.
    """
    safe_path = _escape_applescript_string(image_path)
    L = float(left) if left else 50.0
    T = float(top) if top else 50.0
    W = float(width) if width else 400.0
    H = float(height) if height else 300.0

    script = f'''
tell application "Microsoft PowerPoint"
    tell slide {int(slide_index)} of active presentation
        set newPic to add picture file name (POSIX file "{safe_path}") link to file false save with document true left {L} top {T} width {W} height {H}
        return name of newPic
    end tell
end tell
'''
    out = _run_osascript(script)
    return {"shape_name": out}


@mcp.tool()
def sidecar_get_slide_content(slide_index: int) -> dict[str, Any]:
    """Read all text from a slide's shapes.

    Use this instead of upstream `get_slide_content` — that one is broken (#20473).

    Upstream bug: Anthropic's version writes `if has text frame shp` — missing the `of`
    separator. Correct: `if (has text frame of shp) then ...`.

    Args:
        slide_index: 1-based index of the slide to read.

    Returns:
        dict with `text` (newline-joined) and `shapes` (per-shape entries).
    """
    # We avoid `AppleScript's text item delimiters` — it's unreliable inside `tell
    # application` blocks (caused -2763 in v0.2.0). Instead we concatenate with
    # explicit sentinel substrings and split in Python.
    script = f'''
tell application "Microsoft PowerPoint"
    set activePres to active presentation
    set targetSlide to slide {int(slide_index)} of activePres
    set acc to ""
    repeat with shp in shapes of targetSlide
        try
            if (has text frame of shp) then
                set shpName to ""
                try
                    set shpName to (name of shp) as text
                end try
                set shpText to ""
                try
                    set shpText to (content of text range of text frame of shp) as text
                end try
                set row to shpName & "<<F>>" & shpText
                if acc is "" then
                    set acc to row
                else
                    set acc to acc & "<<NL>>" & row
                end if
            end if
        end try
    end repeat
    return acc
end tell
'''
    out = _run_osascript(script)
    shapes = []
    text_lines = []
    if out:
        for line in out.split("<<NL>>"):
            if "<<F>>" in line:
                name, _, txt = line.partition("<<F>>")
                shapes.append({"name": name, "text": txt})
                text_lines.append(txt)
            elif line:
                shapes.append({"name": "", "text": line})
                text_lines.append(line)
    return {"text": "\n".join(text_lines), "shapes": shapes}


# --- Tools: new — visual feedback, placeholder addressing, layout ops -----

@mcp.tool()
def sidecar_get_slide_thumbnail(slide_index: int, dpi: int = 100) -> Image:
    """Render a single slide as a PNG and return it inline so the assistant can see it.

    This is the most important sidecar tool — without visual feedback the assistant is
    blind to formatting errors and has to ask a human to open PowerPoint and screenshot.

    Implementation: PowerPoint exports the active presentation to a temp PDF, then
    `pdftoppm` (poppler) extracts the requested page as PNG. We went via PDF because
    PowerPoint's `save … as save as PNG` behaviour is fragile (silently writes nothing
    when given a POSIX folder, or fragments naming across versions). PDF export is the
    same path the upstream `export_pdf` handle uses, so it's known to work.

    Requires `pdftoppm` on PATH (ships with Homebrew `poppler`).

    Args:
        slide_index: 1-based index of the slide to render.
        dpi: render resolution. 100 is a good default for inline previews.

    Returns:
        Image (PNG) wrapped in an MCP ImageContent block — visible inline to the model.
    """
    pdftoppm_path = "/opt/homebrew/bin/pdftoppm"
    if not os.path.exists(pdftoppm_path):
        # fallback to PATH lookup
        pdftoppm_path = "pdftoppm"

    tmp_dir = tempfile.mkdtemp(prefix="sidecar_thumb_")
    pdf_path = os.path.join(tmp_dir, "deck.pdf")

    script = f'''
tell application "Microsoft PowerPoint"
    set activePres to active presentation
    save activePres in "{_escape_applescript_string(pdf_path)}" as save as PDF
end tell
'''
    _run_osascript(script, timeout=120)

    if not os.path.exists(pdf_path):
        contents = []
        for root, _dirs, files in os.walk(tmp_dir):
            for f in files:
                contents.append(os.path.join(root, f))
        raise RuntimeError(
            f"PDF export produced no file at {pdf_path}. "
            f"tmp_dir contents: {contents[:20]}"
        )

    prefix = os.path.join(tmp_dir, "slide")
    result = subprocess.run(
        [pdftoppm_path,
         "-f", str(int(slide_index)), "-l", str(int(slide_index)),
         "-png", "-r", str(int(dpi)),
         pdf_path, prefix],
        capture_output=True, text=True, check=False,
    )
    if result.returncode != 0:
        raise RuntimeError(f"pdftoppm failed: {result.stderr.strip()}")

    # pdftoppm names files: prefix-N.png (zero-padded if multi-digit page range)
    candidates = (
        glob.glob(f"{prefix}-{int(slide_index)}.png")
        + glob.glob(f"{prefix}-{int(slide_index):02d}.png")
        + glob.glob(f"{prefix}-{int(slide_index):03d}.png")
    )
    if not candidates:
        raise RuntimeError(
            f"pdftoppm produced no PNG for page {slide_index}. "
            f"tmp_dir: {os.listdir(tmp_dir)}"
        )

    png_bytes = open(candidates[0], "rb").read()
    return Image(data=png_bytes, format="png")


@mcp.tool()
def sidecar_set_text_in_placeholder(
    slide_index: int, placeholder_idx: int, text: str
) -> dict[str, Any]:
    """Write text into the placeholder identified by `placeholder_format.idx`.

    Why this matters: upstream `set_slide_title` only targets the TITLE placeholder,
    and `add_text_to_slide` addresses by *shape ordering index*, which is unstable
    across edits and useless for multi-section layouts (e.g. our template's "3 плашки",
    "4 плашки", "текстовые блоки_4") where you need to dot the body of placeholder
    idx=20 or 27 specifically.

    This tool iterates the slide's placeholder shapes, reads `placeholder index of
    placeholder format` (the OOXML `<p:ph idx>` attribute), and writes into the match.

    Important: setting `content of text range` REPLACES the text but preserves the
    formatting of the first run (PowerPoint extends the run's rPr over the new text).
    If the placeholder originally had multiple paragraphs / runs with distinct styles,
    only the first run's style survives.

    Args:
        slide_index: 1-based index of the target slide.
        placeholder_idx: The OOXML `<p:ph idx>` attribute of the target placeholder
            (NOT the position-based shape index).
        text: New text content.

    Returns:
        dict with `shape_name` of the placeholder that was updated.
    """
    safe_text = _escape_applescript_string(text)
    script = f'''
tell application "Microsoft PowerPoint"
    set targetSlide to slide {int(slide_index)} of active presentation
    set matched to ""
    repeat with shp in shapes of targetSlide
        try
            if (placeholder index of placeholder format of shp) is equal to {int(placeholder_idx)} then
                set content of text range of text frame of shp to "{safe_text}"
                set matched to (name of shp) as text
                exit repeat
            end if
        on error errMsg
            -- shp is not a placeholder; skip
        end try
    end repeat
    if matched is "" then
        error "No placeholder with placeholder_format.idx={int(placeholder_idx)} on slide {int(slide_index)}"
    end if
    return matched
end tell
'''
    out = _run_osascript(script)
    return {"shape_name": out}


@mcp.tool()
def sidecar_move_slide(from_index: int, to_index: int) -> dict[str, Any]:
    """Reorder slides natively via PowerPoint AppleScript.

    Cleaner than python-pptx manipulation of `_sldIdLst` — no zip-duplicate hazard,
    no need to re-pack the file.

    Args:
        from_index: 1-based current position of the slide to move.
        to_index: 1-based target position after the move.

    Returns:
        dict with `new_index` (the slide's index after the move).
    """
    # AppleScript's `move` accepts `to before slide N` / `to after slide N` references.
    # To make the semantics intuitive ("the slide ends up at to_index"), we use
    # `to before slide to_index` for moves that go up, and `to after slide to_index`
    # for moves that go down — the location reference is interpreted against the
    # pre-move slide ordering.
    if int(to_index) <= int(from_index):
        loc = f"to before slide {int(to_index)} of activePres"
    else:
        loc = f"to after slide {int(to_index)} of activePres"

    script = f'''
tell application "Microsoft PowerPoint"
    set activePres to active presentation
    set srcSlide to slide {int(from_index)} of activePres
    move srcSlide {loc}
    return slide index of srcSlide
end tell
'''
    out = _run_osascript(script)
    return {"new_index": int(out) if out.isdigit() else out}


@mcp.tool()
def sidecar_set_slide_layout(slide_index: int, layout_index: int) -> dict[str, Any]:
    """Change the layout of an existing slide without recreating it.

    Args:
        slide_index: 1-based index of the slide.
        layout_index: 1-based index of the new layout in the slide master.

    Returns:
        dict with `slide_index` (echo, for confirmation).
    """
    script = f'''
tell application "Microsoft PowerPoint"
    set activePres to active presentation
    set targetSlide to slide {int(slide_index)} of activePres
    set slide layout of targetSlide to slide layout {int(layout_index)} of slide master of activePres
    return slide index of targetSlide
end tell
'''
    out = _run_osascript(script)
    return {"slide_index": int(out) if out.isdigit() else out}


@mcp.tool()
def sidecar_list_placeholders(slide_index: int) -> dict[str, Any]:
    """Inventory every placeholder on a slide — idx, type, geometry, current text.

    Use as a diagnostic step before calling `sidecar_set_text_in_placeholder` to find
    the right `placeholder_format.idx` for the slot you want to fill.

    Args:
        slide_index: 1-based index of the slide.

    Returns:
        dict with `placeholders` (list of {name, idx, type, left, top, width, height, text}).
    """
    # PowerPoint Mac AppleScript exposes geometry as `left`, `top`, `width`, `height`.
    # `left position` is the Windows-VBA name and doesn't exist in the Mac dictionary
    # (caused -2741 syntax error in v0.2.0).
    # Sentinel-join instead of AppleScript text item delimiters (unreliable inside tell).
    script = f'''
tell application "Microsoft PowerPoint"
    set targetSlide to slide {int(slide_index)} of active presentation
    set acc to ""
    repeat with shp in shapes of targetSlide
        try
            set phIdx to (placeholder index of placeholder format of shp) as text
            -- if shp is not a placeholder, the line above throws and we skip.
            set phType to ""
            try
                set phType to (placeholder type of placeholder format of shp) as text
            end try
            set shpName to ""
            try
                set shpName to (name of shp) as text
            end try
            set L to ""
            try
                set L to (left of shp) as text
            end try
            set T to ""
            try
                set T to (top of shp) as text
            end try
            set W to ""
            try
                set W to (width of shp) as text
            end try
            set H to ""
            try
                set H to (height of shp) as text
            end try
            set shpText to ""
            try
                set shpText to (content of text range of text frame of shp) as text
            end try
            set row to shpName & "<<F>>" & phIdx & "<<F>>" & phType & "<<F>>" & L & "<<F>>" & T & "<<F>>" & W & "<<F>>" & H & "<<F>>" & shpText
            if acc is "" then
                set acc to row
            else
                set acc to acc & "<<NL>>" & row
            end if
        end try
    end repeat
    return acc
end tell
'''
    out = _run_osascript(script)
    placeholders = []
    if out:
        for line in out.split("<<NL>>"):
            parts = line.split("<<F>>", 7)
            if len(parts) != 8:
                continue
            name, idx, ptype, L, T, W, H, text = parts
            try:
                placeholders.append({
                    "name": name,
                    "idx": int(idx),
                    "type": ptype,
                    "left": float(L),
                    "top": float(T),
                    "width": float(W),
                    "height": float(H),
                    "text": text,
                })
            except ValueError:
                placeholders.append({
                    "name": name, "idx": idx, "type": ptype,
                    "left": L, "top": T, "width": W, "height": H, "text": text,
                })
    return {"placeholders": placeholders}


@mcp.tool()
def sidecar_replace_text_in_shape(
    slide_index: int, shape_index: int, old: str, new: str
) -> dict[str, Any]:
    """Replace a substring inside one shape's text, preserving run-level formatting.

    Why not `set text frame.text = ...`: that path collapses every run on the text frame
    into one and resets character properties to the placeholder default. A native
    AppleScript `find/replace` on the text range edits in place, so styled spans
    (bold words, accent-colored phrases) survive.

    Args:
        slide_index: 1-based index of the slide.
        shape_index: 1-based shape index on the slide. Use `sidecar_list_placeholders`
            or upstream's working calls first to find the right index.
        old: Substring to search for (exact match, case-sensitive).
        new: Replacement string.

    Returns:
        dict with `found` (True if at least one replacement happened) and `shape_name`.
    """
    safe_old = _escape_applescript_string(old)
    safe_new = _escape_applescript_string(new)
    script = f'''
tell application "Microsoft PowerPoint"
    set targetSlide to slide {int(slide_index)} of active presentation
    set targetShape to shape {int(shape_index)} of targetSlide
    set foundFlag to false
    try
        set tr to text range of text frame of targetShape
        -- `replace` returns the modified range; we treat non-error as success.
        replace tr what "{safe_old}" replacement "{safe_new}"
        set foundFlag to true
    on error errMsg
        error "replace failed: " & errMsg
    end try
    return (foundFlag as text) & "||" & (name of targetShape)
end tell
'''
    out = _run_osascript(script)
    found_flag, _, shape_name = out.partition("||")
    return {"found": found_flag.strip() == "true", "shape_name": shape_name.strip()}


# --- Entry point ----------------------------------------------------------

def main() -> None:
    """Run the MCP server over stdio (default transport for Claude Desktop)."""
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(name)s %(message)s")
    mcp.run()


if __name__ == "__main__":
    main()
