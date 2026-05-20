"""Sidecar MCP server — drop-in replacement for the broken handles in Anthropic's PowerPoint connector.

The bug: Anthropic's bundled PowerPoint MCP (Claude Desktop, Cowork mode) ships AppleScript
templates that don't match PowerPoint for Mac's AppleScript dictionary. `add_slide`,
`insert_image`, and `get_slide_content` fail with syntax error -2741 on every call. See
upstream issues: https://github.com/anthropics/claude-code/issues/20473 and #26385.

This server speaks AppleScript that's been live-tested against PowerPoint for Mac (16.x)
against an actual presentation, so the dictionary quirks have been resolved:

  - The shapes collection is *not* safe for `repeat with X in <collection>` iteration —
    it hangs PowerPoint indefinitely. We use `repeat with i from 1 to N` everywhere.
  - The geometry property is `left position`, not `left` (despite `left` showing up in
    other PowerPoint AppleScript dialects).
  - `placeholder index` / `placeholder type` are reserved-keyword conflicts; address with
    pipe-quotes `|placeholder index|`. But: in this corporate template, shapes are
    freeform (no placeholder format at all), so the tools that need OOXML idx fall back
    to addressing by shape `name`.
  - `slide layout N of slide master` is not a valid reference form. The slide property
    is `layout` (one identifier), taking enum constants like `slide layout blank`.
  - `duplicate slide N of activePres` returns Parameter error -50 unconditionally;
    PowerPoint Mac AppleScript does not support slide duplication. Use python-pptx.
  - `replace tr what ... replacement ...` is not a valid PowerPoint AppleScript verb;
    we read-modify-write in Python (which loses styled runs).
  - PowerPoint is sandboxed; writing outside `~/Library/Containers/com.microsoft.Powerpoint/`
    triggers TCC prompts and may fail silently. All sidecar I/O lives inside the
    container's `Data/tmp/` directory.
"""

from __future__ import annotations

import glob
import logging
import os
import subprocess
import tempfile
from typing import Any

from mcp.server.fastmcp import FastMCP, Image

logger = logging.getLogger("pptx_sidecar")

mcp = FastMCP("powerpoint-by-anthropic-mac-sidecar")

# Sandbox-writable directory for any artifacts (PDF exports, thumbnails) that the
# PowerPoint process needs to produce. Writing outside this directory triggers TCC
# prompts and may fail silently.
POWERPOINT_SANDBOX_TMP = os.path.expanduser(
    "~/Library/Containers/com.microsoft.Powerpoint/Data/tmp/sidecar"
)


# --- AppleScript runner ---------------------------------------------------

def _escape_applescript_string(s: str) -> str:
    """Escape a Python string so it can be embedded inside an AppleScript string literal."""
    return s.replace("\\", "\\\\").replace('"', '\\"')


def _run_osascript(script: str, timeout: int = 60) -> str:
    """Run an AppleScript via osascript and return stdout."""
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


def _sandbox_tmp_dir(prefix: str) -> str:
    os.makedirs(POWERPOINT_SANDBOX_TMP, exist_ok=True)
    return tempfile.mkdtemp(prefix=prefix, dir=POWERPOINT_SANDBOX_TMP)


# --- Tools: the broken upstream handles, fixed ----------------------------

@mcp.tool()
def sidecar_add_slide(layout: str = "blank", position: int | None = None) -> dict[str, Any]:
    """Add a new slide to the active presentation with one of PowerPoint's built-in layouts.

    Use this instead of upstream `add_slide` — that one is broken (#20473).

    Note: `layout` is a *built-in* PowerPoint enum, NOT a corporate template's custom
    layout. PowerPoint Mac AppleScript exposes ~30 standard slide layouts (blank, title,
    title only, text, chart, etc.). Setting a custom layout from a corporate template
    via AppleScript is not exposed — you'd need to use python-pptx + slide layout XML
    references for that.

    Common enum values: blank, title, title only, text, two column text, chart,
    text and chart, organization chart, table, vertical title and text,
    title and content, section header, two content, comparison.

    Args:
        layout: Built-in layout name without the `slide layout ` prefix
                (so "blank" → AppleScript `slide layout blank`).
        position: 1-based index to insert at. If omitted, the slide is appended at the end.

    Returns:
        dict with `slide_index` of the new slide and the layout that was applied.
    """
    safe_layout = layout.strip().lower()  # AppleScript enums are lowercase

    if position is None:
        insertion = "make new slide at end of p"
    else:
        insertion = f"make new slide at after slide {int(position) - 1} of p" if int(position) > 1 \
            else "make new slide at before slide 1 of p"

    script = f'''
tell application "Microsoft PowerPoint"
    set p to active presentation
    set newSlide to {insertion}
    set layout of newSlide to slide layout {safe_layout}
    return slide index of newSlide
end tell
'''
    out = _run_osascript(script)
    return {
        "slide_index": int(out) if out.isdigit() else out,
        "layout_applied": f"slide layout {safe_layout}",
    }


@mcp.tool()
def sidecar_add_slide_from_template(source_slide_index: int, position: int | None = None) -> dict[str, Any]:
    """Clone an existing slide via PowerPoint's clipboard — full visual copy.

    The new slide inherits every style detail from the source: layout placeholders,
    fonts, theme colors, AND any freeform decorative shapes the author drew on top
    (rectangles, custom icons, hand-placed text boxes — all of it).

    Why this works: PowerPoint Mac's `duplicate` verb returns -50 for slides, but the
    pair `copy object slide N` + `paste object` (both inside `tell active presentation`)
    rides PowerPoint's clipboard machinery and creates a true clone appended at the end.
    Then `move` repositions it.

    Args:
        source_slide_index: 1-based index of the slide to clone.
        position: 1-based target index for the clone. If omitted, the clone stays at
            the end of the deck.

    Returns:
        dict with `slide_index` of the new slide.
    """
    move_clause = ""
    if position is not None:
        # `paste object` appends to the end, then we move into place.
        if int(position) <= 1:
            move_clause = "move newSlide to before slide 1 of activePres"
        else:
            move_clause = f"move newSlide to before slide {int(position)} of activePres"

    script = f'''
tell application "Microsoft PowerPoint"
    set activePres to active presentation
    tell activePres
        copy object slide {int(source_slide_index)}
        paste object
    end tell
    set newSlide to slide (count of slides of activePres) of activePres
    {move_clause}
    return slide index of newSlide
end tell
'''
    out = _run_osascript(script)
    return {"slide_index": int(out) if out.isdigit() else out}


@mcp.tool()
def sidecar_insert_image(
    slide_index: int,
    image_path: str,
    left: float = 50,
    top: float = 50,
    width: float = 400,
    height: float = 300,
) -> dict[str, Any]:
    """Insert an image into a slide.

    Implementation note: PowerPoint Mac's AppleScript dictionary does NOT expose `add
    picture` (despite what the upstream connector tried to use), and the `picture` class
    is read-only — `make new picture` returns -50 Parameter error. The dictionary-correct
    path on Mac is two steps: (1) `make new shape` with a rectangle of the target
    geometry, (2) `user picture <shape> picture file "<path>"` to set its fill to the
    image. Visually identical to a "real" picture shape; in OOXML this is a rectangle
    with a `<a:blipFill>` instead of an `<p:pic>` element.

    Note: on first call PowerPoint may show a one-time TCC permission dialog asking
    whether to allow access to the source image. Approve it; future calls run silently.

    Args:
        slide_index: 1-based index of the target slide.
        image_path: Absolute POSIX path to the image file.
        left, top, width, height: Geometry in points (PowerPoint's native unit).

    Returns:
        dict with `shape_name` of the inserted rectangle (e.g. "Shape_0").
    """
    safe_path = _escape_applescript_string(image_path)
    script = f'''
tell application "Microsoft PowerPoint"
    set sl to slide {int(slide_index)} of active presentation
    set newShape to make new shape at sl with properties {{left position:{float(left)}, top:{float(top)}, width:{float(width)}, height:{float(height)}, auto shape type:autoshape rectangle}}
    user picture newShape picture file "{safe_path}"
    return name of newShape
end tell
'''
    out = _run_osascript(script)
    return {"shape_name": out}


@mcp.tool()
def sidecar_get_slide_content(slide_index: int) -> dict[str, Any]:
    """Read all text from a slide's shapes.

    Use this instead of upstream `get_slide_content` — that one is broken (#20473).

    Returns:
        dict with `text` (newline-joined) and `shapes` (per-shape entries).
    """
    script = f'''
tell application "Microsoft PowerPoint"
    set targetSlide to slide {int(slide_index)} of active presentation
    set n to count of shapes of targetSlide
    set acc to ""
    repeat with i from 1 to n
        set shp to shape i of targetSlide
        try
            if (has text frame of shp) then
                set shpName to (name of shp) as text
                set shpText to ""
                try
                    set shpText to (content of text range of text frame of shp) as text
                end try
                set lineStr to shpName & "<<F>>" & shpText
                if acc is "" then
                    set acc to lineStr
                else
                    set acc to acc & "<<NL>>" & lineStr
                end if
            end if
        end try
    end repeat
    return acc
end tell
'''
    out = _run_osascript(script)
    shapes: list[dict[str, str]] = []
    text_lines: list[str] = []
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


# --- Tools: new — visual feedback, addressing, layout ops -----------------

@mcp.tool()
def sidecar_get_slide_thumbnail(slide_index: int, dpi: int = 100) -> Image:
    """Render a single slide as a PNG and return it inline for the assistant to see.

    Implementation: PowerPoint exports the active presentation to PDF inside its
    sandboxed temp directory (`~/Library/Containers/com.microsoft.Powerpoint/Data/tmp/sidecar/`),
    then `pdftoppm` extracts the requested page as PNG. Writing outside the PowerPoint
    sandbox triggers TCC prompts and silently fails, so we always stay inside.

    Args:
        slide_index: 1-based index of the slide to render.
        dpi: Render resolution. 100 is a good default for inline previews.

    Returns:
        Image (PNG) wrapped as an MCP ImageContent block — visible inline to the model.
    """
    pdftoppm_path = "/opt/homebrew/bin/pdftoppm"
    if not os.path.exists(pdftoppm_path):
        pdftoppm_path = "pdftoppm"

    tmp_dir = _sandbox_tmp_dir("thumb_")
    pdf_path = os.path.join(tmp_dir, "deck.pdf")

    safe_pdf = _escape_applescript_string(pdf_path)
    script = f'''
tell application "Microsoft PowerPoint"
    save active presentation in (POSIX file "{safe_pdf}") as save as PDF
end tell
'''
    _run_osascript(script, timeout=180)

    if not os.path.exists(pdf_path):
        raise RuntimeError(
            f"PDF export produced no file at {pdf_path}. Check that PowerPoint has an "
            f"active presentation open and is responsive."
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

    # pdftoppm zero-pads the page number based on the total page count.
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
def sidecar_list_shapes(slide_index: int) -> dict[str, Any]:
    """Inventory every shape on a slide — name, geometry, current text, and placeholder
    info if the shape has one. Diagnostic step before `sidecar_set_text_in_shape_by_name`.

    Why this exists (and not `list_placeholders`): in heavy corporate templates, slides
    are commonly built from freeform shapes rather than layout placeholders. PowerPoint
    AppleScript reports `count of placeholders of slide` = 0 for these. So we iterate
    `shapes` (which always exist) and surface placeholder info only when present.

    Args:
        slide_index: 1-based index of the slide.

    Returns:
        dict with `shapes`: list of {name, left, top, width, height, text,
        placeholder_idx (or null), placeholder_type (or null)}.
    """
    script = f'''
tell application "Microsoft PowerPoint"
    set targetSlide to slide {int(slide_index)} of active presentation
    set n to count of shapes of targetSlide
    set acc to ""
    repeat with i from 1 to n
        set shp to shape i of targetSlide
        try
            set shpName to (name of shp) as text
            set L to ""
            try
                set L to (left position of shp) as text
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
                if (has text frame of shp) then
                    set shpText to (content of text range of text frame of shp) as text
                end if
            end try
            set phIdx to ""
            set phType to ""
            try
                set phIdx to (|placeholder index| of placeholder format of shp) as text
                try
                    set phType to (|placeholder type| of placeholder format of shp) as text
                end try
            end try
            set rowStr to shpName & "<<F>>" & L & "<<F>>" & T & "<<F>>" & W & "<<F>>" & H & "<<F>>" & phIdx & "<<F>>" & phType & "<<F>>" & shpText
            if acc is "" then
                set acc to rowStr
            else
                set acc to acc & "<<NL>>" & rowStr
            end if
        end try
    end repeat
    return acc
end tell
'''
    out = _run_osascript(script)
    shapes: list[dict[str, Any]] = []
    if out:
        for line in out.split("<<NL>>"):
            parts = line.split("<<F>>", 7)
            if len(parts) != 8:
                continue
            name, L, T, W, H, phIdx, phType, text = parts
            entry: dict[str, Any] = {
                "name": name,
                "text": text,
                "placeholder_idx": int(phIdx) if phIdx.isdigit() else None,
                "placeholder_type": phType or None,
            }
            for key, raw in (("left", L), ("top", T), ("width", W), ("height", H)):
                try:
                    entry[key] = float(raw)
                except ValueError:
                    entry[key] = raw or None
            shapes.append(entry)
    return {"shapes": shapes}


@mcp.tool()
def sidecar_set_text_in_shape_by_name(
    slide_index: int, shape_name: str, text: str
) -> dict[str, Any]:
    """Write text into a shape addressed by its `name` (the stable identifier in
    PowerPoint's AppleScript dictionary).

    Why by name and not by `placeholder_format.idx`: in this template, AppleScript
    reports zero placeholders on every slide — shapes are freeform. The `name` property
    ("Text 0", "Title 1", "Группа 14", etc.) is the only stable identifier available
    through the AppleScript bridge. Use `sidecar_list_shapes` first to find the right
    name.

    Caveat: `set content of text range` replaces the text but preserves the formatting
    of the first run only (PowerPoint extends the first run's rPr over the new content).
    Multi-run formatting is collapsed.

    Args:
        slide_index: 1-based index of the target slide.
        shape_name: Exact `name` of the shape (case-sensitive).
        text: New text content.

    Returns:
        dict with `shape_name` of the updated shape.
    """
    safe_name = _escape_applescript_string(shape_name)
    safe_text = _escape_applescript_string(text)
    script = f'''
tell application "Microsoft PowerPoint"
    set targetSlide to slide {int(slide_index)} of active presentation
    set n to count of shapes of targetSlide
    set matched to ""
    repeat with i from 1 to n
        set shp to shape i of targetSlide
        try
            if (name of shp) is "{safe_name}" then
                set content of text range of text frame of shp to "{safe_text}"
                set matched to (name of shp) as text
                exit repeat
            end if
        end try
    end repeat
    if matched is "" then
        error "No shape named '{safe_name}' on slide {int(slide_index)}"
    end if
    return matched
end tell
'''
    out = _run_osascript(script)
    return {"shape_name": out}


@mcp.tool()
def sidecar_move_slide(from_index: int, to_index: int) -> dict[str, Any]:
    """Reorder slides natively via PowerPoint AppleScript.

    Args:
        from_index: 1-based current position of the slide to move.
        to_index: 1-based target position.

    Returns:
        dict with `new_index` (the slide's index after the move).
    """
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
def sidecar_set_slide_layout(slide_index: int, layout: str) -> dict[str, Any]:
    """Change the *flag* of a slide's layout to one of PowerPoint's built-in enums.

    **Caveat verified live**: this only flips PowerPoint's `layout` property (the enum
    that says "this slide is using layout X"). It does NOT restructure the shapes on the
    slide — existing shapes (placeholders, freeform overlays) are left in place. So the
    visual result is usually identical to before; only the metadata changes.

    To actually reshape a slide to match a different layout, use
    `sidecar_set_slide_layout_from_template` (which adds the new layout's placeholders
    on top — still doesn't delete old shapes, but at least surfaces the layout's slots).

    Args:
        slide_index: 1-based index of the slide.
        layout: Built-in enum name without `slide layout ` prefix (e.g. "blank").

    Returns:
        dict with `slide_index` and the layout flag that was applied.
    """
    safe_layout = layout.strip().lower()
    script = f'''
tell application "Microsoft PowerPoint"
    set targetSlide to slide {int(slide_index)} of active presentation
    set layout of targetSlide to slide layout {safe_layout}
    return slide index of targetSlide
end tell
'''
    out = _run_osascript(script)
    return {
        "slide_index": int(out) if out.isdigit() else out,
        "layout_applied": f"slide layout {safe_layout}",
    }


@mcp.tool()
def sidecar_set_slide_layout_from_template(
    slide_index: int, source_slide_index: int
) -> dict[str, Any]:
    """Apply the *custom layout* of an existing slide to another slide.

    For corporate templates that ship dozens of custom slide layouts (each with its own
    placeholder geometry, fonts, decorative master shapes), this is the only way to
    address a specific layout via AppleScript — built-in enums (`slide layout blank`,
    etc.) only cover ~30 standard layouts, not custom ones.

    Mechanism: `set custom layout of slide N to (custom layout of slide M)`. The target
    slide gains M's layout's placeholder shapes; existing shapes are NOT removed (this
    is layered, not replacement). If you want a clean visual match, delete unwanted
    shapes first.

    Args:
        slide_index: 1-based index of the slide whose layout to change.
        source_slide_index: 1-based index of a slide whose custom layout to copy.

    Returns:
        dict with `slide_index` and `shapes_added` (heuristic: shape-count delta).
    """
    script = f'''
tell application "Microsoft PowerPoint"
    set p to active presentation
    set targetSlide to slide {int(slide_index)} of p
    set beforeCount to count of shapes of targetSlide
    set sourceLayout to custom layout of slide {int(source_slide_index)} of p
    set custom layout of targetSlide to sourceLayout
    set afterCount to count of shapes of targetSlide
    return (slide index of targetSlide as text) & "|" & beforeCount & "|" & afterCount
end tell
'''
    out = _run_osascript(script)
    parts = out.split("|")
    if len(parts) == 3:
        idx, before, after = parts
        return {
            "slide_index": int(idx) if idx.isdigit() else idx,
            "shapes_before": int(before),
            "shapes_after": int(after),
            "shapes_added": int(after) - int(before),
        }
    return {"raw": out}


@mcp.tool()
def sidecar_replace_text_in_shape_by_name(
    slide_index: int, shape_name: str, old: str, new: str
) -> dict[str, Any]:
    """Replace a substring in one shape's text. Caveat: collapses styled runs.

    PowerPoint for Mac's AppleScript dictionary does NOT implement a native find/replace
    on text ranges — every syntax variant returns -2741 syntax error. We read the
    current text, do `str.replace` in Python, and write it back via
    `set content of text range`. This collapses every styled run on the text frame into
    one run with the formatting of the first run.

    If preserving styled runs matters (bold/colored phrases mid-text), use python-pptx
    paragraph/run-level edits instead.

    Args:
        slide_index: 1-based index of the slide.
        shape_name: Exact `name` of the shape.
        old: Substring to search for (exact match, case-sensitive).
        new: Replacement string.

    Returns:
        dict with `found` (True if a replacement happened), `before` text, `after` text.
    """
    safe_name = _escape_applescript_string(shape_name)
    # Step 1: read current text.
    read_script = f'''
tell application "Microsoft PowerPoint"
    set targetSlide to slide {int(slide_index)} of active presentation
    set n to count of shapes of targetSlide
    repeat with i from 1 to n
        set shp to shape i of targetSlide
        try
            if (name of shp) is "{safe_name}" then
                return content of text range of text frame of shp
            end if
        end try
    end repeat
    error "No shape named '{safe_name}' on slide {int(slide_index)}"
end tell
'''
    before_text = _run_osascript(read_script)
    if old not in before_text:
        return {"found": False, "before": before_text, "after": before_text}

    after_text = before_text.replace(old, new)
    safe_after = _escape_applescript_string(after_text)
    write_script = f'''
tell application "Microsoft PowerPoint"
    set targetSlide to slide {int(slide_index)} of active presentation
    set n to count of shapes of targetSlide
    repeat with i from 1 to n
        set shp to shape i of targetSlide
        try
            if (name of shp) is "{safe_name}" then
                set content of text range of text frame of shp to "{safe_after}"
                exit repeat
            end if
        end try
    end repeat
end tell
'''
    _run_osascript(write_script)
    return {"found": True, "before": before_text, "after": after_text}


# --- Entry point ----------------------------------------------------------

def main() -> None:
    """Run the MCP server over stdio (default transport for Claude Desktop)."""
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(name)s %(message)s")
    mcp.run()


if __name__ == "__main__":
    main()
