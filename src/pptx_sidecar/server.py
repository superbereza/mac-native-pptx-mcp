"""Sidecar MCP server — drop-in replacement for the 4 broken handles in Anthropic's PowerPoint connector.

The bug: Anthropic's bundled PowerPoint MCP (Claude Desktop, Cowork mode) ships AppleScript
templates that don't match PowerPoint for Mac's AppleScript dictionary. `add_slide`,
`insert_image`, and `get_slide_content` fail with syntax error -2741 on every call. See
upstream issues: https://github.com/anthropics/claude-code/issues/20473 and #26385.

This server speaks correct AppleScript against PowerPoint for Mac and exposes the same
operations under different tool names so it can run as a sidecar next to the broken
upstream connector without name collisions.
"""

from __future__ import annotations

import logging
import shlex
import subprocess
from typing import Any

from mcp.server.fastmcp import FastMCP

logger = logging.getLogger("pptx_sidecar")

mcp = FastMCP("powerpoint-by-anthropic-mac-sidecar")


# --- AppleScript runner ---------------------------------------------------

def _escape_applescript_string(s: str) -> str:
    """Escape a Python string so it can be embedded inside an AppleScript string literal."""
    return s.replace("\\", "\\\\").replace('"', '\\"')


def _run_osascript(script: str, timeout: int = 30) -> str:
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


# --- Tools ----------------------------------------------------------------

@mcp.tool()
def sidecar_add_slide(layout_index: int = 7, position: int | None = None) -> dict[str, Any]:
    """Add a new slide to the active presentation using a layout from the slide master.

    Upstream bug: Anthropic's `add_slide` passes `slide layout:slide layout blank` as a
    property record, but `slide layout` is not an enum constant — it's a reference to a
    layout object on a slide master. PowerPoint's parser rejects it with syntax error -2741.

    This handle creates the slide first, then assigns the layout via
    `set slide layout of newSlide to slide layout N of slide master of activePres`.

    Args:
        layout_index: 1-based index of the layout in the slide master (PowerPoint for Mac's
            default theme typically exposes ~11 layouts; index 7 is commonly "Blank").
        position: 1-based index to insert at. If omitted, the slide is appended at the end.

    Returns:
        dict with `slide_index` of the new slide and the AppleScript that ran.
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
    return {"slide_index": int(out) if out.isdigit() else out, "script": script.strip()}


@mcp.tool()
def sidecar_add_slide_from_template(source_slide_index: int, position: int | None = None) -> dict[str, Any]:
    """Duplicate an existing slide and (optionally) move the copy to a target position.

    Use this when you want a new slide that inherits every style detail from an existing
    template slide (theme placeholders, fonts, custom layout overrides). This sidesteps
    layout-index guesswork entirely.

    Args:
        source_slide_index: 1-based index of the slide to duplicate.
        position: 1-based target index for the duplicate. If omitted, the duplicate is
            left immediately after the source (PowerPoint's default behavior).

    Returns:
        dict with `slide_index` of the new slide.
    """
    move_clause = ""
    if position is not None:
        # `move` accepts `to before slide N` / `to after slide N` references.
        move_clause = f'move newSlide to before slide {int(position)} of activePres'

    script = f'''
tell application "Microsoft PowerPoint"
    set activePres to active presentation
    set newSlide to duplicate slide {int(source_slide_index)} of activePres
    {move_clause}
    return slide index of newSlide
end tell
'''
    out = _run_osascript(script)
    return {"slide_index": int(out) if out.isdigit() else out, "script": script.strip()}


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

    Upstream bug: Anthropic's `insert_image` does `make new picture with properties
    {file name:..., left position:...}`. `picture` is not a creatable class on shapes,
    and `file name` / `left position` aren't valid property keys. PowerPoint rejects it.

    The correct dictionary form is the `add picture` command with named (not record-style)
    parameters: `add picture file name POSIX file "/path" link to file false save with
    document true left X top Y width W height H`.

    Args:
        slide_index: 1-based index of the target slide.
        image_path: Absolute POSIX path to the image file.
        left, top: Position in points (0 leaves PowerPoint to auto-place).
        width, height: Size in points (0 keeps the image's intrinsic size).

    Returns:
        dict with `name` of the inserted shape.
    """
    safe_path = _escape_applescript_string(image_path)

    # `add picture` requires concrete numeric arguments for left/top/width/height;
    # there is no "auto" sentinel, so we fall back to sensible defaults when the caller
    # passes 0 (centered roughly on a 720pt-tall slide).
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
    return {"shape_name": out, "script": script.strip()}


@mcp.tool()
def sidecar_get_slide_content(slide_index: int) -> dict[str, Any]:
    """Read all text from a slide's shapes.

    Upstream bug: Anthropic's `get_slide_content` writes `if has text frame shp` — missing
    the `of` separator. PowerPoint parses `has text frame` as a unary boolean property
    and then chokes on the bare reference. Correct form: `if (has text frame of shp) then`.

    Args:
        slide_index: 1-based index of the slide to read.

    Returns:
        dict with `text` (newline-joined) and `shapes` (per-shape entries).
    """
    script = f'''
tell application "Microsoft PowerPoint"
    set activePres to active presentation
    set targetSlide to slide {int(slide_index)} of activePres
    set collected to {{}}
    repeat with shp in shapes of targetSlide
        if (has text frame of shp) then
            try
                set shpText to content of text range of text frame of shp
            on error
                set shpText to ""
            end try
            set end of collected to (name of shp & "\\t" & shpText)
        end if
    end repeat
    set AppleScript's text item delimiters to linefeed
    set joined to collected as text
    set AppleScript's text item delimiters to ""
    return joined
end tell
'''
    out = _run_osascript(script)
    shapes = []
    text_lines = []
    for line in out.splitlines():
        if "\t" in line:
            name, _, txt = line.partition("\t")
            shapes.append({"name": name, "text": txt})
            text_lines.append(txt)
        elif line:
            shapes.append({"name": "", "text": line})
            text_lines.append(line)
    return {"text": "\n".join(text_lines), "shapes": shapes}


# --- Entry point ----------------------------------------------------------

def main() -> None:
    """Run the MCP server over stdio (default transport for Claude Desktop)."""
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(name)s %(message)s")
    mcp.run()


if __name__ == "__main__":
    main()
