"""PowerPoint tools for the per-app MCP server.

Reads use JXA: a slide's text lives at shape -> text frame -> text range ->
content, guarded by `has text frame` / `has text`. Writes use AppleScript:
slides are created with `make new slide at end of active presentation` (the
`at end of slides of` form fails), and shape text is set via the content chain.
Status never launches PowerPoint: `.running()` is checked first.
"""

import json

from fastmcp.utilities.types import Image

from office_mcp import bridge

# Friendly names -> MsoAnimEffect / MsoAnimTriggerType terms for ppt_add_animation.
_ANIM_EFFECTS = {
    "appear": "animation type appear",
    "fade": "animation type fade",
    "fly in": "animation type fly",
    "wipe": "animation type wipe",
    "zoom": "animation type zoom",
}
_ANIM_TRIGGERS = {
    "click": "on page click",
    "with_previous": "with previous",
    "after_previous": "after previous",
}


def _rgb(color) -> str:
    """An [r, g, b] list (0-255) as an AppleScript color literal."""
    if not (isinstance(color, (list, tuple)) and len(color) == 3):
        raise ValueError(f"color must be [r, g, b], got {color!r}")
    return "{" + ", ".join(str(int(v)) for v in color) + "}"

# Friendly layout names -> EPPSlideLayout enum terms accepted by ppt_add_slide.
_LAYOUTS = {
    "blank": "slide layout blank",
    "title": "slide layout title slide",
    "title only": "slide layout title only",
    "text": "slide layout text slide",
    "two content": "slide layout two column text",
}

_STATUS = """
const pp = Application('Microsoft PowerPoint');
const out = { running: pp.running() };
if (out.running) {
  try {
    const pres = pp.activePresentation;
    out.presentation = pres.name();
    out.slideCount = pres.slides.length;
    try { out.currentSlide = pp.activeWindow.view.slide.slideIndex(); } catch (e) {}
  } catch (e) { out.presentation = null; }
}
JSON.stringify(out);
"""

_LIST_SLIDES = r"""
const pp = Application('Microsoft PowerPoint');
const out = [];
if (pp.running()) {
  const slides = pp.activePresentation.slides;
  for (let s = 0; s < slides.length; s++) {
    const shapes = slides[s].shapes;
    const parts = [];
    for (let i = 0; i < shapes.length; i++) {
      try { if (shapes[i].hasTextFrame() && shapes[i].textFrame.hasText()) parts.push(shapes[i].textFrame.textRange.content()); } catch (e) {}
    }
    out.push({ slide: s + 1, text: parts.join(' • ') });
  }
}
JSON.stringify(out);
"""

# %d is the 1-based slide index.
_READ_SLIDE = """
const pp = Application('Microsoft PowerPoint');
const shapes = pp.activePresentation.slides[%d - 1].shapes;
const out = [];
for (let i = 0; i < shapes.length; i++) {
  let txt = null, has = false;
  try { has = shapes[i].hasTextFrame(); } catch (e) {}
  if (has) { try { if (shapes[i].textFrame.hasText()) txt = shapes[i].textFrame.textRange.content(); } catch (e) {} }
  out.push({ shape: i + 1, name: shapes[i].name(), text: txt });
}
JSON.stringify(out);
"""

_CURRENT_SLIDE = """
const pp = Application('Microsoft PowerPoint');
let out = null;
if (pp.running()) {
  try {
    const idx = pp.activeWindow.view.slide.slideIndex();
    const shapes = pp.activePresentation.slides[idx - 1].shapes;
    const sh = [];
    for (let i = 0; i < shapes.length; i++) {
      let txt = null, has = false;
      try { has = shapes[i].hasTextFrame(); } catch (e) {}
      if (has) { try { if (shapes[i].textFrame.hasText()) txt = shapes[i].textFrame.textRange.content(); } catch (e) {} }
      sh.push({ shape: i + 1, name: shapes[i].name(), text: txt });
    }
    out = { slide: idx, shapes: sh };
  } catch (e) {}
}
JSON.stringify(out);
"""

# %s are shape index then slide index; the new text comes in via argv.
_SET_TEXT = """
on run argv
  tell application "Microsoft PowerPoint"
    set content of text range of text frame of shape %s of slide %s of active presentation to (item 1 of argv)
  end tell
  return "ok"
end run
"""

# PowerPoint for Mac reports the selection type and the selected text reliably,
# but won't hand over the selected shape objects (`shape range of selection` reads
# as empty via both JXA and AppleScript, even when type is "shapes"). So this
# reports type/slide/selectedText; to see which shape is selected, take a
# screenshot — the selection handles show.
_SELECTION = """
const pp = Application('Microsoft PowerPoint');
let out = { type: 'none' };
if (pp.running()) {
  const sel = pp.activeWindow.selection;
  out.type = sel.selectionType().replace('selection type ', '');
  try { out.slide = pp.activeWindow.view.slide.slideIndex(); } catch (e) {}
  if (out.type === 'text') { try { out.selectedText = sel.textRange.content(); } catch (e) {} }
}
JSON.stringify(out);
"""

_SET_SELECTED_TEXT = """
on run argv
  tell application "Microsoft PowerPoint"
    set sel to selection of active window
    if (selection type of sel) is selection type shapes then
      set content of text range of text frame of (item 1 of shape range of sel) to (item 1 of argv)
    else
      set content of (text range of sel) to (item 1 of argv)
    end if
  end tell
  return "ok"
end run
"""

# Notes live in the "Notes Placeholder" shape on a slide's notes page; find it by
# name (its index varies). %d is the 1-based slide index, %s the JSON-escaped text.
_GET_NOTES = """
const pp = Application('Microsoft PowerPoint');
const sh = pp.activePresentation.slides[%d - 1].notesPage.shapes;
let t = null;
for (let i = 0; i < sh.length; i++) {
  if (sh[i].name().indexOf('Notes') >= 0) {
    try { t = sh[i].textFrame.textRange.content(); } catch (e) {}
    break;
  }
}
JSON.stringify(t);
"""

_SET_NOTES = """
const pp = Application('Microsoft PowerPoint');
const sh = pp.activePresentation.slides[%d - 1].notesPage.shapes;
let done = false;
for (let i = 0; i < sh.length; i++) {
  if (sh[i].name().indexOf('Notes') >= 0) {
    sh[i].textFrame.textRange.content = %s;
    done = true;
    break;
  }
}
JSON.stringify(done);
"""


def register(mcp):
    @mcp.tool
    def ppt_status() -> dict:
        """Whether PowerPoint is running and, if so, the active presentation name,
        slide count, and current slide index. Does not launch PowerPoint."""
        return bridge.run_jxa(_STATUS)

    @mcp.tool
    def ppt_list_slides() -> list:
        """Every slide as {slide, text}, where text joins the slide's shape texts."""
        return bridge.run_jxa(_LIST_SLIDES)

    @mcp.tool
    def ppt_read_slide(slide_index: int) -> list:
        """Shapes on a slide (1-based) as {shape, name, text}."""
        return bridge.run_jxa(_READ_SLIDE % int(slide_index))

    @mcp.tool
    def ppt_get_current_slide() -> dict | None:
        """The slide currently shown in the editor: {slide, shapes}."""
        return bridge.run_jxa(_CURRENT_SLIDE)

    @mcp.tool
    def ppt_add_slide(layout: str = "text", position: int | None = None) -> int:
        """Add a slide and return its index. `layout` is one of: blank, title,
        title only, text, two content. By default appends at the end; pass
        `position` to insert after that 1-based slide index."""
        if layout not in _LAYOUTS:
            raise ValueError(f"unknown layout {layout!r}; choose from {sorted(_LAYOUTS)}")
        where = (
            "end of active presentation"
            if position is None
            else f"after (slide {int(position)} of active presentation)"
        )
        script = (
            'tell application "Microsoft PowerPoint"\n'
            f"set s to make new slide at {where} with properties {{layout:{_LAYOUTS[layout]}}}\n"
            "return slide index of s\n"
            "end tell"
        )
        return int(bridge.run_applescript(script))

    @mcp.tool
    def ppt_set_text(slide_index: int, shape_index: int, text: str) -> str:
        """Set the text of a shape (1-based) on a slide (1-based). Use ppt_read_slide
        to find shape indexes."""
        return bridge.run_applescript(_SET_TEXT % (int(shape_index), int(slide_index)), text)

    @mcp.tool
    def ppt_get_selection() -> dict:
        """What the user has selected: {type, slide, selectedText}. type is
        none / slides / shapes / text. PowerPoint won't reveal which shape is
        selected — use ppt_screenshot to see it (the handles show)."""
        return bridge.run_jxa(_SELECTION)

    @mcp.tool
    def ppt_set_selected_text(text: str) -> str:
        """Set the text where the user is working: replaces the selected text /
        inserts at the cursor (text mode), or the whole selected shape's text
        (shape mode)."""
        return bridge.run_applescript(_SET_SELECTED_TEXT, text)

    @mcp.tool
    def ppt_get_notes(slide_index: int) -> str | None:
        """The speaker-notes text for a slide (1-based)."""
        return bridge.run_jxa(_GET_NOTES % int(slide_index))

    @mcp.tool
    def ppt_set_notes(slide_index: int, text: str) -> bool:
        """Set a slide's speaker notes (1-based). Good for leaving context that
        belongs with the slide but shouldn't appear on it."""
        return bridge.run_jxa(_SET_NOTES % (int(slide_index), json.dumps(text)))

    @mcp.tool
    def ppt_set_shape_position(
        slide_index: int,
        shape_index: int,
        left: float | None = None,
        top: float | None = None,
        width: float | None = None,
        height: float | None = None,
    ) -> str:
        """Move and/or resize a shape (1-based indexes), in points. Only the
        arguments you pass are changed."""
        lines = []
        if left is not None:
            lines.append(f"set left position of sh to {float(left)}")
        if top is not None:
            lines.append(f"set top of sh to {float(top)}")
        if width is not None:
            lines.append(f"set width of sh to {float(width)}")
        if height is not None:
            lines.append(f"set height of sh to {float(height)}")
        if not lines:
            return "nothing to change"
        body = "\n    ".join(lines)
        return bridge.run_applescript(
            'tell application "Microsoft PowerPoint"\n'
            f"    set sh to shape {int(shape_index)} of slide {int(slide_index)} of active presentation\n"
            f"    {body}\nend tell\nreturn \"ok\""
        )

    @mcp.tool
    def ppt_add_animation(
        slide_index: int,
        shape_index: int,
        effect: str = "fade",
        trigger: str = "click",
        exit: bool = False,
    ) -> str:
        """Animate a shape (1-based). effect: appear, fade, fly in, wipe, zoom.
        trigger: click, with_previous, after_previous. exit=True animates it out
        instead of in (e.g. a cover that disappears on click to reveal what's under it)."""
        if effect not in _ANIM_EFFECTS:
            raise ValueError(f"unknown effect {effect!r}; choose from {sorted(_ANIM_EFFECTS)}")
        if trigger not in _ANIM_TRIGGERS:
            raise ValueError(f"unknown trigger {trigger!r}; choose from {sorted(_ANIM_TRIGGERS)}")
        return bridge.run_applescript(
            'tell application "Microsoft PowerPoint"\n'
            f"  set sld to slide {int(slide_index)} of active presentation\n"
            f"  set fx to add effect (main sequence of timeline of sld) for shape {int(shape_index)} of sld fx {_ANIM_EFFECTS[effect]} trigger {_ANIM_TRIGGERS[trigger]}\n"
            f"  set exit animation of fx to {'true' if exit else 'false'}\n"
            '  return "ok"\n'
            "end tell"
        )

    @mcp.tool
    def ppt_format_shape(
        slide_index: int,
        shape_index: int,
        fill_color: list[int] | None = None,
        border_color: list[int] | None = None,
        border_weight: float | None = None,
    ) -> str:
        """Format a shape's fill and border (1-based indexes). Colors are [r, g, b]
        (0-255); border_weight is in points. Only what you pass changes."""
        lines = []
        if fill_color is not None:
            lines.append(f"set fore color of fill format of sh to {_rgb(fill_color)}")
        if border_color is not None:
            lines.append(f"set fore color of line format of sh to {_rgb(border_color)}")
        if border_weight is not None:
            lines.append(f"set line weight of line format of sh to {float(border_weight)}")
        if not lines:
            return "nothing to change"
        body = "\n    ".join(lines)
        return bridge.run_applescript(
            'tell application "Microsoft PowerPoint"\n'
            f"    set sh to shape {int(shape_index)} of slide {int(slide_index)} of active presentation\n"
            f"    {body}\nend tell\nreturn \"ok\""
        )

    @mcp.tool
    def ppt_format_text(
        slide_index: int,
        shape_index: int,
        bold: bool | None = None,
        italic: bool | None = None,
        underline: bool | None = None,
        size: float | None = None,
        color: list[int] | None = None,
    ) -> str:
        """Format a shape's text font (1-based indexes). color is [r, g, b]; size
        in points. Only the arguments you pass change."""
        lines = ["set fnt to font of text range of text frame of sh"]
        if bold is not None:
            lines.append(f"set bold of fnt to {'true' if bold else 'false'}")
        if italic is not None:
            lines.append(f"set italic of fnt to {'true' if italic else 'false'}")
        if underline is not None:
            lines.append(f"set underline of fnt to {'true' if underline else 'false'}")
        if size is not None:
            lines.append(f"set font size of fnt to {float(size)}")
        if color is not None:
            lines.append(f"set font color of fnt to {_rgb(color)}")
        if len(lines) == 1:
            return "nothing to change"
        body = "\n    ".join(lines)
        return bridge.run_applescript(
            'tell application "Microsoft PowerPoint"\n'
            f"    set sh to shape {int(shape_index)} of slide {int(slide_index)} of active presentation\n"
            f"    {body}\nend tell\nreturn \"ok\""
        )

    @mcp.tool
    def ppt_add_textbox(
        slide_index: int,
        text: str = "",
        left: float = 100.0,
        top: float = 100.0,
        width: float = 300.0,
        height: float = 50.0,
    ) -> int:
        """Add a text box to a slide (1-based), positioned/sized in points. Returns
        the new shape's index (usable with ppt_format_shape / ppt_set_shape_position)."""
        script = (
            "on run argv\n"
            'tell application "Microsoft PowerPoint"\n'
            f"  set s to slide {int(slide_index)} of active presentation\n"
            f"  set tb to make new text box at s with properties {{left position:{float(left)}, top:{float(top)}, width:{float(width)}, height:{float(height)}}}\n"
            "  set content of text range of text frame of tb to (item 1 of argv)\n"
            "  return (count of shapes of s) as string\n"
            "end tell\n"
            "end run"
        )
        return int(bridge.run_applescript(script, text))

    @mcp.tool
    def ppt_add_image(
        slide_index: int,
        path: str,
        left: float = 100.0,
        top: float = 100.0,
        width: float = 0.0,
        height: float = 0.0,
    ) -> int:
        """Add an image file to a slide (1-based). width/height of 0 keep the image's
        native size. Returns the new shape's index."""
        props = ["file name:(item 1 of argv)", f"left position:{float(left)}", f"top:{float(top)}"]
        if width > 0:
            props.append(f"width:{float(width)}")
        if height > 0:
            props.append(f"height:{float(height)}")
        script = (
            "on run argv\n"
            'tell application "Microsoft PowerPoint"\n'
            f"  set s to slide {int(slide_index)} of active presentation\n"
            f"  make new picture at s with properties {{{', '.join(props)}}}\n"
            "  return (count of shapes of s) as string\n"
            "end tell\n"
            "end run"
        )
        return int(bridge.run_applescript(script, path))

    @mcp.tool
    def ppt_delete_slide(slide_index: int) -> str:
        """Delete a slide (1-based)."""
        bridge.run_applescript(
            f'tell application "Microsoft PowerPoint" to delete slide {int(slide_index)} of active presentation'
        )
        return f"deleted slide {slide_index}"

    @mcp.tool
    def ppt_move_slide(slide_index: int, before_index: int) -> str:
        """Move a slide so it sits just before the slide currently at before_index
        (both 1-based)."""
        bridge.run_applescript(
            f'tell application "Microsoft PowerPoint" to move slide {int(slide_index)} '
            f"of active presentation to before slide {int(before_index)} of active presentation"
        )
        return f"moved slide {slide_index} before {before_index}"

    @mcp.tool
    def ppt_screenshot() -> Image:
        """A PNG screenshot of the PowerPoint window, to visually check the slide."""
        return Image(data=bridge.screenshot("Microsoft PowerPoint"), format="png")

    @mcp.tool
    def run_applescript(script: str) -> str:
        """Escape hatch: run arbitrary AppleScript and return its result. Use only
        when a dedicated PowerPoint tool doesn't cover the operation."""
        return bridge.run_applescript(script)
