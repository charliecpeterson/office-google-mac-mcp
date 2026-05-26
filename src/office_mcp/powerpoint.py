"""PowerPoint tools for the per-app MCP server.

Reads use JXA: a slide's text lives at shape -> text frame -> text range ->
content, guarded by `has text frame` / `has text`. Writes use AppleScript:
slides are created with `make new slide at end of active presentation` (the
`at end of slides of` form fails), and shape text is set via the content chain.
Status never launches PowerPoint: `.running()` is checked first.
"""

from fastmcp.utilities.types import Image

from office_mcp import bridge

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
    def ppt_screenshot() -> Image:
        """A PNG screenshot of the PowerPoint window, to visually check the slide."""
        return Image(data=bridge.screenshot("Microsoft PowerPoint"), format="png")

    @mcp.tool
    def run_applescript(script: str) -> str:
        """Escape hatch: run arbitrary AppleScript and return its result. Use only
        when a dedicated PowerPoint tool doesn't cover the operation."""
        return bridge.run_applescript(script)
