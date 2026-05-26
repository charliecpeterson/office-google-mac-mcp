"""Word tools for the per-app MCP server.

Reads use JXA (returns JSON we parse). Writes use AppleScript with `on run argv`
so user text passes as an argument, sidestepping quote-escaping. Status never
launches Word: `.running()` is checked first.
"""

from fastmcp.utilities.types import Image

from office_mcp import bridge

# Friendly names -> WdBuiltinStyle enum terms for word_set_style.
_STYLES = {"normal": "style normal", "title": "style title", "subtitle": "style subtitle"}
for _i in range(1, 10):
    _STYLES[f"heading {_i}"] = f"style heading{_i}"

_STATUS = """
const W = Application('Microsoft Word');
const out = { running: W.running() };
if (out.running) {
  out.documentCount = W.documents.length;
  if (out.documentCount > 0) {
    const d = W.activeDocument;
    out.name = d.name();
    try { out.path = d.fullName(); } catch (e) { out.path = null; }
    let sel = '';
    try { sel = W.selection.textObject.content() || ''; } catch (e) {}
    out.hasSelection = sel.length > 0;
    out.selectionLength = sel.length;
  }
}
JSON.stringify(out);
"""

_SELECTION = """
const W = Application('Microsoft Word');
let r = null;
if (W.running()) { try { r = W.selection.textObject.content(); } catch (e) {} }
JSON.stringify(r);
"""

_DOCUMENT_TEXT = """
const W = Application('Microsoft Word');
let r = null;
if (W.running() && W.documents.length > 0) {
  r = W.activeDocument.textObject.content();
  if (%d > 0 && r.length > %d) r = r.slice(0, %d);
}
JSON.stringify(r);
"""

_REPLACE_SELECTION = """
on run argv
  tell application "Microsoft Word"
    set content of (text object of selection) to (item 1 of argv)
  end tell
  return "ok"
end run
"""

_INSERT_AT_END = """
on run argv
  tell application "Microsoft Word"
    end key selection move (a story item)
    type text selection text (item 1 of argv)
  end tell
  return "ok"
end run
"""

_INSERT_AT_CURSOR = """
on run argv
  tell application "Microsoft Word"
    type text selection text (item 1 of argv)
  end tell
  return "ok"
end run
"""

# %s are table, row, column (1-based); cell text comes via argv.
_SET_TABLE_CELL = """
on run argv
  tell application "Microsoft Word"
    set c to get cell from table (table %s of active document) row %s column %s
    set content of text object of c to (item 1 of argv)
  end tell
  return "ok"
end run
"""

_GET_TABLE_CELL = (
    'tell application "Microsoft Word" to return content of text object of '
    "(get cell from table (table %s of active document) row %s column %s)"
)

_INSERT_PICTURE = """
on run argv
  tell application "Microsoft Word"
    make new inline picture at (text object of selection) with properties {file name:(item 1 of argv)}
  end tell
  return "ok"
end run
"""

# %s is the match-case boolean literal; find/replace text come in via argv.
_FIND_REPLACE = """
on run argv
  tell application "Microsoft Word"
    set f to find object of (text object of active document)
    set didReplace to execute find f find text (item 1 of argv) ¬
      replace with (item 2 of argv) replace replace all ¬
      wrap find find continue match case %s
  end tell
  return (didReplace as string)
end run
"""

# WdColorIndex terms accepted by word_apply_formatting's `color`.
_COLORS = frozenset({
    "auto", "black", "blue", "turquoise", "bright green", "pink", "red",
    "yellow", "white", "dark blue", "teal", "green", "violet", "dark red",
    "dark yellow", "gray50", "gray25",
})

_OUTLINE = r"""
const W = Application('Microsoft Word');
const out = [];
if (W.running() && W.documents.length > 0) {
  const paras = W.activeDocument.paragraphs;
  for (let i = 0; i < paras.length; i++) {
    let name = '';
    try { name = paras[i].style.nameLocal(); } catch (e) {}
    const m = name.match(/^Heading (\d)$/);
    if (m) {
      let text = '';
      try { text = paras[i].textObject.content(); } catch (e) {}
      out.push({ level: parseInt(m[1], 10), text: text.replace(/[\r\n]+$/, ''), paragraph: i + 1 });
    }
  }
}
JSON.stringify(out);
"""


def register(mcp):
    @mcp.tool
    def word_status() -> dict:
        """Whether Word is running and, if so, the active document name/path and
        whether any text is selected. Does not launch Word."""
        return bridge.run_jxa(_STATUS)

    @mcp.tool
    def word_get_document_text(max_chars: int = 0) -> str | None:
        """Full text of the active document. max_chars > 0 truncates."""
        return bridge.run_jxa(_DOCUMENT_TEXT % (max_chars, max_chars, max_chars))

    @mcp.tool
    def word_get_selection() -> str | None:
        """Currently selected text in the active document, or null if nothing is selected."""
        return bridge.run_jxa(_SELECTION)

    @mcp.tool
    def word_replace_selection(text: str) -> str:
        """Replace the current selection with text."""
        return bridge.run_applescript(_REPLACE_SELECTION, text)

    @mcp.tool
    def word_insert_text(text: str) -> str:
        """Insert text at the end of the active document."""
        return bridge.run_applescript(_INSERT_AT_END, text)

    @mcp.tool
    def word_insert_at_cursor(text: str) -> str:
        """Insert text at the current cursor position, before any selection."""
        return bridge.run_applescript(_INSERT_AT_CURSOR, text)

    @mcp.tool
    def word_find_replace(find: str, replace: str, match_case: bool = False) -> str:
        """Replace every occurrence of `find` with `replace` across the active
        document. Returns "true" if a match was found."""
        script = _FIND_REPLACE % ("true" if match_case else "false")
        return bridge.run_applescript(script, find, replace)

    @mcp.tool
    def word_apply_formatting(
        bold: bool | None = None,
        italic: bool | None = None,
        underline: bool | None = None,
        size: float | None = None,
        color: str | None = None,
    ) -> str:
        """Apply font formatting to the current selection. Only the arguments you
        pass are changed. `color` is a name like "red", "blue", "dark green"."""
        lines = ["set fnt to font object of selection"]
        if bold is not None:
            lines.append(f"set bold of fnt to {'true' if bold else 'false'}")
        if italic is not None:
            lines.append(f"set italic of fnt to {'true' if italic else 'false'}")
        if underline is not None:
            lines.append(f"set underline of fnt to {'underline single' if underline else 'underline none'}")
        if size is not None:
            lines.append(f"set font size of fnt to {float(size)}")
        if color is not None:
            if color not in _COLORS:
                raise ValueError(f"unknown color {color!r}; choose from {sorted(_COLORS)}")
            lines.append(f"set color index of fnt to {color}")
        body = "\n    ".join(lines)
        return bridge.run_applescript(f'tell application "Microsoft Word"\n    {body}\nend tell\nreturn "ok"')

    @mcp.tool
    def word_get_outline() -> list:
        """The heading structure of the active document: a list of
        {level, text, paragraph} for each Heading 1-9 paragraph."""
        return bridge.run_jxa(_OUTLINE)

    @mcp.tool
    def word_set_style(style: str, paragraph: int | None = None) -> str:
        """Set a paragraph's style. `style` is one of: normal, title, subtitle,
        heading 1 .. heading 9. Applies to the selected paragraph(s) by default;
        pass `paragraph` (1-based) to target a specific one."""
        key = style.strip().lower()
        if key not in _STYLES:
            raise ValueError(f"unknown style {style!r}; choose from {sorted(_STYLES)}")
        target = "selection" if paragraph is None else f"paragraph {int(paragraph)} of active document"
        return bridge.run_applescript(
            f'tell application "Microsoft Word" to set style of {target} to {_STYLES[key]}'
        )

    @mcp.tool
    def word_insert_table(rows: int, columns: int) -> str:
        """Insert an empty table at the cursor."""
        return bridge.run_applescript(
            'tell application "Microsoft Word"\n'
            "  set r to text object of selection\n"
            f"  make new table at r with properties {{number of rows:{int(rows)}, number of columns:{int(columns)}}}\n"
            "end tell\n"
            'return "ok"'
        )

    @mcp.tool
    def word_set_table_cell(row: int, column: int, text: str, table: int = 1) -> str:
        """Set a table cell's text (1-based row/column; `table` is the Nth table, default first)."""
        return bridge.run_applescript(_SET_TABLE_CELL % (int(table), int(row), int(column)), text)

    @mcp.tool
    def word_get_table_cell(row: int, column: int, table: int = 1) -> str:
        """Read a table cell's text (1-based)."""
        raw = bridge.run_applescript(_GET_TABLE_CELL % (int(table), int(row), int(column)))
        return raw.rstrip("\r\n\x07")

    @mcp.tool
    def word_insert_picture(path: str) -> str:
        """Insert an image file as an inline picture at the cursor."""
        return bridge.run_applescript(_INSERT_PICTURE, path)

    @mcp.tool
    def word_screenshot() -> Image:
        """A PNG screenshot of the Word window, to visually check the document."""
        return Image(data=bridge.screenshot("Microsoft Word"), format="png")

    @mcp.tool
    def run_applescript(script: str) -> str:
        """Escape hatch: run arbitrary AppleScript and return its result. Use only
        when a dedicated Word tool doesn't cover the operation."""
        return bridge.run_applescript(script)
