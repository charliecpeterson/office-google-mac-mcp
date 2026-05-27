"""Word tools for the per-app MCP server.

Reads use JXA (returns JSON we parse). Writes use AppleScript with `on run argv`
so user text passes as an argument, sidestepping quote-escaping. Status never
launches Word: `.running()` is checked first.
"""

from fastmcp.utilities.types import Image

from office_mcp import bridge

# Sent to MCP clients on connect (FastMCP `instructions`) — usage context that
# works for any client, including small local models.
INSTRUCTIONS = """\
You are working live inside the user's open Microsoft Word document (macOS, via Apple events).

Workflow: to verify edits, prefer the structured reads — word_get_outline (headings),
word_get_paragraphs ({index, style, text} per paragraph), and word_get_stats ({pages, words,
paragraphs}) — over word_screenshot, which is the most expensive call; screenshot only to check
true visual layout. Paragraph/table indices are 1-based.

Editing is paragraph-anchored: read structure with word_get_paragraphs to get the index you want,
then edit by index — word_insert_paragraph(after=N or before=N, style=...) for mid-document
inserts, word_replace_paragraph(N, ...), word_delete_paragraph(N). word_get_paragraph(N) returns
one paragraph's full text (e.g. to build a precise find/replace anchor). word_insert_text only
appends at the end; word_insert_at_cursor / word_replace_selection act at the live cursor/selection.

Building from scratch: word_add_section (styled heading + body), and word_insert_table(data=...) or
word_fill_table for tables (one call, not many word_set_table_cell). word_set_style sets a
paragraph's heading level by index.

Gotchas: set font size/color with word_apply_formatting (it acts on the selection) — raw Apple-event
font sizing fails. word_find_replace understands Word codes (^p, ^l, ^t) but caps the replacement at
255 chars (it errors past that — use word_replace_paragraph / word_insert_paragraph for big text).
Not scriptable in Word: comments. If no tool fits, use run_applescript.

First use prompts a macOS Automation grant (and Screen Recording for word_screenshot) on the
terminal app — ask the user to approve it; a denial returns a clear "not authorized" error.
"""

# Friendly names -> WdBuiltinStyle enum terms for word_set_style.
_STYLES = {"normal": "style normal", "title": "style title", "subtitle": "style subtitle"}
for _i in range(1, 10):
    _STYLES[f"heading {_i}"] = f"style heading{_i}"

# %s is the heading-style enum; heading text and body come via argv.
_ADD_SECTION = """
on run argv
  tell application "Microsoft Word"
    end key selection move (a story item)
    type text selection text (item 1 of argv)
    set style of selection to %s
    type paragraph selection
    set style of selection to style normal
    type text selection text (item 2 of argv)
    type paragraph selection
  end tell
  return "ok"
end run
"""

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
    type paragraph selection
  end tell
  return "ok"
end run
"""

_STATS = (
    'tell application "Microsoft Word"\n'
    "set d to active document\n"
    "return ((compute statistics d statistic statistic pages) as string) & \",\" & "
    "((compute statistics d statistic statistic words) as string) & \",\" & "
    "((compute statistics d statistic statistic paragraphs) as string)\n"
    "end tell"
)

# %d caps the text length returned per paragraph. Table cells (Word marks them
# with \x07) get tagged `table:true` and de-noised.
_PARAGRAPHS = r"""
const W = Application('Microsoft Word');
const out = [];
if (W.running() && W.documents.length > 0) {
  const paras = W.activeDocument.paragraphs;
  for (let i = 0; i < paras.length; i++) {
    let style = '', raw = '';
    try { style = paras[i].style.nameLocal(); } catch (e) {}
    try { raw = paras[i].textObject.content(); } catch (e) {}
    const inTable = raw.indexOf('\x07') >= 0;
    const text = raw.replace(/[\r\n\x07]+$/, '').replace(/\x07/g, ' | ');
    const o = { index: i + 1, style: style, text: text.slice(0, %d) };
    if (inTable) o.table = true;
    out.push(o);
  }
}
JSON.stringify(out);
"""

_GET_PARAGRAPH = r"""
const W = Application('Microsoft Word');
let r = null;
if (W.running() && W.documents.length > 0) {
  try { r = W.activeDocument.paragraphs[%d - 1].textObject.content().replace(/[\r\n\x07]+$/, ''); } catch (e) {}
}
JSON.stringify(r);
"""

# %s is the 1-based paragraph index; new text comes via argv. `& return` keeps the
# paragraph mark so the paragraph isn't merged into the next one.
_REPLACE_PARAGRAPH = """
on run argv
  tell application "Microsoft Word"
    set content of (text object of paragraph %s of active document) to ((item 1 of argv) & return)
  end tell
  return "ok"
end run
"""

_DELETE_PARAGRAPH = """
tell application "Microsoft Word"
  set d to active document
  set s to start of content of (text object of paragraph %s of d)
  set e to end of content of (text object of paragraph %s of d)
  select (create range d start s end e)
  set content of (text object of selection) to ""
end tell
return "ok"
"""

# Fill an existing table. %s are: table index, rows, columns, columns (index math).
_FILL_TABLE = """
on run argv
  tell application "Microsoft Word"
    set t to table %s of active document
    repeat with r from 1 to %s
      repeat with c from 1 to %s
        set content of text object of (get cell from table t row r column c) to (item ((r - 1) * %s + c) of argv)
      end repeat
    end repeat
  end tell
  return "ok"
end run
"""

# Create a table at the cursor and fill it. %s are: rows, columns, rows, columns, columns.
_INSERT_TABLE_DATA = """
on run argv
  tell application "Microsoft Word"
    set t to make new table at (text object of selection) with properties {number of rows:%s, number of columns:%s}
    repeat with r from 1 to %s
      repeat with c from 1 to %s
        set content of text object of (get cell from table t row r column c) to (item ((r - 1) * %s + c) of argv)
      end repeat
    end repeat
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

# %s are left/top/width/height (points); the box text comes via argv.
_ADD_TEXTBOX = """
on run argv
  tell application "Microsoft Word"
    set tb to make new text box at active document with properties {left position:%s, top:%s, width:%s, height:%s}
    set content of text range of text frame of tb to (item 1 of argv)
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
        """Append text as its own new paragraph at the end of the document (it is
        paragraph-terminated, so a following heading/section won't glue onto it)."""
        return bridge.run_applescript(_INSERT_AT_END, text)

    @mcp.tool
    def word_insert_at_cursor(text: str) -> str:
        """Insert text at the current cursor position, before any selection."""
        return bridge.run_applescript(_INSERT_AT_CURSOR, text)

    @mcp.tool
    def word_find_replace(find: str, replace: str, match_case: bool = False) -> str:
        """Replace every occurrence of `find` with `replace` across the active
        document. Returns "true" if a match was found. `find`/`replace` understand
        Word codes (^p paragraph, ^l line break, ^t tab)."""
        if len(replace) > 255:
            raise ValueError(
                "Word's find/replace caps the replacement at 255 characters; for "
                "larger text use word_replace_paragraph or word_insert_paragraph"
            )
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
    def word_get_stats() -> dict:
        """Document statistics: {pages, words, paragraphs}. Use this instead of a
        screenshot to check length and pagination."""
        pages, words, paragraphs = bridge.run_applescript(_STATS).split(",")
        return {"pages": int(pages), "words": int(words), "paragraphs": int(paragraphs)}

    @mcp.tool
    def word_get_paragraphs(max_chars: int = 200) -> list:
        """Every paragraph as {index, style, text} (table cells are tagged
        `table:true`) — a structured view to verify edits and find paragraph
        indexes without a screenshot. Each text is truncated to max_chars."""
        return bridge.run_jxa(_PARAGRAPHS % int(max_chars))

    @mcp.tool
    def word_get_paragraph(index: int) -> str | None:
        """The full, untruncated text of one paragraph (1-based). Use it to build a
        precise find/replace anchor or to read a paragraph in full."""
        return bridge.run_jxa(_GET_PARAGRAPH % int(index))

    @mcp.tool
    def word_insert_paragraph(
        text: str,
        after: int | None = None,
        before: int | None = None,
        style: str | None = None,
    ) -> str:
        """Insert a new paragraph relative to an existing one (1-based): pass
        exactly one of `after` or `before`. Optionally style it (normal, title,
        heading 1-9). This is the way to add text mid-document — read structure with
        word_get_paragraphs first to find the index."""
        if (after is None) == (before is None):
            raise ValueError("pass exactly one of `after` or `before` (a 1-based paragraph index)")
        if after is not None:
            pos, new_index = f"end of content of (text object of paragraph {int(after)} of d)", int(after) + 1
        else:
            pos, new_index = f"start of content of (text object of paragraph {int(before)} of d)", int(before)
        style_line = ""
        if style is not None:
            key = style.strip().lower()
            if key not in _STYLES:
                raise ValueError(f"unknown style {style!r}; choose from {sorted(_STYLES)}")
            style_line = f"\n    set style of paragraph {new_index} of d to {_STYLES[key]}"
        script = (
            "on run argv\n"
            '  tell application "Microsoft Word"\n'
            "    set d to active document\n"
            f"    set r to create range d start ({pos}) end ({pos})\n"
            "    select r\n"
            "    type text selection text (item 1 of argv)\n"
            "    type paragraph selection"
            f"{style_line}\n"
            "  end tell\n"
            '  return "ok"\n'
            "end run"
        )
        return bridge.run_applescript(script, text)

    @mcp.tool
    def word_replace_paragraph(index: int, text: str) -> str:
        """Replace the text of paragraph `index` (1-based), keeping it a separate
        paragraph and its style."""
        return bridge.run_applescript(_REPLACE_PARAGRAPH % int(index), text)

    @mcp.tool
    def word_delete_paragraph(index: int) -> str:
        """Delete paragraph `index` (1-based) entirely."""
        return bridge.run_applescript(_DELETE_PARAGRAPH % (int(index), int(index)))

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
    def word_insert_table(rows: int = 0, columns: int = 0, data: list[list[str]] | None = None) -> str:
        """Insert a table at the cursor. Pass `data` (a 2-D list) to size the table
        to the data and fill it in one call; otherwise give `rows` and `columns` for
        an empty table."""
        if data:
            n_rows = len(data)
            n_cols = max(len(r) for r in data)
            flat = [str(v) for r in data for v in (list(r) + [""] * (n_cols - len(r)))]
            return bridge.run_applescript(
                _INSERT_TABLE_DATA % (n_rows, n_cols, n_rows, n_cols, n_cols), *flat
            )
        if rows < 1 or columns < 1:
            raise ValueError("provide `data`, or `rows` and `columns` >= 1")
        return bridge.run_applescript(
            'tell application "Microsoft Word"\n'
            "  set r to text object of selection\n"
            f"  make new table at r with properties {{number of rows:{int(rows)}, number of columns:{int(columns)}}}\n"
            "end tell\n"
            'return "ok"'
        )

    @mcp.tool
    def word_fill_table(rows: list[list[str]], table: int = 1) -> str:
        """Fill an existing table from a 2-D list in one call (much cheaper than
        many word_set_table_cell calls). The table (the Nth, 1-based) must be at
        least as large as the data."""
        if not rows:
            raise ValueError("rows must be a non-empty 2-D list")
        n_cols = max(len(r) for r in rows)
        flat = [str(v) for r in rows for v in (list(r) + [""] * (n_cols - len(r)))]
        return bridge.run_applescript(_FILL_TABLE % (int(table), len(rows), n_cols, n_cols), *flat)

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
    def word_add_textbox(
        text: str = "",
        left: float = 100.0,
        top: float = 100.0,
        width: float = 200.0,
        height: float = 80.0,
    ) -> str:
        """Add a floating text box to the document, positioned/sized in points."""
        return bridge.run_applescript(
            _ADD_TEXTBOX % (float(left), float(top), float(width), float(height)), text
        )

    @mcp.tool
    def word_add_section(heading: str, body: str = "", level: int = 1) -> str:
        """Append a styled heading and a body paragraph in one step (the common
        'add a section' workflow). level is the heading level, 1-9."""
        if not 1 <= level <= 9:
            raise ValueError("level must be 1-9")
        return bridge.run_applescript(_ADD_SECTION % f"style heading{int(level)}", heading, body)

    @mcp.tool
    def word_screenshot() -> Image:
        """A PNG screenshot of the Word window, to visually check the document."""
        return Image(data=bridge.screenshot("Microsoft Word"), format="png")

    @mcp.tool
    def run_applescript(script: str) -> str:
        """Escape hatch: run arbitrary AppleScript and return its result. Use only
        when a dedicated Word tool doesn't cover the operation."""
        return bridge.run_applescript(script)
