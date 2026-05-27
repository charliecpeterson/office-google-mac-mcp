"""Excel tools for the per-app MCP server.

All Excel automation goes through JXA. Ranges are addressed with bracket
notation (`sheet.ranges['A1:B2']`) and values round-trip as JSON, so a 2-D range
reads and writes as a nested list. Dynamic values are embedded with json.dumps,
which produces valid JS literals and sidesteps escaping. Status never launches
Excel: `.running()` is checked first.
"""

import json
import re
from typing import Any

from fastmcp.utilities.types import Image

from office_mcp import bridge

_CHART_TYPES = {
    "column": "column clustered",
    "bar": "bar clustered",
    "line": "line markers",
    "scatter": "xy scatter lines",
    "pie": "pie",
    "area": "area",
}
_BORDER_WEIGHTS = {
    "hairline": "border weight hairline",
    "thin": "border weight thin",
    "medium": "border weight medium",
    "thick": "border weight thick",
}

# Sent to MCP clients on connect (FastMCP `instructions`).
INSTRUCTIONS = """\
You are working live inside the user's open Microsoft Excel workbook (macOS, via Apple events).

Ranges are A1-style; reads return 2-D lists. Most tools take an optional `sheet` (name) — omit
it for the active sheet; passing it is how you work across tabs. Colors are [r, g, b] (0-255).
Confirm visual results with excel_screenshot.

Common tasks: prefer excel_write_table (writes values + bold header + borders + autofit in one
call) for dropping in a formatted table. Use excel_set_array_formula for CSE/array formulas.

If no tool fits, use run_applescript. First use prompts a macOS Automation grant (and Screen
Recording for excel_screenshot) on the terminal app — ask the user to approve it.
"""

_STATUS = """
const xl = Application('Microsoft Excel');
const out = { running: xl.running() };
if (out.running) {
  out.workbookCount = xl.workbooks.length;
  if (out.workbookCount > 0) {
    out.workbook = xl.activeWorkbook.name();
    out.activeSheet = xl.activeSheet.name();
    out.sheets = xl.activeWorkbook.worksheets.name();
    try { out.selection = xl.selection.getAddress(); } catch (e) {}
  }
}
JSON.stringify(out);
"""

_LIST_SHEETS = """
const xl = Application('Microsoft Excel');
const ok = xl.running() && xl.workbooks.length > 0;
JSON.stringify(ok ? xl.activeWorkbook.worksheets.name() : []);
"""

_GET_SELECTION = """
const xl = Application('Microsoft Excel');
let out = null;
if (xl.running() && xl.workbooks.length > 0) {
  const s = xl.selection;
  out = { address: s.getAddress(), value: s.value() };
}
JSON.stringify(out);
"""


def _target(sheet):
    """JS expression for the worksheet to act on: a named sheet or the active one."""
    if sheet is None:
        return "xl.activeSheet"
    return f"xl.activeWorkbook.worksheets[{json.dumps(sheet)}]"


def _rgb_list(color):
    """Validate an [r, g, b] color (0-255) and return it as a list of ints."""
    if not (isinstance(color, (list, tuple)) and len(color) == 3):
        raise ValueError(f"color must be [r, g, b], got {color!r}")
    return [int(v) for v in color]


def _col_to_num(col: str) -> int:
    n = 0
    for ch in col.upper():
        n = n * 26 + (ord(ch) - 64)
    return n


def _num_to_col(n: int) -> str:
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _range_command(clause: str, rng: str, sheet):
    """Run an Excel command on a range. `clause` has a {ref} placeholder for the
    range reference; the sheet name (if any) is passed via argv to avoid escaping."""
    if sheet is None:
        ref = f'range "{rng}" of active sheet'
        return bridge.run_applescript(f'tell application "Microsoft Excel" to {clause.format(ref=ref)}')
    ref = f'range "{rng}" of worksheet (item 1 of argv) of active workbook'
    return bridge.run_applescript(
        f'on run argv\ntell application "Microsoft Excel" to {clause.format(ref=ref)}\nend run', sheet
    )


def _run_on_sheet(clause: str, sheet):
    """Run a multi-line Excel clause inside a tell block. `clause` uses the literal
    placeholder __SREF__ for the sheet reference (plain replace, so AppleScript
    list braces in the clause are left alone); the sheet name is passed via argv."""
    if sheet is None:
        body = clause.replace("__SREF__", "active sheet")
        return bridge.run_applescript(f'tell application "Microsoft Excel"\n{body}\nend tell')
    body = clause.replace("__SREF__", "worksheet (item 1 of argv) of active workbook")
    return bridge.run_applescript(f'on run argv\ntell application "Microsoft Excel"\n{body}\nend tell\nend run', sheet)


def register(mcp):
    @mcp.tool
    def excel_status() -> dict:
        """Whether Excel is running and, if so, the active workbook, active sheet,
        all sheet names, and the current selection address. Does not launch Excel."""
        return bridge.run_jxa(_STATUS)

    @mcp.tool
    def excel_list_sheets() -> list:
        """Names of every worksheet in the active workbook."""
        return bridge.run_jxa(_LIST_SHEETS)

    @mcp.tool
    def excel_read_range(cell_range: str, sheet: str | None = None) -> Any:
        """Values in an A1-style range (e.g. "A1:C10"). A multi-cell range returns a
        2-D list; a single cell returns a scalar. Defaults to the active sheet."""
        script = (
            "const xl = Application('Microsoft Excel');\n"
            f"JSON.stringify({_target(sheet)}.ranges[{json.dumps(cell_range)}].value());"
        )
        return bridge.run_jxa(script)

    @mcp.tool
    def excel_get_selection() -> dict | None:
        """Address and values of the current selection."""
        return bridge.run_jxa(_GET_SELECTION)

    @mcp.tool
    def excel_set_selection(value: float | str) -> bool:
        """Set every cell in the current selection to a number or text."""
        script = (
            "const xl = Application('Microsoft Excel');\n"
            f"xl.selection.value = {json.dumps(value)};\n"
            "JSON.stringify(true);"
        )
        return bridge.run_jxa(script)

    @mcp.tool
    def excel_write_range(
        cell_range: str, values: list[list[float | str]], sheet: str | None = None
    ) -> bool:
        """Write a 2-D list of values into an A1-style range. The shape of `values`
        must match the range. Defaults to the active sheet."""
        script = (
            "const xl = Application('Microsoft Excel');\n"
            f"{_target(sheet)}.ranges[{json.dumps(cell_range)}].value = {json.dumps(values)};\n"
            "JSON.stringify(true);"
        )
        return bridge.run_jxa(script)

    @mcp.tool
    def excel_set_cell(cell: str, value: float | str, sheet: str | None = None) -> bool:
        """Set a single cell (e.g. "B2") to a number or text. Defaults to the active sheet."""
        script = (
            "const xl = Application('Microsoft Excel');\n"
            f"{_target(sheet)}.ranges[{json.dumps(cell)}].value = {json.dumps(value)};\n"
            "JSON.stringify(true);"
        )
        return bridge.run_jxa(script)

    @mcp.tool
    def excel_set_formula(cell: str, formula: str, sheet: str | None = None) -> bool:
        """Set a cell's formula in A1 notation (e.g. "=SUM(A1:A10)"). Defaults to the active sheet."""
        script = (
            "const xl = Application('Microsoft Excel');\n"
            f"{_target(sheet)}.ranges[{json.dumps(cell)}].formula = {json.dumps(formula)};\n"
            "JSON.stringify(true);"
        )
        return bridge.run_jxa(script)

    @mcp.tool
    def excel_set_array_formula(cell_range: str, formula: str, sheet: str | None = None) -> bool:
        """Set an array (CSE) formula on a cell or range, e.g. "=SUM(A1:A3*B1:B3)".
        For a multi-cell result, pass the full output range."""
        script = (
            "const xl = Application('Microsoft Excel');\n"
            f"{_target(sheet)}.ranges[{json.dumps(cell_range)}].formulaArray = {json.dumps(formula)};\n"
            "JSON.stringify(true);"
        )
        return bridge.run_jxa(script)

    @mcp.tool
    def excel_format_range(
        cell_range: str,
        sheet: str | None = None,
        bold: bool | None = None,
        italic: bool | None = None,
        size: float | None = None,
        font_color: list[int] | None = None,
        fill_color: list[int] | None = None,
        number_format: str | None = None,
    ) -> bool:
        """Format an A1-style range: font (bold/italic/size/color), cell fill, and
        number format (e.g. "0.00", "0.0%"). Colors are [r, g, b] (0-255). Only the
        arguments you pass change. Defaults to the active sheet."""
        r = f"{_target(sheet)}.ranges[{json.dumps(cell_range)}]"
        lines = ["const xl = Application('Microsoft Excel');", f"const r = {r};"]
        if bold is not None:
            lines.append(f"r.fontObject.bold = {json.dumps(bold)};")
        if italic is not None:
            lines.append(f"r.fontObject.italic = {json.dumps(italic)};")
        if size is not None:
            lines.append(f"r.fontObject.fontSize = {json.dumps(size)};")
        if font_color is not None:
            lines.append(f"r.fontObject.color = {json.dumps(_rgb_list(font_color))};")
        if fill_color is not None:
            lines.append(f"r.interiorObject.color = {json.dumps(_rgb_list(fill_color))};")
        if number_format is not None:
            lines.append(f"r.numberFormat = {json.dumps(number_format)};")
        lines.append("JSON.stringify(true);")
        return bridge.run_jxa("\n".join(lines))

    @mcp.tool
    def excel_insert_rows(at_row: int, count: int = 1, sheet: str | None = None) -> str:
        """Insert `count` blank rows before row `at_row` (1-based)."""
        rng = f"{int(at_row)}:{int(at_row) + int(count) - 1}"
        _range_command("insert into range ({ref})", rng, sheet)
        return f"inserted {count} row(s) at {at_row}"

    @mcp.tool
    def excel_delete_rows(at_row: int, count: int = 1, sheet: str | None = None) -> str:
        """Delete `count` rows starting at row `at_row` (1-based)."""
        rng = f"{int(at_row)}:{int(at_row) + int(count) - 1}"
        _range_command("delete range ({ref})", rng, sheet)
        return f"deleted {count} row(s) at {at_row}"

    @mcp.tool
    def excel_insert_columns(at_col: str, count: int = 1, sheet: str | None = None) -> str:
        """Insert `count` blank columns before column `at_col` (a letter, e.g. "C")."""
        end = _num_to_col(_col_to_num(at_col) + int(count) - 1)
        _range_command("insert into range ({ref})", f"{at_col.upper()}:{end}", sheet)
        return f"inserted {count} column(s) at {at_col.upper()}"

    @mcp.tool
    def excel_delete_columns(at_col: str, count: int = 1, sheet: str | None = None) -> str:
        """Delete `count` columns starting at column `at_col` (a letter, e.g. "C")."""
        end = _num_to_col(_col_to_num(at_col) + int(count) - 1)
        _range_command("delete range ({ref})", f"{at_col.upper()}:{end}", sheet)
        return f"deleted {count} column(s) at {at_col.upper()}"

    @mcp.tool
    def excel_autofit(cell_range: str, sheet: str | None = None) -> str:
        """Auto-fit the column widths spanning an A1-style range (e.g. "A:D")."""
        _range_command("autofit (entire column of {ref})", cell_range, sheet)
        return f"autofit {cell_range}"

    @mcp.tool
    def excel_add_sheet(name: str | None = None) -> str:
        """Add a worksheet at the end of the workbook, optionally named."""
        if name is None:
            bridge.run_applescript('tell application "Microsoft Excel" to make new worksheet at end of active workbook')
            return "added sheet"
        bridge.run_applescript(
            'on run argv\ntell application "Microsoft Excel" to set name of '
            "(make new worksheet at end of active workbook) to (item 1 of argv)\nend run",
            name,
        )
        return f"added sheet '{name}'"

    @mcp.tool
    def excel_delete_sheet(name: str) -> str:
        """Delete a worksheet by name."""
        bridge.run_applescript(
            'on run argv\ntell application "Microsoft Excel" to delete worksheet (item 1 of argv) '
            "of active workbook\nend run",
            name,
        )
        return f"deleted sheet '{name}'"

    @mcp.tool
    def excel_rename_sheet(old_name: str, new_name: str) -> str:
        """Rename a worksheet."""
        bridge.run_applescript(
            'on run argv\ntell application "Microsoft Excel" to set name of worksheet (item 1 of argv) '
            "of active workbook to (item 2 of argv)\nend run",
            old_name,
            new_name,
        )
        return f"renamed '{old_name}' to '{new_name}'"

    @mcp.tool
    def excel_activate_sheet(name: str) -> str:
        """Switch to (activate) a worksheet by name."""
        bridge.run_applescript(
            'on run argv\ntell application "Microsoft Excel" to activate object worksheet (item 1 of argv) '
            "of active workbook\nend run",
            name,
        )
        return f"activated '{name}'"

    @mcp.tool
    def excel_sort(
        cell_range: str,
        key_column: str,
        ascending: bool = True,
        has_header: bool = True,
        sheet: str | None = None,
    ) -> str:
        """Sort an A1 range by a column (a letter, e.g. "C"). has_header keeps the
        first row in place."""
        m = re.match(r"[A-Za-z]+(\d+)", cell_range)
        key = f"{key_column.upper()}{m.group(1) if m else '1'}"
        order = "sort ascending" if ascending else "sort descending"
        header = "header yes" if has_header else "header no"
        clause = (
            f'sort (range "{cell_range}" of __SREF__) key1 (range "{key}" of __SREF__) '
            f"order1 {order} header {header}"
        )
        _run_on_sheet(clause, sheet)
        return f"sorted {cell_range} by {key_column.upper()}"

    @mcp.tool
    def excel_set_borders(
        cell_range: str, weight: str = "thin", color: list[int] | None = None, sheet: str | None = None
    ) -> str:
        """Put borders (all edges + inner grid) on a range. weight: hairline / thin /
        medium / thick. color is [r, g, b] (default black)."""
        if weight not in _BORDER_WEIGHTS:
            raise ValueError(f"unknown weight {weight!r}; choose from {sorted(_BORDER_WEIGHTS)}")
        rgb = _rgb_list(color) if color is not None else [0, 0, 0]
        color_lit = "{" + ", ".join(str(v) for v in rgb) + "}"
        clause = (
            f'set rng to range "{cell_range}" of __SREF__\n'
            "repeat with idx in {edge top, edge bottom, edge left, edge right, inside horizontal, inside vertical}\n"
            "set b to get border rng which border idx\n"
            "set line style of b to continuous\n"
            f"set weight of b to {_BORDER_WEIGHTS[weight]}\n"
            f"set color of b to {color_lit}\n"
            "end repeat"
        )
        _run_on_sheet(clause, sheet)
        return f"bordered {cell_range}"

    @mcp.tool
    def excel_autofilter(cell_range: str, sheet: str | None = None) -> str:
        """Toggle AutoFilter (the dropdown filters) on a range with headers."""
        _range_command("autofilter range ({ref})", cell_range, sheet)
        return f"autofilter on {cell_range}"

    @mcp.tool
    def excel_create_chart(cell_range: str, chart_type: str = "column", sheet: str | None = None) -> str:
        """Create a chart from an A1 range. chart_type: column, bar, line, scatter,
        pie, area."""
        if chart_type not in _CHART_TYPES:
            raise ValueError(f"unknown chart_type {chart_type!r}; choose from {sorted(_CHART_TYPES)}")
        lines = []
        if sheet is not None:
            lines.append("activate object worksheet (item 1 of argv) of active workbook")
        lines.append(f'select range "{cell_range}" of active sheet')
        lines.append("set co to make new chart object at end of active sheet")
        lines.append(f"set chart type of chart of co to {_CHART_TYPES[chart_type]}")
        body = "\n".join(lines)
        if sheet is None:
            bridge.run_applescript(f'tell application "Microsoft Excel"\n{body}\nend tell')
        else:
            bridge.run_applescript(f'on run argv\ntell application "Microsoft Excel"\n{body}\nend tell\nend run', sheet)
        return f"created {chart_type} chart from {cell_range}"

    @mcp.tool
    def excel_write_table(
        start_cell: str,
        values: list[list[float | str]],
        header: bool = True,
        sheet: str | None = None,
    ) -> str:
        """Write a 2-D table starting at `start_cell` and format it in one step: a
        bold, filled header row (if header), thin borders, and autofit columns. The
        'drop in a clean table' workflow."""
        m = re.match(r"([A-Za-z]+)(\d+)$", start_cell)
        if not m:
            raise ValueError(f"start_cell must be a cell like 'A1', got {start_cell!r}")
        col0, row0 = m.group(1).upper(), int(m.group(2))
        rows = len(values)
        cols = max((len(r) for r in values), default=0)
        if rows == 0 or cols == 0:
            raise ValueError("values must be a non-empty 2-D list")
        end_col = _num_to_col(_col_to_num(col0) + cols - 1)
        full = f"{col0}{row0}:{end_col}{row0 + rows - 1}"
        js = [
            "const xl = Application('Microsoft Excel');",
            f"const t = {_target(sheet)};",
            f"t.ranges[{json.dumps(full)}].value = {json.dumps(values)};",
        ]
        if header:
            hdr = f"{col0}{row0}:{end_col}{row0}"
            js.append(f"t.ranges[{json.dumps(hdr)}].fontObject.bold = true;")
            js.append(f"t.ranges[{json.dumps(hdr)}].interiorObject.color = [220, 230, 242];")
        js.append("JSON.stringify(true);")
        bridge.run_jxa("\n".join(js))
        clause = (
            f'set rng to range "{full}" of __SREF__\n'
            "repeat with idx in {edge top, edge bottom, edge left, edge right, inside horizontal, inside vertical}\n"
            "set b to get border rng which border idx\n"
            "set line style of b to continuous\n"
            "set weight of b to border weight thin\n"
            "end repeat\n"
            "autofit (entire column of rng)"
        )
        _run_on_sheet(clause, sheet)
        return f"wrote table {full}"

    @mcp.tool
    def excel_screenshot() -> Image:
        """A PNG screenshot of the Excel window, to visually check the workbook."""
        return Image(data=bridge.screenshot("Microsoft Excel"), format="png")

    @mcp.tool
    def run_applescript(script: str) -> str:
        """Escape hatch: run arbitrary AppleScript and return its result. Use only
        when a dedicated Excel tool doesn't cover the operation."""
        return bridge.run_applescript(script)
