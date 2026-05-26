"""Excel tools for the per-app MCP server.

All Excel automation goes through JXA. Ranges are addressed with bracket
notation (`sheet.ranges['A1:B2']`) and values round-trip as JSON, so a 2-D range
reads and writes as a nested list. Dynamic values are embedded with json.dumps,
which produces valid JS literals and sidesteps escaping. Status never launches
Excel: `.running()` is checked first.
"""

import json
from typing import Any

from fastmcp.utilities.types import Image

from office_mcp import bridge

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
    def excel_screenshot() -> Image:
        """A PNG screenshot of the Excel window, to visually check the workbook."""
        return Image(data=bridge.screenshot("Microsoft Excel"), format="png")

    @mcp.tool
    def run_applescript(script: str) -> str:
        """Escape hatch: run arbitrary AppleScript and return its result. Use only
        when a dedicated Excel tool doesn't cover the operation."""
        return bridge.run_applescript(script)
