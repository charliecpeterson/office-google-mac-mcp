"""Excel tools for the per-app MCP server.

All Excel automation goes through JXA. Ranges are addressed with bracket
notation (`sheet.ranges['A1:B2']`) and values round-trip as JSON, so a 2-D range
reads and writes as a nested list. Dynamic values are embedded with json.dumps,
which produces valid JS literals and sidesteps escaping. Status never launches
Excel: `.running()` is checked first.
"""

import json
from typing import Any

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
    def run_applescript(script: str) -> str:
        """Escape hatch: run arbitrary AppleScript and return its result. Use only
        when a dedicated Excel tool doesn't cover the operation."""
        return bridge.run_applescript(script)
