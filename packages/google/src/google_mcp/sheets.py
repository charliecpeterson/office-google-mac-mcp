"""Sheets tools — A1-anchored editing on the user's open Google Sheet.

Everything goes through the Sheets v4 API client built in `auth.py`. Edits to
the spreadsheet on Google's servers show up live in the user's open browser tab
via Google's realtime collaboration sync.
"""

import re
from typing import Any

from google_mcp import auth

INSTRUCTIONS = """\
You are working live on the user's open Google Sheet, via the Sheets API. Edits
show up in their open browser tab through Google's realtime sync.

First, point the server at the right spreadsheet: call sheets_set_active(url) with
the URL from the browser address bar. After that all other tools default to that
spreadsheet. Confirm with sheets_status. Ranges are A1-style (e.g. "A1:C10");
pass `sheet` (a tab name) to target a specific tab, omit it for the first tab.

Values: sheets_read_range, sheets_write_range (2-D), sheets_set_cell (1-cell),
sheets_set_formula. Strings starting with "=" land as formulas in write_range
and set_cell, but set_formula makes intent explicit.

Tables: sheets_write_table writes values, bolds the header, adds thin borders,
and autofits in one call.

Formatting: sheets_format_range (bold/italic/font/size/colors/number_format/
alignment), sheets_set_borders. Colors are [r, g, b] 0-255.

Structure: sheets_insert_rows / delete_rows / insert_columns / delete_columns,
sheets_autofit, sheets_sort, sheets_autofilter. Sheet management:
sheets_add_sheet, sheets_delete_sheet, sheets_rename_sheet,
sheets_activate_sheet. Charts: sheets_create_chart.

Escape hatch: sheets_batch_update(requests) takes a raw list of Sheets API
batchUpdate requests when no tool fits.
"""


_A1_CELL = re.compile(r"^([A-Za-z]+)(\d+)$")
_A1_ENDPOINT = re.compile(r"^(?P<col>[A-Za-z]+)?(?P<row>\d+)?$")


def _col_letters_to_index(letters: str) -> int:
    """A -> 0, B -> 1, ..., Z -> 25, AA -> 26."""
    idx = 0
    for c in letters.upper():
        idx = idx * 26 + (ord(c) - ord("A") + 1)
    return idx - 1


def _col_index_to_letters(idx: int) -> str:
    letters = ""
    n = idx + 1
    while n > 0:
        n, r = divmod(n - 1, 26)
        letters = chr(ord("A") + r) + letters
    return letters


def _parse_a1(cell_range: str) -> tuple[str | None, int | None, int | None, int | None, int | None]:
    """Parse 'A1:B5', 'Sheet1!A1:B5', 'A:C' (cols only), or '1:5' (rows only).
    Returns (sheet_name, start_row, start_col, end_row_excl, end_col_excl) —
    0-based, end bounds exclusive. Row or column bounds may be None for an
    unconstrained side (the API will treat None as 'all rows'/'all columns')."""
    s = cell_range.strip()
    sheet = None
    if "!" in s:
        sheet, s = s.split("!", 1)
        sheet = sheet.strip().strip("'")
    if ":" in s:
        a, b = s.split(":", 1)
    else:
        a = b = s
    ma = _A1_ENDPOINT.match(a.strip())
    mb = _A1_ENDPOINT.match(b.strip())
    if not ma or not mb or not (ma.group("col") or ma.group("row")):
        raise ValueError(f"invalid A1 range {cell_range!r}")
    sc = _col_letters_to_index(ma.group("col")) if ma.group("col") else None
    sr = int(ma.group("row")) - 1 if ma.group("row") else None
    ec = _col_letters_to_index(mb.group("col")) + 1 if mb.group("col") else None
    er = int(mb.group("row")) if mb.group("row") else None
    return sheet, sr, sc, er, ec


def _qualified(cell_range: str, sheet: str | None) -> str:
    """An API range string: 'A1:B2' (active tab) or 'Sheet1!A1:B2' (named tab)."""
    return f"'{sheet}'!{cell_range}" if sheet else cell_range


def _sheets_meta(svc, spreadsheet_id: str) -> list[dict]:
    return (
        svc.spreadsheets()
        .get(spreadsheetId=spreadsheet_id, fields="sheets.properties")
        .execute()
        .get("sheets", [])
    )


def _sheet_id(svc, spreadsheet_id: str, sheet_name: str | None) -> int:
    sheets = _sheets_meta(svc, spreadsheet_id)
    if not sheets:
        raise RuntimeError("spreadsheet has no sheets")
    if sheet_name is None:
        return sheets[0]["properties"]["sheetId"]
    for s in sheets:
        if s["properties"]["title"] == sheet_name:
            return s["properties"]["sheetId"]
    raise ValueError(f"sheet {sheet_name!r} not found")


def _grid_range(svc, spreadsheet_id: str, cell_range: str, default_sheet: str | None) -> dict:
    """Convert an A1 range + sheet name to a Sheets API GridRange. Missing row
    or column bounds (e.g. 'A:C') are omitted from the result so the API treats
    them as unconstrained."""
    parsed_sheet, sr, sc, er, ec = _parse_a1(cell_range)
    sheet_name = parsed_sheet if parsed_sheet is not None else default_sheet
    grid: dict = {"sheetId": _sheet_id(svc, spreadsheet_id, sheet_name)}
    if sr is not None:
        grid["startRowIndex"] = sr
    if er is not None:
        grid["endRowIndex"] = er
    if sc is not None:
        grid["startColumnIndex"] = sc
    if ec is not None:
        grid["endColumnIndex"] = ec
    return grid


def _rgb_color(rgb) -> dict:
    if len(rgb) != 3 or not all(0 <= int(c) <= 255 for c in rgb):
        raise ValueError("color must be [r, g, b] 0-255")
    r, g, b = (int(c) for c in rgb)
    return {"red": r / 255, "green": g / 255, "blue": b / 255}


_BORDER_STYLES = {"SOLID", "SOLID_MEDIUM", "SOLID_THICK", "DOTTED", "DASHED", "DOUBLE", "NONE"}

_CHART_TYPES = {
    "column": ("BASIC", "COLUMN"),
    "bar": ("BASIC", "BAR"),
    "line": ("BASIC", "LINE"),
    "area": ("BASIC", "AREA"),
    "scatter": ("BASIC", "SCATTER"),
    "combo": ("BASIC", "COMBO"),
    "pie": ("PIE", None),
}

_NUMBER_FORMAT_TYPES = {
    "TEXT", "NUMBER", "PERCENT", "CURRENCY", "DATE", "TIME", "DATE_TIME", "SCIENTIFIC",
}

_ALIGNMENTS_H = {"LEFT", "CENTER", "RIGHT"}


def _detect_number_format_type(pattern: str) -> str:
    p = pattern.lower()
    if "%" in p:
        return "PERCENT"
    if "$" in p or "€" in p or "£" in p:
        return "CURRENCY"
    if any(t in p for t in ("yyyy", "yy", "mm", "dd")) and any(t in p for t in (":", "h", "s")):
        return "DATE_TIME"
    if any(t in p for t in ("yyyy", "yy", "mm", "dd")):
        return "DATE"
    if any(t in p for t in ("hh", "ss")) or ":" in p:
        return "TIME"
    return "NUMBER"


def register(mcp):
    @mcp.tool
    def sheets_set_active(url_or_id: str) -> str:
        """Point the server at a spreadsheet by URL (or ID). Persists across sessions."""
        sid = auth.set_active("sheets", url_or_id)
        return f"active spreadsheet set: {sid}"

    @mcp.tool
    def sheets_status() -> dict:
        """The active spreadsheet's title, ID, URL, and tab names."""
        sid = auth.require_active("sheets")
        meta = (
            auth.sheets_service()
            .spreadsheets()
            .get(spreadsheetId=sid, fields="properties.title,sheets.properties(title,sheetId,gridProperties)")
            .execute()
        )
        return {
            "id": sid,
            "title": meta["properties"]["title"],
            "url": f"https://docs.google.com/spreadsheets/d/{sid}/edit",
            "tabs": [s["properties"]["title"] for s in meta.get("sheets", [])],
        }

    @mcp.tool
    def sheets_list_tabs() -> list:
        """Tab names + sheetIds + grid dimensions for the active spreadsheet."""
        sid = auth.require_active("sheets")
        meta = (
            auth.sheets_service()
            .spreadsheets()
            .get(spreadsheetId=sid, fields="sheets.properties(title,sheetId,gridProperties)")
            .execute()
        )
        return [
            {
                "title": s["properties"]["title"],
                "sheetId": s["properties"]["sheetId"],
                "rows": s["properties"].get("gridProperties", {}).get("rowCount"),
                "cols": s["properties"].get("gridProperties", {}).get("columnCount"),
            }
            for s in meta.get("sheets", [])
        ]

    @mcp.tool
    def sheets_read_range(cell_range: str, sheet: str | None = None) -> Any:
        """Values in an A1-style range. Multi-cell returns a 2-D list (ragged rows
        possible — Sheets returns only used cells per row); single cell returns a
        scalar."""
        sid = auth.require_active("sheets")
        resp = (
            auth.sheets_service()
            .spreadsheets()
            .values()
            .get(spreadsheetId=sid, range=_qualified(cell_range, sheet))
            .execute()
        )
        values = resp.get("values", [])
        if not values:
            return None
        if len(values) == 1 and len(values[0]) == 1:
            return values[0][0]
        return values

    @mcp.tool
    def sheets_write_range(
        cell_range: str, values: list[list], sheet: str | None = None
    ) -> dict:
        """Write a 2-D list into an A1-style range. Strings starting with '=' are
        treated as formulas. Returns {updatedRange, updatedCells}."""
        sid = auth.require_active("sheets")
        resp = (
            auth.sheets_service()
            .spreadsheets()
            .values()
            .update(
                spreadsheetId=sid,
                range=_qualified(cell_range, sheet),
                valueInputOption="USER_ENTERED",
                body={"values": values},
            )
            .execute()
        )
        return {"updatedRange": resp.get("updatedRange"), "updatedCells": resp.get("updatedCells")}

    @mcp.tool
    def sheets_set_cell(cell: str, value, sheet: str | None = None) -> dict:
        """Set a single cell. Strings starting with '=' land as formulas."""
        return sheets_write_range(cell, [[value]], sheet)

    @mcp.tool
    def sheets_set_formula(cell: str, formula: str, sheet: str | None = None) -> dict:
        """Set a formula in a single cell. Leading '=' is added if missing."""
        f = formula if formula.startswith("=") else "=" + formula
        return sheets_write_range(cell, [[f]], sheet)

    @mcp.tool
    def sheets_format_range(
        cell_range: str,
        sheet: str | None = None,
        bold: bool | None = None,
        italic: bool | None = None,
        underline: bool | None = None,
        strikethrough: bool | None = None,
        font: str | None = None,
        size: float | None = None,
        font_color: list | None = None,
        fill_color: list | None = None,
        number_format: str | None = None,
        horizontal_alignment: str | None = None,
    ) -> dict:
        """Format an A1 range. Pass only the attributes to change.
        number_format takes a pattern like "0.00", "0.0%", "$#,##0.00", "yyyy-mm-dd".
        horizontal_alignment is LEFT, CENTER, or RIGHT. Colors are [r, g, b] 0-255."""
        text_fmt: dict = {}
        text_field_paths: list[str] = []
        if bold is not None:
            text_fmt["bold"] = bold
            text_field_paths.append("textFormat.bold")
        if italic is not None:
            text_fmt["italic"] = italic
            text_field_paths.append("textFormat.italic")
        if underline is not None:
            text_fmt["underline"] = underline
            text_field_paths.append("textFormat.underline")
        if strikethrough is not None:
            text_fmt["strikethrough"] = strikethrough
            text_field_paths.append("textFormat.strikethrough")
        if font is not None:
            text_fmt["fontFamily"] = font
            text_field_paths.append("textFormat.fontFamily")
        if size is not None:
            text_fmt["fontSize"] = int(size)
            text_field_paths.append("textFormat.fontSize")
        if font_color is not None:
            text_fmt["foregroundColor"] = _rgb_color(font_color)
            text_field_paths.append("textFormat.foregroundColor")

        cell_format: dict = {}
        field_paths: list[str] = []
        if text_field_paths:
            cell_format["textFormat"] = text_fmt
            field_paths.extend(text_field_paths)
        if fill_color is not None:
            cell_format["backgroundColor"] = _rgb_color(fill_color)
            field_paths.append("backgroundColor")
        if number_format is not None:
            cell_format["numberFormat"] = {
                "type": _detect_number_format_type(number_format),
                "pattern": number_format,
            }
            field_paths.append("numberFormat")
        if horizontal_alignment is not None:
            if horizontal_alignment.upper() not in _ALIGNMENTS_H:
                raise ValueError(
                    f"horizontal_alignment must be one of {sorted(_ALIGNMENTS_H)}"
                )
            cell_format["horizontalAlignment"] = horizontal_alignment.upper()
            field_paths.append("horizontalAlignment")
        if not field_paths:
            raise ValueError(
                "no formatting changes requested; pass at least one of "
                "bold/italic/underline/strikethrough/font/size/font_color/"
                "fill_color/number_format/horizontal_alignment"
            )

        sid = auth.require_active("sheets")
        svc = auth.sheets_service()
        fields = ",".join(f"userEnteredFormat.{p}" for p in field_paths)
        svc.spreadsheets().batchUpdate(
            spreadsheetId=sid,
            body={"requests": [
                {"repeatCell": {
                    "range": _grid_range(svc, sid, cell_range, sheet),
                    "cell": {"userEnteredFormat": cell_format},
                    "fields": fields,
                }}
            ]},
        ).execute()
        return {"formatted": cell_range, "fields": field_paths}

    @mcp.tool
    def sheets_set_borders(
        cell_range: str,
        style: str = "SOLID",
        color: list | None = None,
        sheet: str | None = None,
        inner: bool = True,
    ) -> dict:
        """Put borders on a range. style: SOLID, SOLID_MEDIUM, SOLID_THICK, DOTTED,
        DASHED, DOUBLE, NONE. color is [r, g, b] 0-255 (default black). If `inner`
        is True (default), grid lines between cells are also drawn."""
        if style not in _BORDER_STYLES:
            raise ValueError(f"unknown style {style!r}; choose {sorted(_BORDER_STYLES)}")
        border = {"style": style}
        if color is not None:
            border["color"] = _rgb_color(color)
        sid = auth.require_active("sheets")
        svc = auth.sheets_service()
        req = {
            "range": _grid_range(svc, sid, cell_range, sheet),
            "top": border, "bottom": border, "left": border, "right": border,
        }
        if inner:
            req["innerHorizontal"] = border
            req["innerVertical"] = border
        svc.spreadsheets().batchUpdate(
            spreadsheetId=sid, body={"requests": [{"updateBorders": req}]}
        ).execute()
        return {"bordered": cell_range, "style": style}

    @mcp.tool
    def sheets_sort(
        cell_range: str,
        key_column: str,
        ascending: bool = True,
        has_header: bool = True,
        sheet: str | None = None,
    ) -> dict:
        """Sort a range by one column. key_column is a letter ("A", "B", ...).
        has_header skips the first row."""
        sid = auth.require_active("sheets")
        svc = auth.sheets_service()
        grid = _grid_range(svc, sid, cell_range, sheet)
        if has_header:
            grid = {**grid, "startRowIndex": grid["startRowIndex"] + 1}
        col_idx = _col_letters_to_index(key_column)
        svc.spreadsheets().batchUpdate(
            spreadsheetId=sid,
            body={"requests": [
                {"sortRange": {
                    "range": grid,
                    "sortSpecs": [{
                        "dimensionIndex": col_idx,
                        "sortOrder": "ASCENDING" if ascending else "DESCENDING",
                    }],
                }}
            ]},
        ).execute()
        return {"sorted": cell_range, "key": key_column, "ascending": ascending}

    @mcp.tool
    def sheets_autofilter(cell_range: str, sheet: str | None = None) -> dict:
        """Set a basic filter (the dropdown filter UI) over a range with headers.
        Calling on an unfiltered sheet adds it; the first row is the header."""
        sid = auth.require_active("sheets")
        svc = auth.sheets_service()
        svc.spreadsheets().batchUpdate(
            spreadsheetId=sid,
            body={"requests": [
                {"setBasicFilter": {
                    "filter": {"range": _grid_range(svc, sid, cell_range, sheet)}
                }}
            ]},
        ).execute()
        return {"filter_on": cell_range}

    @mcp.tool
    def sheets_create_chart(
        cell_range: str,
        chart_type: str = "column",
        sheet: str | None = None,
        title: str | None = None,
    ) -> dict:
        """Create a chart from an A1 range. chart_type: column, bar, line, area,
        scatter, combo, pie. Returns the new chart's id."""
        if chart_type not in _CHART_TYPES:
            raise ValueError(
                f"unknown chart_type {chart_type!r}; choose {sorted(_CHART_TYPES)}"
            )
        sid = auth.require_active("sheets")
        svc = auth.sheets_service()
        grid = _grid_range(svc, sid, cell_range, sheet)
        # Each ChartSourceRange.sources entry must be a single row OR single column.
        # Split the input range: first column = domain (labels); the remaining
        # columns each become a series.
        sheet_id = grid["sheetId"]
        sr = grid.get("startRowIndex", 0)
        er = grid.get("endRowIndex")
        sc = grid.get("startColumnIndex", 0)
        ec = grid.get("endColumnIndex")
        if er is None or ec is None or ec - sc < 2:
            raise ValueError(
                "create_chart range must include a label column plus at least one "
                "data column (e.g. 'A1:C10') and have explicit row bounds"
            )

        def col_grid(col: int) -> dict:
            return {
                "sheetId": sheet_id,
                "startRowIndex": sr,
                "endRowIndex": er,
                "startColumnIndex": col,
                "endColumnIndex": col + 1,
            }

        domain_range = {"sourceRange": {"sources": [col_grid(sc)]}}
        spec_kind, basic_kind = _CHART_TYPES[chart_type]
        spec: dict = {}
        if title is not None:
            spec["title"] = title
        if spec_kind == "BASIC":
            spec["basicChart"] = {
                "chartType": basic_kind,
                "legendPosition": "BOTTOM_LEGEND",
                "headerCount": 1,
                "domains": [{"domain": domain_range}],
                "series": [
                    {
                        "series": {"sourceRange": {"sources": [col_grid(col)]}},
                        "targetAxis": "LEFT_AXIS",
                    }
                    for col in range(sc + 1, ec)
                ],
            }
        else:  # PIE
            # Pie takes a single data series.
            spec["pieChart"] = {
                "legendPosition": "RIGHT_LEGEND",
                "domain": domain_range,
                "series": {"sourceRange": {"sources": [col_grid(sc + 1)]}},
            }
        resp = svc.spreadsheets().batchUpdate(
            spreadsheetId=sid,
            body={"requests": [
                {"addChart": {"chart": {
                    "spec": spec,
                    "position": {"overlayPosition": {
                        "anchorCell": {
                            "sheetId": sheet_id,
                            "rowIndex": er,
                            "columnIndex": sc,
                        },
                        "widthPixels": 600, "heightPixels": 371,
                    }},
                }}}
            ]},
        ).execute()
        new_chart = resp["replies"][0]["addChart"]["chart"]
        return {"chartId": new_chart["chartId"], "type": chart_type}

    @mcp.tool
    def sheets_autofit(cell_range: str, sheet: str | None = None) -> dict:
        """Auto-resize columns spanning an A1 range to fit content. Pass a column-
        only range like "A:D" or any range — the column dimension is what's resized."""
        sid = auth.require_active("sheets")
        svc = auth.sheets_service()
        grid = _grid_range(svc, sid, cell_range, sheet)
        svc.spreadsheets().batchUpdate(
            spreadsheetId=sid,
            body={"requests": [
                {"autoResizeDimensions": {"dimensions": {
                    "sheetId": grid["sheetId"],
                    "dimension": "COLUMNS",
                    "startIndex": grid["startColumnIndex"],
                    "endIndex": grid["endColumnIndex"],
                }}}
            ]},
        ).execute()
        return {"autofit": cell_range}

    @mcp.tool
    def sheets_insert_rows(at_row: int, count: int = 1, sheet: str | None = None) -> dict:
        """Insert `count` blank rows before row `at_row` (1-based)."""
        sid = auth.require_active("sheets")
        svc = auth.sheets_service()
        sheet_id = _sheet_id(svc, sid, sheet)
        svc.spreadsheets().batchUpdate(
            spreadsheetId=sid,
            body={"requests": [
                {"insertDimension": {"range": {
                    "sheetId": sheet_id, "dimension": "ROWS",
                    "startIndex": at_row - 1, "endIndex": at_row - 1 + count,
                }, "inheritFromBefore": at_row > 1}}
            ]},
        ).execute()
        return {"inserted_rows": count, "at": at_row}

    @mcp.tool
    def sheets_delete_rows(at_row: int, count: int = 1, sheet: str | None = None) -> dict:
        """Delete `count` rows starting at row `at_row` (1-based)."""
        sid = auth.require_active("sheets")
        svc = auth.sheets_service()
        sheet_id = _sheet_id(svc, sid, sheet)
        svc.spreadsheets().batchUpdate(
            spreadsheetId=sid,
            body={"requests": [
                {"deleteDimension": {"range": {
                    "sheetId": sheet_id, "dimension": "ROWS",
                    "startIndex": at_row - 1, "endIndex": at_row - 1 + count,
                }}}
            ]},
        ).execute()
        return {"deleted_rows": count, "at": at_row}

    @mcp.tool
    def sheets_insert_columns(at_col: str, count: int = 1, sheet: str | None = None) -> dict:
        """Insert `count` blank columns before column `at_col` (a letter, e.g. "C")."""
        sid = auth.require_active("sheets")
        svc = auth.sheets_service()
        sheet_id = _sheet_id(svc, sid, sheet)
        start = _col_letters_to_index(at_col)
        svc.spreadsheets().batchUpdate(
            spreadsheetId=sid,
            body={"requests": [
                {"insertDimension": {"range": {
                    "sheetId": sheet_id, "dimension": "COLUMNS",
                    "startIndex": start, "endIndex": start + count,
                }, "inheritFromBefore": start > 0}}
            ]},
        ).execute()
        return {"inserted_cols": count, "at": at_col.upper()}

    @mcp.tool
    def sheets_delete_columns(at_col: str, count: int = 1, sheet: str | None = None) -> dict:
        """Delete `count` columns starting at column `at_col` (a letter)."""
        sid = auth.require_active("sheets")
        svc = auth.sheets_service()
        sheet_id = _sheet_id(svc, sid, sheet)
        start = _col_letters_to_index(at_col)
        svc.spreadsheets().batchUpdate(
            spreadsheetId=sid,
            body={"requests": [
                {"deleteDimension": {"range": {
                    "sheetId": sheet_id, "dimension": "COLUMNS",
                    "startIndex": start, "endIndex": start + count,
                }}}
            ]},
        ).execute()
        return {"deleted_cols": count, "at": at_col.upper()}

    @mcp.tool
    def sheets_add_sheet(name: str | None = None) -> dict:
        """Add a new tab at the end of the spreadsheet. Returns {sheetId, title}."""
        sid = auth.require_active("sheets")
        svc = auth.sheets_service()
        props: dict = {}
        if name is not None:
            props["title"] = name
        resp = svc.spreadsheets().batchUpdate(
            spreadsheetId=sid,
            body={"requests": [{"addSheet": {"properties": props}}]},
        ).execute()
        new_props = resp["replies"][0]["addSheet"]["properties"]
        return {"sheetId": new_props["sheetId"], "title": new_props["title"]}

    @mcp.tool
    def sheets_delete_sheet(name: str) -> dict:
        """Delete a tab by name."""
        sid = auth.require_active("sheets")
        svc = auth.sheets_service()
        sheet_id = _sheet_id(svc, sid, name)
        svc.spreadsheets().batchUpdate(
            spreadsheetId=sid,
            body={"requests": [{"deleteSheet": {"sheetId": sheet_id}}]},
        ).execute()
        return {"deleted": name}

    @mcp.tool
    def sheets_rename_sheet(old_name: str, new_name: str) -> dict:
        """Rename a tab."""
        sid = auth.require_active("sheets")
        svc = auth.sheets_service()
        sheet_id = _sheet_id(svc, sid, old_name)
        svc.spreadsheets().batchUpdate(
            spreadsheetId=sid,
            body={"requests": [
                {"updateSheetProperties": {
                    "properties": {"sheetId": sheet_id, "title": new_name},
                    "fields": "title",
                }}
            ]},
        ).execute()
        return {"renamed": {"from": old_name, "to": new_name}}

    @mcp.tool
    def sheets_write_table(
        start_cell: str,
        values: list[list],
        header: bool = True,
        sheet: str | None = None,
    ) -> dict:
        """Write a 2-D table, bold the header row (if `header`), add thin borders,
        and autofit the columns. The 'drop in a clean table' workflow."""
        m = _A1_CELL.match(start_cell.strip())
        if not m:
            raise ValueError(f"start_cell must be a cell like 'A1', got {start_cell!r}")
        col0_letters = m.group(1).upper()
        row0 = int(m.group(2))
        rows = len(values)
        cols = max((len(r) for r in values), default=0)
        if rows == 0 or cols == 0:
            raise ValueError("values must be a non-empty 2-D list")
        start_col = _col_letters_to_index(col0_letters)
        end_col_letters = _col_index_to_letters(start_col + cols - 1)
        full_range = f"{col0_letters}{row0}:{end_col_letters}{row0 + rows - 1}"
        sheets_write_range(full_range, values, sheet)
        if header:
            header_range = f"{col0_letters}{row0}:{end_col_letters}{row0}"
            sheets_format_range(header_range, sheet=sheet, bold=True)
        sheets_set_borders(full_range, style="SOLID", sheet=sheet, inner=True)
        col_a1 = f"{col0_letters}:{end_col_letters}"
        sheets_autofit(col_a1, sheet=sheet)
        return {"wrote_table": full_range, "header": header}

    @mcp.tool
    def sheets_batch_update(requests: list) -> dict:
        """Raw escape hatch: send a list of Sheets API batchUpdate requests. Use
        for anything no semantic tool covers. Each request is a dict matching the
        Sheets v4 batchUpdate spec (e.g. {"mergeCells": ...})."""
        if not isinstance(requests, list) or not requests:
            raise ValueError("requests must be a non-empty list of API request dicts")
        sid = auth.require_active("sheets")
        resp = (
            auth.sheets_service()
            .spreadsheets()
            .batchUpdate(spreadsheetId=sid, body={"requests": requests})
            .execute()
        )
        return {"replies": resp.get("replies", [])}
