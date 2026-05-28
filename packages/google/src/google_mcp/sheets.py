"""Sheets tools for the Google MCP — the first app implemented.

Everything goes through the Sheets v4 API client built in `auth.py`. Edits to the
spreadsheet on Google's servers show up live in the user's open browser tab via
Google's realtime collaboration sync.
"""

from typing import Any

from google_mcp import auth

INSTRUCTIONS = """\
You are working live on the user's open Google Sheet, via the Sheets API. Edits
show up in their open browser tab through Google's realtime sync.

First, point the server at the right spreadsheet: call sheets_set_active(url) with
the URL from the browser address bar. After that all other tools default to that
spreadsheet. Confirm with sheets_status. Ranges are A1-style (e.g. "A1:C10"); pass
`sheet` (a tab name) to target a specific tab, omit it for the first/active tab.

The OAuth consent flow runs once on first use (a browser tab opens for approval);
afterwards the cached token is reused.
"""


def _qualified(cell_range: str, sheet: str | None) -> str:
    """An API range string: 'A1:B2' (active tab) or 'Sheet1!A1:B2' (named tab)."""
    return f"'{sheet}'!{cell_range}" if sheet else cell_range


def register(mcp):
    @mcp.tool
    def sheets_set_active(url_or_id: str) -> str:
        """Point the server at a spreadsheet by its URL (or ID). Persists across
        sessions. All other Sheets tools default to this spreadsheet."""
        sid = auth.set_active("sheets", url_or_id)
        return f"active spreadsheet set: {sid}"

    @mcp.tool
    def sheets_status() -> dict:
        """The active spreadsheet's title, ID, URL, and tab names. Use after
        sheets_set_active to confirm you're pointed at the right doc."""
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
