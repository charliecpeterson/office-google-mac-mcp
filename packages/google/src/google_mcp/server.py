"""Build a per-app Google MCP server (sheets / docs / slides).

Mirrors the Office MCP's per-app structure: each app is a module with a
`register(mcp)` function and an `INSTRUCTIONS` string, listed in APPS.
"""

from fastmcp import FastMCP

from google_mcp import docs, sheets, slides

APPS: dict = {"docs": docs, "sheets": sheets, "slides": slides}


def build(app: str) -> FastMCP:
    module = APPS.get(app)
    if module is None:
        available = ", ".join(sorted(APPS)) or "(none yet — sheets is next)"
        raise SystemExit(f"unknown or unimplemented app {app!r}; available: {available}")
    mcp = FastMCP(f"google-{app}", instructions=getattr(module, "INSTRUCTIONS", None))
    module.register(mcp)
    return mcp
