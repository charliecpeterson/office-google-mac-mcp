"""Build a per-app MCP server.

Each launch registers only one app's tools, keeping the toolset small and
focused. New apps are added by writing a module with a `register(mcp)` function
and an `INSTRUCTIONS` string, and listing it in APPS.
"""

from fastmcp import FastMCP

from office_mcp import excel, powerpoint, word

APPS = {
    "word": word,
    "excel": excel,
    "powerpoint": powerpoint,
}


def build(app: str) -> FastMCP:
    module = APPS.get(app)
    if module is None:
        available = ", ".join(sorted(APPS))
        raise SystemExit(f"unknown app {app!r}; available: {available}")
    mcp = FastMCP(f"office-{app}", instructions=getattr(module, "INSTRUCTIONS", None))
    module.register(mcp)
    return mcp
