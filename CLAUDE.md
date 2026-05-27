# docs-mcp workspace

A `uv` workspace of document co-editor MCPs — two packages, separate backends, one
shared design.

- `packages/office` — macOS Office (Word/Excel/PowerPoint) via AppleScript/osascript
  (`bridge.py`). See `packages/office/CLAUDE.md` for its conventions and the
  per-app AppleScript dictionary notes, and `packages/office/PLAN.md`.
- `packages/google` — Google Docs/Sheets/Slides via the Workspace REST APIs + OAuth
  (`auth.py`). Work in progress; see `packages/google/PLAN.md`.

## Shared conventions (both packages)

- Per-app servers: `<tool> <app>` launches a FastMCP server exposing only that
  app's tools. Add an app = a module with `register(mcp)` + an `INSTRUCTIONS`
  string, listed in that package's `server.APPS`.
- Semantic tools + a few composite ("thick") tools for common workflows + a raw
  escape hatch. Validated enum choices return the valid options on bad input.
- **Index/range-anchored editing** (paragraph index in Docs/Word, A1 ranges in
  Sheets/Excel, object IDs in Slides). Status/reads never launch the app.
- Verify every new tool **live** against the open app/doc; `tests/` only covers the
  OS/transport layer. Aim for a unified tool surface so one skill covers both
  backends.

## Commands

```bash
uv sync
uv run --directory packages/office office-mcp word                  # run a server
uv run --directory packages/office --with pytest pytest             # office tests
uv run --directory packages/google google-mcp sheets                # WIP
```

Commit one feature per commit; keep each package's `PLAN.md` tool surface current.
