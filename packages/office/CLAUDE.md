# office-mcp

A macOS-only MCP server that drives the Microsoft Office apps the user already
has open (Word / Excel / PowerPoint) via Apple events, so a CLI agent can read
and edit live documents. See `PLAN.md` for the full design, the per-app
AppleScript dictionary notes, and the roadmap — read it before extending a module.

## Architecture

- `office-mcp <app>` (`__main__.py` → `server.py`) builds a FastMCP server that
  registers **only that one app's** tools. Per-app servers keep each session's
  toolset small. Add an app by writing a module with `register(mcp)` and listing
  it in `server.APPS`.
- `bridge.py` is the **only** module that touches the OS. Everything else builds
  script strings and calls it. Don't shell out elsewhere.
- One module per app: `word.py`, `excel.py`, `powerpoint.py`. Each is
  self-contained (its own status/read/write tools, `*_screenshot`, and a
  `run_applescript` escape hatch). Structure: docstring → enums/helpers → script
  constants → `register(mcp)` with `@mcp.tool` functions.

## Conventions (follow these when adding tools)

- **Reads → `bridge.run_jxa(script)`**: the JXA script ends in a
  `JSON.stringify(...)` expression; the parsed value comes back. JXA gives real
  nested lists/dicts. Polymorphic reads are typed `Any` (value lands in the MCP
  text content).
- **Writes with user text → `bridge.run_applescript(script, *args)`** where the
  script is `on run argv ... (item N of argv)`. **Never interpolate user text or
  sheet names or file paths into a script string** — pass them as args (avoids
  quote-escaping bugs). Numbers/bools/enum choices are safe to inline.
- Excel is JXA end-to-end (including writes via property assignment:
  `range.value = …`, `range.formulaArray = …`). Word/PowerPoint writes need
  AppleScript. Excel addresses ranges with **bracket** notation
  (`sheet.ranges['A1:B2']`), never `ranges('A1:B2')`.
- Colors are `[r, g, b]` integer lists (0-255).
- Friendly enum choices (styles, layouts, chart types, weights) map through a
  module-level dict and are validated (raise `ValueError` listing the options).
- Status/read tools must **not** launch the app — check `.running()` first.
- Verify every new tool **live** against the open app, and confirm visually with
  the app's `*_screenshot`. `uv run pytest` only covers the OS bridge.

## Hard constraints

- **TCC permissions**: Automation (per app, prompts once, attaches to the parent
  terminal app) and Screen Recording (for `*_screenshot`, separate prompt, needs a
  terminal restart). Denials surface as `bridge.NotAuthorized`.
- **Office sandbox**: the apps can only write near the open document. Writing a
  PNG/export to `/tmp` silently no-ops — that's why screenshots go through
  `screencapture`, not the apps' own export.
- **Not scriptable (confirmed dead-ends, see PLAN.md)**: PowerPoint slide
  duplicate / charts / tables-on-slides, and Word comments. Don't retry these —
  they error in the dictionary, not in our code. They'd need GUI scripting.

## Commands

```bash
uv run office-mcp word          # run a server over stdio (word|excel|powerpoint)
uv run pytest                   # bridge tests (no Office app needed)
sdef "/Applications/Microsoft Word.app"   # dump an app's AppleScript dictionary
```

Commit one feature per commit (see the git history). Update `PLAN.md`'s tool
surface and dictionary notes in the same change.
