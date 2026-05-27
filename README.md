# office-mcp

An MCP server that lets a CLI agent (Claude Code, opencode, local LLMs) work as a
live coworker inside the Microsoft Office apps you have open on macOS — read the
active document, edit it, see it via a screenshot, and watch it update live.

macOS only. It drives the running apps through Apple events (`osascript`) and
captures windows with the built-in `screencapture`; nothing compiles locally. See
[PLAN.md](PLAN.md) for the design, the per-app dictionary notes, and the roadmap.

## Per-app servers

Each launch exposes one app's tools, keeping the toolset small and focused (which
also makes tool selection sharper, especially for local models):

```
office-mcp word
office-mcp excel
office-mcp powerpoint
```

## What it can do

72 tools across the three apps. Every app also has a `*_screenshot` (a PNG of the
window, even when occluded) and a `run_applescript` escape hatch.

- **Word (25)** — read document / selection / heading outline / per-paragraph
  structure / page-word stats; **paragraph-anchored editing** (insert before/after
  paragraph N, replace, delete) plus end-append, cursor-insert, and replace-selection;
  find-and-replace; font formatting; paragraph styles (headings); tables (insert +
  bulk-fill + read/write cells); inline pictures; floating text boxes.
- **Excel (26)** — read/write ranges (2-D) and cells; formulas and array formulas;
  the current selection; cell & number formatting; borders; insert/delete rows and
  columns; autofit; sort; autofilter; charts; sheet management (add/delete/rename/
  activate) for cross-tab work.
- **PowerPoint (21)** — list/read slides and the current slide; add/delete/move
  slides; set text by index or at the selection; add text boxes and images; move,
  resize, and format shapes (fill/border/font); animations (incl. reveal-on-click);
  speaker notes.

Each app also has a few **composite tools** that fold a common multi-step workflow
into one call — `word_add_section`, `excel_write_table`, `ppt_add_content_slide` —
which keeps things reliable for smaller models. And each server ships usage
`instructions` (workflow, conventions, gotchas) that any MCP client receives on
connect; there's also a Claude Code skill at `.claude/skills/office-mcp/`.

## Configure (Claude Code)

Copy `.mcp.json.example` to `.mcp.json` and set the `--directory` path to your
checkout. It wires up all three apps (drop the ones you don't need — each server
carries only its own tools):

```json
{
  "mcpServers": {
    "word": {
      "command": "uv",
      "args": ["run", "--directory", "/path/to/office-mcp", "office-mcp", "word"]
    }
  }
}
```

`.mcp.json` is gitignored since the path is machine-specific. opencode reads the
same shape. Once published to PyPI, the command simplifies to
`uvx office-mcp <app>` and no path is needed.

## Permissions

The first time a tool drives a given Office app, macOS shows a one-time
**Automation** prompt (each app — Word, Excel, PowerPoint — prompts separately). It
attaches to the **parent terminal app** (Terminal / iTerm / VS Code), not Python —
grant "<your terminal> → control Microsoft Word" under System Settings → Privacy &
Security → Automation. If denied, tools return a clear "not authorized" error.

The `*_screenshot` tools additionally need **Screen Recording** for the same
terminal app (a separate prompt under Privacy & Security → Screen Recording), and
it only takes effect after you restart the terminal.

## Develop

```bash
uv run office-mcp word     # run the Word server over stdio
uv run pytest              # bridge tests (no Office app needed)
```

Tools are verified live against the open apps; `tests/` covers the OS bridge
(osascript mechanics and error translation), which is the only part testable
without Office running.
