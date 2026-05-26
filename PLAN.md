# Office Coworker MCP — Plan

A from-scratch MCP server that lets a CLI agent (Claude Code, opencode, local
LLMs) act as a live coworker inside the Microsoft Office apps you already have
open on macOS — read what's on screen, edit the active document, and watch it
update live.

This replaces the cloned Windows `officemcp` project entirely. None of the old
COM/`pywin32` code survives the move to macOS.

## Goal

Match the functional capability of the in-app Claude Office add-in (read the
active document, understand it, write back into it) but driven from the
terminal over MCP, so it works with any agentic CLI instead of a single app's
chat panel.

## Mechanism

macOS has no COM. The way to drive a *running* Office app on Mac is **Apple
events** (what AppleScript speaks). Word, Excel, and PowerPoint each ship a
scriptable dictionary that mirrors the VBA object model. Apple events attach to
the already-running instance, so edits land in the document you have open and
you see them happen live.

Decisions settled:

- **No `appscript` Python library.** It's the most ergonomic option but compiles
  against system frameworks and chronically lags new Python releases — the exact
  sustainability risk to avoid. We use the built-in `osascript`, which ships
  with macOS and keeps working across upgrades.
- **No file-based editing** (python-docx / openpyxl / python-pptx). That edits
  files on disk, not the live app, and fights unsaved changes. Wrong tool for a
  live coworker.
- **No VBA.** Microsoft removed "do Visual Basic" from Office for Mac years ago,
  so macros can't be driven from outside. AppleScript dictionary only.

## Architecture

One package, **multiple launch modes** — one per app. Each launch spins up a
FastMCP server that registers only that app's tools, so a session carries a
small, focused toolset. This keeps context lean and, more importantly, makes
tool selection sharper (especially for local LLMs, which degrade as the toolset
grows).

- Single binary with an app argument: `office-mcp word` / `office-mcp excel` /
  `office-mcp powerpoint`.
- Shared code (the `osascript` bridge, error handling) is written once and
  reused — no per-app duplication.
- Each per-app server is self-contained: it includes that app's status,
  selection, read, and write tools plus the raw-script escape hatch.
- Transport: stdio (what Claude Code and opencode launch). SSE can come later if
  there's a reason.

### Tool philosophy

Semantic tools as the primary interface (`word_get_selection`,
`excel_read_range`, …), not a freeform code sandbox. Reasons: local LLMs are far
more reliable with a curated toolset; it's safer than `exec`-ing model output;
it's self-documenting. A single `run_applescript` escape hatch covers whatever
the semantic tools don't.

## Project structure

```
office-mcp/
├── pyproject.toml          # fastmcp dep, python >=3.12, office-mcp script
├── README.md
├── PLAN.md                 # this file
├── src/
│   └── office_mcp/
│       ├── __init__.py
│       ├── __main__.py     # parse app arg, run the right server
│       ├── server.py       # build FastMCP for an app, register its tools
│       ├── bridge.py       # the ONLY code that touches osascript
│       ├── word.py         # Word tools + AppleScript/JXA snippets
│       ├── excel.py
│       └── powerpoint.py
└── tests/
    └── test_bridge.py      # script-generation tests (no live app needed)
```

## The bridge (`bridge.py`)

The single OS-touching layer. Everything else builds script strings and hands
them here.

- `run_applescript(script) -> str` — runs via `osascript -e`/stdin, returns
  stdout, raises with the AppleScript error text on failure. Used for actions.
- `run_jxa(script) -> Any` — runs via `osascript -l JavaScript`, expects the
  script to `JSON.stringify` its result, parses JSON. Used for structured reads
  (e.g. a 2D Excel range comes back as a real nested list).
- Error translation: distinguish app-not-running, **not-authorized (TCC error
  -1743)**, and script errors, and surface clean messages instead of raw
  `osascript` noise.
- Read/status tools must **not** auto-launch the app — they report "not running"
  instead. Launching only happens on an explicit action.

## Tool surface (v1, shipped)

Each app server also exposes `<app>_screenshot()` — a PNG of the app window so
the model can see its own work — and `run_applescript(script)` as the escape hatch.

### Word
- `word_status` — running? active document name/path, whether text is selected
- `word_get_document_text` — full text (with optional length cap)
- `word_get_selection` — currently selected text
- `word_insert_text(text)` — append at end of document
- `word_replace_selection(text)`
- `word_find_replace(find, replace, match_case)`
- `word_apply_formatting(...)` — bold/italic/underline/size/color on the selection
- `word_get_outline` — heading structure
- `word_screenshot`

### Excel
- `excel_status` — running? active workbook, active sheet, all sheets, selection
- `excel_list_sheets`
- `excel_read_range(range, sheet)` — values as a 2D list (JXA → JSON)
- `excel_write_range(range, values, sheet)`
- `excel_set_cell(cell, value, sheet)`
- `excel_set_formula(cell, formula, sheet)`
- `excel_get_selection`
- `excel_screenshot`

### PowerPoint
- `ppt_status` — running? active presentation, slide count, current slide
- `ppt_list_slides` — text summary per slide
- `ppt_read_slide(index)` — shapes {shape, name, text}
- `ppt_add_slide(layout, position)`
- `ppt_set_text(slide, shape, text)`
- `ppt_get_current_slide`
- `ppt_screenshot`

## Client configuration

Per-project `.mcp.json` (Claude Code) — wire up only the app you need:

```json
{ "mcpServers": { "word": { "command": "uvx", "args": ["office-mcp", "word"] } } }
```

Spreadsheet project: swap to `"args": ["office-mcp", "excel"]`. opencode reads
the same shape. Need two apps at once? List two servers — each still carries
only its own tools.

## Permissions

First time a tool drives an app, macOS shows a one-time Automation prompt. It
attaches to the **parent terminal app** (Terminal / iTerm / VS Code), not Python
— so you grant "iTerm → control Microsoft Word" once under System Settings →
Privacy & Security → Automation. Each app prompts separately. Denial surfaces as
a clear "not authorized" error (TCC -1743), not a crash.

The `screenshot` tools additionally need **Screen Recording** permission for the
same terminal app (a separate prompt, and it only takes effect after the terminal
is restarted). Denial surfaces as a clear "not authorized" error.

## Scope

- **In, v1:** Word, Excel, PowerPoint.
- **Maybe later:** Outlook, OneNote (both scriptable on Mac).
- **Out:** Access, Publisher, Visio, Project (don't exist on macOS); Windows
  support; file-based editing; the old `RunPython` tool.

## Dependencies

`fastmcp`, `pyobjc-framework-Quartz` (window-id lookup for screenshots — a clean
binary wheel, no compile), Python ≥ 3.12. macOS only. No `pywin32`, no `PIL`, no
`appscript`.

## Testing

Apple events can't run without the live apps, so unit tests cover
script-generation logic (the strings we build) and the bridge's error parsing.
End-to-end verification is manual against open apps, app by app.

## Build order

1. ~~Scaffold package + `bridge.py` + `server.py` + `word_status`.~~ Done.
2. ~~Word reads (text, selection, outline).~~ Done.
3. ~~Word writes (insert, replace, find/replace, formatting).~~ Done — all 9 Word
   tools verified live.
4. ~~Excel module.~~ Done — all 8 Excel tools verified live.
5. ~~PowerPoint module.~~ Done — all 7 PowerPoint tools verified live.
6. README + config examples + polish.
7. ~~Vision: `<app>_screenshot`.~~ Done — verified end-to-end through the MCP.

## v2 roadmap — toward a first-class coworker

Priorities, learned from driving the tools on a real deck (the win: a
do → see → verify loop; the recurring friction: editing blind, and falling back
to `run_applescript` for common operations).

### Tier 1
- **Screenshot / vision.** Shipped. Closes the do → see → verify loop and makes
  every other edit checkable. (See "Vision findings".)
- **Selection-aware editing.** Act where the user is pointing. Today Word only
  appends at the end and PowerPoint addresses shapes by index — neither works off
  the live cursor / selected shape. Add: read the current selection (and selected
  shape in PowerPoint) and edit it in place. The most "coworker" interaction and
  currently the weakest.

### Tier 2 — promote escape-hatch operations to real tools
- PowerPoint: `ppt_add_click_reveal(slide, shape)` (the reveal-on-click build we
  did by hand), delete / duplicate / reorder slide, format / move / resize shape,
  add textbox / image.
- Word: insert-at-cursor, set a paragraph's style (we read the outline but can't
  set "Heading 2"), tables, comments.
- Excel: cell formatting + number formats, insert / delete rows-cols, sort/filter.

### Tier 3 — polish
- Safety: a backup/checkpoint tool. The sandbox blocks `/tmp`, but a copy can be
  saved beside the original before destructive edits.
- Richer structural reads: shape geometry, types, and existing animations —
  geometry + z-order is what let us identify a cover shape *without* vision, so
  feeding structure makes the model smarter between screenshots.
- Cross-app: pull an Excel chart / range into a slide.

### Principles
Keep per-app servers lean. Vision and selection are universal → shared core,
per-app wrappers. Default to non-destructive; bias reads toward returning
structure the model can reason over.

### Vision findings (learned from live runs)

- `screencapture -l<windowid> -o -x` grabs one app window cleanly, even when
  occluded, without moving it. The window id comes from Quartz
  (`CGWindowListCopyWindowInfo`, match `kCGWindowOwnerName`, largest area).
- PowerPoint's own PNG export is a dead end here: the Office **sandbox** blocks
  writes to normal paths (it can only write near the open file), and the
  "every slide / just this one" modal kills the AppleScript `save … as PNG`
  (returns success, writes nothing).
- Screenshots need **Screen Recording** permission (separate from Automation),
  and it only takes effect after the terminal restarts.

### Word dictionary notes (learned from live runs)

- Empty selection reads as `null`, not `""`.
- End of text: `end of content of (text object …)`.
- The `insert` command and `create range` ranges fail over osascript (-1708 /
  -10006). Insert via `type text selection text …` (after `end key selection
  move (a story item)`); replace via `set content of (text object of selection)`.
- A paragraph's `style` is a Word style object; its name is `name local`
  (e.g. "Heading 1"), not `name`.
- Find/replace: `execute find (find object of (text object of active document))
  … replace replace all wrap find find continue`.
- Select-all: `home key … move (a story item)` then `end key … extend (by
  selecting) move (a story item)`.

### Excel dictionary notes (learned from live runs)

- Unlike Word, JXA does everything cleanly, including writes via property
  assignment (`range.value = …`, `range.formula = …`).
- Range addressing uses **bracket** notation: `sheet.ranges['A1:B2']` and
  `workbook.worksheets['Sheet1']`. The function-call form `ranges('A1:B2')`
  fails (-1728).
- A range's address comes from `getAddress()` (the `address` *property* is the
  hyperlink address, unrelated). Values round-trip as JSON 2-D lists.
- Read tools return `Any`, so the value lands in the MCP text content, not the
  structured `.data` channel — fine for LLM clients.

### PowerPoint dictionary notes (learned from live runs)

- Slide text: shape -> `text frame` -> `text range` -> `content`, guarded by
  `has text frame` and `has text`.
- Add a slide with `make new slide at end of active presentation` (or `at after
  (slide N of active presentation)`). The `at end of slides of …` form and the
  JXA `make` both fail (-2710 / -1708).
- Current slide on screen: `slide of view of active window`, then `slide index`.
- Layouts are the `EPPSlideLayout` enum (`slide layout title only`, etc.);
  ppt_add_slide maps friendly names onto them.

## Assumptions taken from discussion

- Single binary with an app arg (not three separate commands).
- Word built first.

Flag either if wrong.
