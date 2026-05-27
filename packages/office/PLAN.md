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

### Small-model support

Three things help thin/local models run these tools as well as large ones:

1. **Per-app servers + server `instructions`.** Each server carries only one
   app's tools and ships a FastMCP `instructions=` string (the module-level
   `INSTRUCTIONS`) with the workflow, conventions, and dead-ends — portable to any
   client (Claude Code, opencode, local). A Claude Code skill at
   `.claude/skills/office-mcp/` adds richer on-demand recipes (copy to
   `~/.claude/skills/` to use across projects).
2. **Composite ("thick") tools** that fold a multi-step workflow into one call, so
   the model supplies intent, not orchestration: `word_add_section`,
   `excel_write_table`, `ppt_add_content_slide`. A deterministic composite result
   also needs less visual re-checking (a weak spot for small models). Keep these
   few — primitives stay underneath for flexibility.
3. **Validated enum choices** (styles, layouts, chart types, weights, colors)
   return the valid options on bad input, so a model self-corrects.

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

Composite tools (one call, multi-step — preferred for common tasks and small
models): `word_add_section`, `excel_write_table`, `ppt_add_content_slide`.

### Word
- `word_add_section(heading, body, level)` — styled heading + body (composite)
- `word_status` — running? active document name/path, whether text is selected
- `word_get_document_text` — full text (with optional length cap)
- `word_get_selection` — currently selected text
- `word_insert_text(text)` — append at end of document
- `word_insert_at_cursor(text)` — insert at the cursor, before any selection
- `word_replace_selection(text)`
- `word_find_replace(find, replace, match_case)`
- `word_apply_formatting(...)` — bold/italic/underline/size/color on the selection
- `word_get_outline` — heading structure
- `word_get_paragraphs(max_chars)` — {index, style, text, table?} per paragraph
- `word_get_paragraph(index)` — one paragraph's full untruncated text
- `word_get_stats` — {pages, words, paragraphs}
- `word_insert_paragraph(text, after, before, style)` — mid-document insert by index
- `word_replace_paragraph(index, text)` / `word_delete_paragraph(index)`
- `word_set_style(style, paragraph)` — normal / title / heading 1-9
- `word_insert_table(rows, columns, data)` — empty table, or sized-and-filled from `data`
- `word_fill_table(rows, table)` — bulk-fill an existing table in one call
- `word_set_table_cell(row, column, text, table)` / `word_get_table_cell(...)`
- `word_insert_picture(path)` — inline picture at the cursor
- `word_add_textbox(text, left, top, width, height)` — floating text box
- `word_screenshot`

### Excel
- `excel_write_table(start_cell, values, header)` — write + header + borders + autofit (composite)
- `excel_status` — running? active workbook, active sheet, all sheets, selection
- `excel_list_sheets`
- `excel_read_range(range, sheet)` — values as a 2D list (JXA → JSON)
- `excel_write_range(range, values, sheet)`
- `excel_set_cell(cell, value, sheet)`
- `excel_set_formula(cell, formula, sheet)`
- `excel_set_array_formula(range, formula, sheet)` — CSE/array formula
- `excel_get_selection`
- `excel_set_selection(value)` — set every cell in the selection
- `excel_format_range(range, sheet, bold, italic, size, font_color, fill_color, number_format)`
- `excel_insert_rows(at_row, count, sheet)` / `excel_delete_rows(...)`
- `excel_insert_columns(at_col, count, sheet)` / `excel_delete_columns(...)`
- `excel_autofit(range, sheet)` — auto-fit column widths
- `excel_add_sheet` / `excel_delete_sheet` / `excel_rename_sheet` / `excel_activate_sheet`
- `excel_sort(range, key_column, ascending, has_header, sheet)`
- `excel_set_borders(range, weight, color, sheet)` — edges + inner grid
- `excel_autofilter(range, sheet)`
- `excel_create_chart(range, chart_type, sheet)` — column/bar/line/scatter/pie/area
- `excel_screenshot`

### PowerPoint
- `ppt_add_content_slide(title, bullets)` — titled, bulleted slide (composite)
- `ppt_status` — running? active presentation, slide count, current slide
- `ppt_list_slides` — text summary per slide
- `ppt_read_slide(index)` — shapes {shape, name, text}
- `ppt_add_slide(layout, position)`
- `ppt_set_text(slide, shape, text)`
- `ppt_get_current_slide`
- `ppt_get_selection` — type / slide / selected text
- `ppt_set_selected_text(text)` — edit where the user is working
- `ppt_get_notes(slide)` / `ppt_set_notes(slide, text)` — speaker notes
- `ppt_add_textbox(slide, text, left, top, width, height)` — returns new shape index
- `ppt_add_image(slide, path, left, top, width, height)` — returns new shape index
- `ppt_set_shape_position(slide, shape, left, top, width, height)` — move/resize
- `ppt_format_shape(slide, shape, fill_color, border_color, border_weight)`
- `ppt_format_text(slide, shape, bold, italic, underline, size, color)`
- `ppt_add_animation(slide, shape, effect, trigger, exit)` — e.g. reveal-on-click
- `ppt_delete_slide(slide)` / `ppt_move_slide(slide, before)`
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
- **Selection-aware editing.** Shipped: `word_insert_at_cursor`,
  `excel_set_selection`, `ppt_get_selection` + `ppt_set_selected_text`. Act where
  the user is pointing. Caveat: PowerPoint reports selection *type* and selected
  *text* but won't reveal the selected *shape* (`shape range of selection` reads
  empty via both JXA and AppleScript) — a screenshot shows the handles instead.

### Tier 2 — promote escape-hatch operations to real tools
- PowerPoint: shipped `ppt_add_animation` (reveal-on-click etc.),
  `ppt_set_shape_position`, `ppt_format_shape` (fill/border), `ppt_format_text`
  (font), `ppt_add_textbox`, `ppt_add_image`, `ppt_delete_slide`, `ppt_move_slide`,
  `ppt_get_notes` / `ppt_set_notes`. **Not scriptable** in PowerPoint's dictionary
  (confirmed): slide duplicate (no `duplicate`/`paste`, only `copy object`), charts
  (no chart class), and tables on slides (`make new table` → -2710). These need GUI
  scripting or manual steps.
- Word: shipped insert-at-cursor, `word_set_style`, `word_insert_table` (with
  `data`), `word_fill_table`, `word_set_table_cell` / `word_get_table_cell`,
  `word_insert_picture`, `word_add_textbox` (floating), the structured reads
  `word_get_paragraphs` / `word_get_paragraph` / `word_get_stats`, and
  paragraph-anchored editing (`word_insert_paragraph` after/before N,
  `word_replace_paragraph`, `word_delete_paragraph`) — the co-editor model, added
  from dogfooding feedback. Comments aren't scriptable. Still to do: table/cell
  formatting; anchoring by matched text rather than paragraph index.
- Excel: shipped formatting, rows/cols, autofit, sheet management (add/delete/
  rename/activate), sort, borders, autofilter, charts, and array formulas
  (`excel_set_array_formula`). Cross-sheet refs work via the `sheet` param and
  `Sheet!A1` formulas. Still to do: pivot tables, conditional formatting.

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
- Set a paragraph/selection style with the `WdBuiltinStyle` enum: `set style of
  selection to style heading2` (no space: `style heading2`, not `style heading 2`).
- Insert a table: `make new table at <text range> with properties {number of rows,
  number of columns}`.
- Comments are not a scriptable class (no comment element / make), only
  `delete all comments`.
- Table cells: `get cell from table (table N of active document) row R column C`,
  then `content of text object of` the cell (reads end with `\\r\\x07` markers —
  strip them).
- Insert a picture: `make new inline picture at <text range> with properties
  {file name:"/path"}`.
- Floating text box: `make new text box at active document with properties
  {left position, top, width, height}`, then set `content of text range of text
  frame`. Bulk `delete every shape` fails (-1708) — delete shapes by index.
- `word_insert_text` ends with `type paragraph` so appended content is its own
  paragraph — otherwise a following heading glues onto it and inherits its style
  (a real bug found dogfooding; see the session notes that drove the read tools,
  bulk table fill, and stats below).
- Stats: `compute statistics <doc> statistic statistic pages|words|paragraphs`.
  Find/replace honors Word codes in the text (`^p` paragraph, `^l` line break,
  `^t` tab) but caps the replacement at 255 chars (raises a clear error past that).
- Editing is paragraph-anchored. Mid-document insert: `create range` at `start of
  content` / `end of content` of paragraph N, `select` it, then `type text` +
  `type paragraph` (plain `insert`/setting a collapsed range's content fail).
  Replace a paragraph with `set content of (text object of paragraph N) to (text
  & return)` — the `& return` keeps the paragraph mark (omitting it merges into the
  next paragraph, the bug found dogfooding). Delete: select the paragraph's
  start..end range and set its content to "". Table cells carry a `\x07` marker.

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
- Formatting is JXA too: `range.fontObject.{bold,italic,fontSize,color}`,
  `range.interiorObject.color`, `range.numberFormat`. Colors are `[r, g, b]`
  lists (0-255). Font size is `fontSize` / `font size`, not `size`.
- Rows/cols use AppleScript `insert into range` / `delete range` on a row range
  ("5:7") or column range ("C:E"); `autofit (entire column of range …)`. The sheet
  name is passed via argv to avoid escaping.
- Sheets: `make new worksheet at end of active workbook`, `activate object
  worksheet …`, `set name of worksheet …`, `delete worksheet …`.
- Sort: `sort <range> key1 <cell-in-column> order1 sort ascending|descending
  header header yes|no`. Borders: `get border <range> which border <XlBordersIndex>`,
  then `line style` (continuous), `weight` (`border weight thin`…), `color`.
- Charts: select the range first, then `make new chart object`, then `set chart
  type of chart of …` (XlChartType). There's no clean SetSourceData, so the
  selection is what seeds the chart. Filter: `autofilter range <range>`.
- Array formula: range `formula array` is settable (JXA `range.formulaArray = …`).

### PowerPoint dictionary notes (learned from live runs)

- Slide text: shape -> `text frame` -> `text range` -> `content`, guarded by
  `has text frame` and `has text`.
- Add a slide with `make new slide at end of active presentation` (or `at after
  (slide N of active presentation)`). The `at end of slides of …` form and the
  JXA `make` both fail (-2710 / -1708).
- Current slide on screen: `slide of view of active window`, then `slide index`.
- Layouts are the `EPPSlideLayout` enum (`slide layout title only`, etc.);
  ppt_add_slide maps friendly names onto them.
- Selection: `selection type` (none/slides/shapes/text) and `text range of
  selection` read fine, but `shape range of selection` reads empty even when the
  type is "shapes" — so we can't enumerate selected shapes. Clicking a text
  placeholder enters "text" mode (cursor), not "shapes" mode. Editing via the
  selection's text range works (replace highlighted / insert at cursor).
- Shape position is `left position` / `top` / `width` / `height` (note: `top`,
  not `top position` — the latter reads null).
- Speaker notes: the "Notes Placeholder" shape on `notes page of slide N` (its
  index varies, so find it by name). JXA can set its `content` directly.
- Animation: `add effect (main sequence of timeline of slide) for <shape> fx
  <MsoAnimEffect> trigger <MsoAnimTriggerType>`, then `set exit animation` for an
  exit effect. The trigger isn't readable back (write-only), so verify in a show.
- `move slide X to before slide Y` works; `duplicate slide X` returns -50.
- Add shapes: `make new text box at <slide> with properties {left position, top,
  width, height}` (then set its text), `make new picture at <slide> with properties
  {file name, left position, top, …}`. The new shape is the last one (its index =
  count of shapes).
- Dead ends (dictionary doesn't support): slide `duplicate`/`paste` (only
  `copy object`), `chart` class, and `make new table` (-2710). These need GUI
  scripting (System Events) or manual steps — not done.
- Colors are integer `{r, g, b}` lists (0-255): fill = `fore color of fill format`,
  border = `fore color of line format` (+ `line weight`), text = `font color of
  font of text range`. The PowerPoint `font` has boolean `underline` (unlike Word's
  enum).

## Assumptions taken from discussion

- Single binary with an app arg (not three separate commands).
- Word built first.

Flag either if wrong.
