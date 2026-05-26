---
name: office-mcp
description: Use when reading or editing the user's open Microsoft Word, Excel, or PowerPoint documents on macOS through the office-mcp tools (word_*, excel_*, ppt_*, run_applescript). Covers the read -> edit -> screenshot workflow, conventions (1-based indexes, [r,g,b] colors, the sheet param for cross-tab), the composite tools to prefer, and what AppleScript can't do.
---

# Driving Office with office-mcp

These tools edit the Microsoft Office documents the user **already has open** on
macOS, live (via Apple events). Each app is a separate MCP server (`word`,
`excel`, `powerpoint`); the connected server's tools are prefixed `word_*` /
`excel_*` / `ppt_*`. Every server also exposes `run_applescript` (escape hatch)
and `*_screenshot`.

## Core workflow: read -> edit -> see

1. **Read first** — `*_status` to confirm the app/document is open, then read the
   relevant content (`word_get_document_text`, `excel_read_range`,
   `ppt_read_slide`) so you act on real indices/values, not guesses.
2. **Edit** with the most specific tool.
3. **Verify visually** — call `*_screenshot` after any layout/format change and
   look at the result. Don't assume; confirm.

## Conventions

- Indexes (slides, shapes, paragraphs, tables, rows) are **1-based**.
- Colors are `[r, g, b]`, 0-255.
- Excel ranges are A1-style; most Excel tools take an optional `sheet` (name) —
  omit for the active sheet, pass it to work **across tabs**. Reads return 2-D lists.
- Prefer the **composite tools** for common multi-step tasks (fewer calls, more
  reliable — especially for smaller models):
  - `word_add_section(heading, body, level)` — styled heading + paragraph.
  - `excel_write_table(start_cell, values, header)` — values + header + borders + autofit.
  - `ppt_add_content_slide(title, bullets)` — a titled, bulleted slide.

## Per-app notes

- **Word**: `word_insert_text` appends at the end; `word_insert_at_cursor` /
  `word_replace_selection` act at the cursor/selection; `word_set_style` for
  headings; tables via `word_insert_table` + `word_set_table_cell`.
- **Excel**: formatting via `excel_format_range`; structure via
  `excel_insert_rows`/`_columns`, `excel_sort`, `excel_autofilter`,
  `excel_set_borders`, `excel_create_chart`; sheets via `excel_add_sheet` etc.
- **PowerPoint**: PowerPoint **can't report which shape is selected** — take a
  screenshot to see it. Reveal-on-click = `ppt_add_animation(..., exit=True)` on a
  cover shape. Add objects with `ppt_add_textbox` / `ppt_add_image`.

## Not possible (don't attempt — they error in the dictionary)

- PowerPoint: duplicating slides, creating charts, creating tables on slides.
- Word: comments.

Do these manually or tell the user. For anything else uncovered, use `run_applescript`.

## Permissions

First use of an app prompts a macOS **Automation** grant on the terminal app
(each app prompts once). The `*_screenshot` tools also need **Screen Recording**
(separate prompt, takes effect only after the terminal restarts). A denial returns
a clear "not authorized" error — guide the user to System Settings -> Privacy &
Security.
