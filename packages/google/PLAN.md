# Google Drive MCP — Plan

A from-scratch MCP server that lets a CLI agent (Claude Code, opencode, local
LLMs) act as a live co-editor of the user's Google Docs / Sheets / Slides — read
the document, edit it, and watch the change appear live in the open browser tab.

Same goal and design philosophy as the macOS Office MCP (`~/mcps/OfficeMCP`),
but a different mechanism. Read that project's `PLAN.md` first — most of the tool
design transfers; only the backend changes.

## Goal

The Office-MCP experience, for Google Workspace: a coworker that edits the doc you
have open and you see it update live. Cross-platform (any OS, any browser), driven
from a CLI agent rather than an in-app panel.

## Mechanism — the key difference from Office

Office MCP **attaches to the local running app** via Apple events and edits the
document on screen. Google has no local app to attach to: the document lives on
Google's servers. So this MCP edits the document **through the official Google REST
APIs** (Docs / Sheets / Slides / Drive), addressed by document ID. Because
Workspace docs are realtime-collaborative, **the user's open browser tab updates
live via Google's own sync** — exactly as if a collaborator were typing. The
"live in my tab" experience is real; it just arrives by a different path:

> API edits the server document → Google's collaboration sync reflects it in the
> open tab. (Not "the tool drives the tab.")

Decisions settled:

- **Official REST APIs, not browser/DOM automation.** Driving the Docs/Sheets web
  UI (Playwright, an extension) is brittle and strictly worse than the APIs.
- **The doc need not be "open."** The API edits it regardless; the tab is just a
  live viewer. To get the live experience, the user keeps the tab open.
- **Addressed by document ID/URL, not "frontmost app."** This is the main UX
  difference from Office — see "Document selection" below.

## What carries over from the Office MCP (reuse the spec)

The *design* transfers almost 1:1; only the implementation differs.

- Per-app servers (`docs` / `sheets` / `slides`), one package, launch mode
  (`google-mcp docs`), stdio transport.
- Semantic tools + a few composite ("thick") tools for common workflows + a raw
  escape hatch (here: a raw `batchUpdate` request runner per API).
- Server `instructions` per app (portable usage context) + a usage skill.
- Validated enum choices that return the valid options on bad input.
- **Index/range-anchored editing.** The paragraph-anchored model we built for Word
  maps perfectly — the Docs API is *natively* index-based, Sheets is A1-range
  based (like our Excel tools), Slides is object-ID based.
- do → read-structure → verify. Reads return full structured content, so
  screenshots are rarely needed (unlike Office); export-to-image is optional.

## Relationship to the official Google Drive MCP (compose, don't merge)

Google ships a hosted, managed **Drive MCP** (developer preview,
`drivemcp.googleapis.com`) that does Drive *file management*: `search_files`,
`list_recent_files`, `read_file_content`, `get_file_metadata`, permissions,
copy/create. It does **not** edit the *content* of Docs/Sheets/Slides — that's
exactly our job. So they compose:

- **Official Drive MCP** (or Anthropic's Google Drive connector) = file discovery +
  reading files for context ("get me the relevant files").
- **This MCP** = deep content co-editing of a specific document.

Consequence: **we don't build Drive search/recent at all.** Lean on the official
Drive MCP / connector for discovery and context; this MCP only needs "edit the
document at this URL/ID." The Drive section below shrinks to just active-document
selection + optional export-to-image. Run both servers in a session.

## Packaging — separate packages, eventually one monorepo

Decision: separate *packages* (never one tangled package), and when both Office and
Google exist, house them in **one monorepo (a `uv` workspace)** rather than isolated
repos — because the goal is to **co-evolve the tool surface** (mirror a Word tool
improvement into Google Docs, and vice versa). A monorepo makes that a single
side-by-side PR; isolated repos make it a two-clone, two-PR, drift-prone dance.

Structure when built:
```
docs-mcp/                # monorepo, uv workspace (renamed/scoped for both backends)
  packages/core/         # FastMCP plumbing, enum validation, shared conventions
  packages/office/       # macOS, pyobjc, AppleScript bridge  (move OfficeMCP in)
  packages/google/       # cross-platform, google clients, OAuth
  skills/office-mcp/      # one shared skill
  PLAN.md  README.md
```

Notes / constraints:
- A `uv` workspace keeps deps per-package, so the macOS-only `pyobjc` stays in
  `office` and the Google clients in `google` — "one repo" ≠ "one dependency blob".
- No common *engine*: Office is AppleScript, Google is REST. ~90% of each package is
  backend-specific and can't be shared. The monorepo makes parallel changes
  *visible and co-committable*, not automatic; what truly shares code is `core`
  (plumbing/helpers) and the composite-tool *ideas* + the skill.
- Aim for a unified interface: a `doc_insert_paragraph` that behaves the same on
  either backend, so one skill and one mental model cover both.
- "Load either Word *or* Google Docs in a session" already works via the per-app
  server + `.mcp.json` model regardless of layout — list whichever servers you want.
- **Don't reorganize the existing Office repo yet.** It's a working, published,
  macOS-scoped repo; restructuring it for not-yet-built Google code is premature
  churn. Do the monorepo move when there's real Google code to co-evolve with, and
  rename/rescope the repo deliberately at that point.

## What's new / different (Google-specific)

- **Auth: OAuth 2.0 + a Google Cloud project.** Enable the Docs/Sheets/Slides/Drive
  APIs, configure a consent screen, generate desktop-app credentials, store the
  token. One-time setup, heavier than Office's TCC prompt. Service-account +
  domain-wide delegation is the Workspace-org option. Scopes: `documents`,
  `spreadsheets`, `presentations`, and `drive.file` (or `drive` for search).
- **Document selection.** The agent needs to know *which* doc. Options (support
  more than one): accept a URL/ID; a `drive_recent` / `drive_search` to list and
  pick; an `set_active_document(id)` that the per-app tools default to. There's no
  "frontmost" — this is the cost of the API model.
- **More capable, fewer dead-ends than AppleScript.** Notably Docs **comments**
  (which Word's AppleScript can't do) and Slides **duplicate/reorder** (which
  PowerPoint's AppleScript can't) are first-class in the APIs.
- **Index gotcha (Docs).** Docs character indices *shift* as you insert/delete, so
  a multi-edit `batchUpdate` must be ordered carefully (apply edits back-to-front,
  or recompute indices between requests). Sheets (A1) and Slides (object IDs) don't
  have this problem.
- **Quotas / rate limits.** Per-minute API quotas; batch related edits into one
  `batchUpdate` where possible (also atomic).

## Architecture

```
google-drive-mcp/
├── pyproject.toml          # fastmcp + google-api-python-client + google-auth-oauthlib
├── README.md
├── PLAN.md                 # this file
├── src/google_mcp/
│   ├── __main__.py         # google-mcp <app>
│   ├── server.py           # APPS registry → per-app FastMCP + instructions
│   ├── auth.py             # OAuth flow + cached token; builds API service clients
│   ├── drive.py            # document selection: search/recent/active, export
│   ├── docs.py
│   ├── sheets.py
│   └── slides.py
└── tests/
```

Language: **Python + FastMCP**, to match the Office MCP's conventions, skill, and
shared design (the alternative is forking the Node-based `a-bonus/google-docs-mcp`
— see below). `auth.py` is the analog of Office's `bridge.py`: the one module that
holds credentials and builds the Google service clients; every tool calls the API
through it.

## Tool surface (target)

Each app server also exposes its raw `*_batch_update(requests)` escape hatch.

### Document selection (minimal — discovery is the official Drive MCP's job)
- `set_active_document(id_or_url)` / `get_active_document` — the default target the
  per-app tools edit (accepts a pasted Docs/Sheets/Slides URL)
- `export(id, format)` — PDF/PNG for an occasional visual check
- (No search/recent here — use the official Google Drive MCP or Anthropic's Drive
  connector for finding files and pulling them in as context.)

### Docs (index/range-anchored, like Word)
- `docs_get_structure` — paragraphs/elements with indices and styles
- `docs_get_text(range?)`
- `docs_insert_paragraph(text, after/before index, style)`,
  `docs_replace_range`, `docs_delete_range`
- `docs_apply_style` (headings), `docs_format_text` (bold/size/color)
- `docs_find_replace`, `docs_insert_table` / `docs_fill_table`,
  `docs_insert_image`, `docs_add_comment` (a Google advantage)

### Sheets (A1, like Excel)
- `sheets_list_tabs`, `sheets_add_tab` / `rename` / `delete` / `activate`
- `sheets_read_range`, `sheets_write_range`, `sheets_set_formula`
- `sheets_format_range`, `sheets_set_borders`, `sheets_sort`, `sheets_autofilter`
- `sheets_create_chart`, `sheets_conditional_format`, `sheets_data_validation`
- composite: `sheets_write_table` (values + header + borders)

### Slides (the gap a-bonus doesn't cover)
- `slides_list` / `slides_read(index)`
- `slides_add` / `slides_duplicate` / `slides_delete` / `slides_move` (all clean here)
- `slides_add_textbox` / `slides_add_image` / `slides_set_text`
- `slides_format_shape` (fill/border/font), `slides_set_notes`
- composite: `slides_add_content_slide(title, bullets)`

## Relationship to `a-bonus/google-docs-mcp`

That project is a mature API-based Google MCP: Docs, Sheets, Drive, Gmail, Calendar
(no Slides); official REST APIs; OAuth2 / OAuth2.1 / service account; needs a GCP
project. It already covers Docs + Sheets richly (incl. Docs comments, Sheets
charts/validation/conditional formatting).

So the real decision is **fork/reuse vs. greenfield**:

- **Reuse it** for Docs + Sheets (don't rebuild what works), and build a separate
  Slides MCP to fill its gap. Fastest path to coverage. Cost: it's Node, so it
  won't share the Office MCP's Python design/skill, and you'd run two codebases.
- **Greenfield** all three in Python, matching the Office MCP's design exactly.
  More work up front, but one consistent design, one skill, and a **shared semantic
  tool surface** with Office (same tool names/semantics, two backends) so an agent
  and skill work across both. This is the cleaner long-term architecture.

Recommendation: prototype against `a-bonus` first to learn the APIs and confirm the
"live in the tab" experience with minimal effort; decide greenfield-vs-fork after,
based on how much the shared-design payoff matters.

## Build order (mirrors the Office MCP's incremental, verify-live approach)

1. `auth.py` + `set_active_document` (paste a URL). Get OAuth working and confirm
   an API edit shows up live in an open tab. (File discovery is the official Drive
   MCP's job, not ours.)
2. **Sheets** first — cleanest API (A1 ranges), most like our Excel tools.
3. **Docs** — the index-shifting model; reuse the paragraph-anchored design.
4. **Slides** — the new ground; the APIs support duplicate/reorder cleanly.
5. Composite tools + server instructions + skill.

## Scope

- **In:** Docs, Sheets, Slides editing as a co-editor.
- **Out:** Gmail / Calendar (not a doc co-editor concern; `a-bonus` covers them if
  needed); browser/DOM automation; offline file manipulation.

## Open questions

- Greenfield Python vs. fork the Node `a-bonus` server (the shared-design question).
- Best "active document" UX — paste URL each time, vs. recent-files picker, vs. a
  persisted active-doc setting.
- Auth distribution for other users (per-user GCP project is friction; a published
  OAuth app with verification is the smoother but heavier path).
- Whether to converge on a shared semantic tool surface with the Office MCP now or
  later.
