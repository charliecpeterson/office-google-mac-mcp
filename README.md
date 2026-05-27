# docs-mcp — document co-editor MCPs

A monorepo of MCP servers that let a CLI agent (Claude Code, opencode, local LLMs)
act as a live co-editor of the documents you have open — read, edit, and watch
them update live.

- **[packages/office](packages/office)** — Microsoft Word / Excel / PowerPoint on
  macOS, driven through Apple events. Working, 72 tools, published. macOS only.
- **[packages/google](packages/google)** — Google Docs / Sheets / Slides, through
  the Google Workspace REST APIs (cross-platform; edits show up live in your open
  browser tab via Google's sync). Work in progress — Sheets first.

Both share a design — per-app servers, semantic + a few composite ("thick") tools,
the read → edit → verify workflow, and one [skill](.claude/skills/office-mcp) — so
the two backends feel like one coworker. The engines differ (AppleScript vs. REST
API); the tool surface is kept aligned. See each package's `PLAN.md`.

## Layout

A `uv` workspace; each package installs and runs independently:

```bash
uv sync                                                # workspace deps
uv run --directory packages/office office-mcp word     # Word server (stdio)
uv run --directory packages/google google-mcp sheets   # WIP
```

Wire whichever apps you need into your MCP client — copy `.mcp.json.example` to
`.mcp.json` and set the paths. Each server carries only its own app's tools.
