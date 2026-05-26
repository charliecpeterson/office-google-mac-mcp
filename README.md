# office-mcp

An MCP server that lets a CLI agent (Claude Code, opencode, local LLMs) work as
a live coworker inside the Microsoft Office apps you have open on macOS — read
the active document, edit it, and watch it update live.

macOS only. It drives the running apps through Apple events (`osascript`), so
there are no native build dependencies. See [PLAN.md](PLAN.md) for the design
and roadmap.

## Per-app servers

Each launch exposes one app's tools, keeping the toolset small and focused:

```
office-mcp word
office-mcp excel
office-mcp powerpoint   # not yet implemented
```

Word and Excel are implemented; PowerPoint follows.

## Configure (Claude Code)

Per-project `.mcp.json`, wiring up only the app you need:

```json
{ "mcpServers": { "word": { "command": "uvx", "args": ["office-mcp", "word"] } } }
```

opencode reads the same shape.

## Permissions

The first time a tool drives Word, macOS shows a one-time Automation prompt. It
attaches to the **parent terminal app** (Terminal / iTerm / VS Code), not Python
— grant "<your terminal> -> control Microsoft Word" under System Settings ->
Privacy & Security -> Automation. If denied, tools return a clear "not
authorized" error.

## Develop

```bash
uv run office-mcp word     # run the Word server over stdio
uv run pytest              # bridge tests (no Office app needed)
```
