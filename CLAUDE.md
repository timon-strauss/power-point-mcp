# CLAUDE.md — power-point-mcp

Instructions for any future Claude (or other agent) session that opens
this repository. The user's global CLAUDE.md and the project-scope file
in the parent directory take precedence over this file; this file is for
project-internal conventions.

## What this project is

A small MCP server, written in Python and exposed via `mcp[cli]`, that
lets Claude read, create, and modify exactly **one** PowerPoint file per
process. The bound target and (optional) template are picked at server
startup via environment variables.

The full feature set, scope, and out-of-scope list live in
`docs/iteration-1-plan.md` and `docs/iteration-2-plan.md`. Read those
before adding features.

## Hard rules for working on this repo

1. **Never break the single-file boundary.** Every file operation must
   route through `power_point_mcp.security.assert_within_target`. No
   tool may accept a path argument. If you add a tool that needs to
   touch the filesystem, route it through the same helper.
2. **Never invent slide text.** Every byte of text inserted into a
   placeholder, title, or shape must come from a tool argument supplied
   by the caller. The server never derives titles, summaries, or body
   text on its own. If you write a tool that fills a placeholder with
   anything other than a caller-supplied value, you have broken the
   contract.
3. **No network code.** No `requests`, `httpx`, `urllib.request`, or
   similar imports anywhere in `src/power_point_mcp/`. The hygiene test
   `tests/test_no_network_imports.py` enforces this.
4. **No global installs.** All dependencies go through `uv` into the
   project's local `.venv`. Never tell the user to `pip install` or
   `pipx install` anything.
5. **Tests must stay green.** Run `uv run pytest -q` before committing
   any code change. The full suite finishes in under a second.
6. **Pin the python-pptx major.** `delete_slide` and `reorder_slide`
   poke `prs.slides._sldIdLst`, a private API. The `<2` upper bound in
   `pyproject.toml` is load-bearing — do not loosen it without
   replacing the workaround.

## Layering

```
config.py        ← parses env vars into a frozen ServerConfig
   ↓
security.py      ← assert_within_target — the single chokepoint
   ↓
pptx_ops.py      ← pure functions over python-pptx; no MCP imports
   ↓
server.py        ← thin FastMCP wrappers; opens, calls helper, saves
   ↓
__main__.py      ← CLI entry point (`uv run power-point-mcp`)
```

When adding a tool: write the pure helper in `pptx_ops.py` first, test
it in isolation, then expose it through `server.py`. This keeps
`server.py` boring and testable.

## Workflow conventions

The orchestrator drives this project in iterations. One iteration =
plan → write → test → review → save (commit). Every iteration:

- Adds a `docs/iteration-N-plan.md` describing scope, signatures,
  testing, out-of-scope, and acceptance criteria.
- Bumps `version` in `pyproject.toml` and `__version__` in
  `src/power_point_mcp/__init__.py` (semver).
- Appends an entry to `changelog.md` under a `[X.Y.Z] — YYYY-MM-DD`
  heading. Newest entry on top.
- Commits all changes in a single commit with a `feat:` / `fix:` /
  `chore:` prefix.

## What is in scope

- Anything `python-pptx` can express: slides, shapes, placeholders,
  layouts, masters, text frames, runs.
- Reading and writing the bound `.pptx`.
- Honouring a user-supplied template.

## What is out of scope (hard limits)

- Touching any file other than `PPTX_TARGET`.
- Network calls of any kind.
- Cleaning up orphan slide parts left by `delete_slide` (the file
  stays slightly larger than a clean delete; PowerPoint tolerates it).
- `.ppt` (binary) format support — `python-pptx` does not handle it.
- PDF export, image rasterisation, video embedding.
- Concurrent writes from multiple clients (the server assumes serial
  use; concurrency is the client's problem to coordinate).

## Running the project

```bash
uv sync --group dev          # one-time
PPTX_TARGET=/abs/path.pptx uv run power-point-mcp
```

For tests: `uv run pytest -q`.

## Files an agent should know

- `src/power_point_mcp/server.py` — tool registrations.
- `src/power_point_mcp/pptx_ops.py` — slide-level operations.
- `src/power_point_mcp/security.py` — the boundary helper.
- `tests/conftest.py` — fixtures for a synthetic template.
- `docs/iteration-1-plan.md`, `docs/iteration-2-plan.md` — design
  decisions and rationale.
- `docs/architecture.md` — short architectural overview.

If you need to add a tool, the pattern is:

1. Add a pure helper to `pptx_ops.py` and unit-test it in
   `tests/test_pptx_ops.py` (or a new file).
2. Add a thin tool wrapper in `server.py` inside `create_server`,
   wrapped in `_safe`.
3. Add an end-to-end test in `tests/test_server_e2e.py` that goes
   through `server.call_tool(...)`.
4. Update the README "Tools exposed" section.
5. Bump version, write changelog, commit.
