# Changelog

All notable changes to power-point-mcp will be documented in this file.

The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.1.0] — 2026-06-18

First iteration. Read-mostly MCP server bound to a single PowerPoint file.

### Added
- Package skeleton under `src/power_point_mcp/` (config, security, pptx_ops, server, `__main__`).
- Environment-variable configuration: `PPTX_TARGET` (required) selects the single
  presentation the server may touch; `PPTX_TEMPLATE` (optional) selects the
  template used for new presentations.
- Single-file security boundary via `assert_within_target` — every file
  operation resolves through this helper; no tool accepts a path argument.
- Five MCP tools:
  - `presentation_info` — slide count, dimensions, layout names.
  - `list_slides` — per-slide layout, title, text snippet.
  - `read_slide` — full per-shape contents of one slide.
  - `create_presentation_from_template` — initialise the target file from
    `PPTX_TEMPLATE`; refuses to overwrite unless `overwrite=True`.
  - `add_slide` — append a slide using a layout name from the template; only
    placeholders the caller supplies are populated. Unknown placeholder keys
    are returned in a `warnings` array; the server never invents text.
- `pytest` suite (19 tests) covering config validation, security boundary,
  pptx operations, end-to-end tool calls, and a network-import guard.
- `docs/iteration-1-plan.md` and `docs/architecture.md`.
- `README.md` with prerequisites, `uv`-based setup, Claude Desktop config
  snippet, tool list, and recommendation for the companion
  [academic-pptx-skill](https://github.com/Gabberflast/academic-pptx-skill).

### Configuration
- `pyproject.toml`: hatchling build backend, `src/` layout,
  `[project.scripts] power-point-mcp`, pytest config, `dev` dependency
  group with `pytest>=8`.
- `.gitignore`: ignores `.venv/`, `__pycache__/`, `.DS_Store`,
  `.pytest_cache/`, `.ruff_cache/`, `.coverage`, `*.pptx`.

### Out of scope (deferred)
- Editing or deleting existing slides.
- Image, chart, table, or media insertion.
- Theme / colour / font manipulation.
- PDF or image export.
- Speaker notes.
- Path arguments on tools (forbidden by design).
- HTTP / SSE transports — stdio only.
