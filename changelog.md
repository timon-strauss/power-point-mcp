# Changelog

All notable changes to power-point-mcp will be documented in this file.

The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.3.0] — 2026-06-18

Setup-ergonomics iteration. Easier first-run experience without changing
the security boundary or runtime behaviour.

### Added
- `--version` and `--doctor` CLI flags on `power-point-mcp`. `--doctor`
  validates `PPTX_TARGET`, `PPTX_TEMPLATE`, and dependency versions
  without starting the server, prefixing each line `[ok]` / `[warn]` /
  `[fail]` and exiting non-zero on any failure.
- `.env` file support in `load_config_from_env`. Stdlib parser, no new
  dependency. Precedence: `os.environ` > `.env`.
- `.env.example` at repo root as a starting point.
- `examples/` folder: `build_template.py` (regenerates the template),
  the generated `template.pptx` (tracked), and `examples/README.md`
  with a "from clone to first slide" walkthrough.
- New MCP tool `list_layouts()` — first-class layout discovery
  (returns name, idx, and placeholder names per layout).
- New `tests/test_cli.py` covering `--version`, `--doctor` happy/sad
  paths. Suite is now 45 tests.

### Changed
- `pyproject.toml`: `version = 0.3.0`.
- `src/power_point_mcp/__init__.py`: `__version__ = "0.3.0"`.
- `.gitignore`: ignores `.env` but allows `examples/*.pptx` (intentional
  fixture).
- `README.md` and `SETUP.md`: document `.env` workflow, `--doctor`,
  `--version`, the example template, and the new tool count (11).

## [0.2.1] — 2026-06-18

Documentation iteration. No runtime changes.

### Added
- `SETUP.md` — step-by-step setup, Claude Desktop wiring, and a
  troubleshooting section.
- `CLAUDE.md` — project-internal conventions for future Claude (or
  agent) sessions: hard rules, layering, workflow, scope.
- README cross-link to `SETUP.md`.

### Changed
- `pyproject.toml`: `version = 0.2.1`.
- `src/power_point_mcp/__init__.py`: `__version__ = "0.2.1"`.

## [0.2.0] — 2026-06-18

Mutation iteration. Adds the ability to modify and reorder existing slides
without breaking the single-file boundary or hallucinating text.

### Added
- `set_slide_placeholder(slide_index, placeholder_name, text)` — set a
  placeholder by name. Reports `previous_text` and warns when multiple
  placeholders share a name (first match wins).
- `set_slide_placeholder_by_idx(slide_index, placeholder_idx, text)` —
  same, keyed on `placeholder_format.idx` for templates with non-unique
  placeholder names.
- `set_slide_title(slide_index, text)` — convenience wrapper that
  resolves the title via `slide.shapes.title` with an idx-0 fallback.
  Errors loudly on slides with no title placeholder.
- `delete_slide(slide_index)` — remove a slide via the documented
  `prs.slides._sldIdLst` workaround; isolated to a single helper.
- `reorder_slide(slide_index, new_index)` — move a slide; `new_index`
  is clamped to the valid range; same-index calls are no-ops.

### Changed
- `pyproject.toml`: `version = 0.2.0`; pinned `python-pptx>=1.0.2,<2`
  to bound the private-API hazard from `_sldIdLst`.
- `src/power_point_mcp/__init__.py`: `__version__ = "0.2.0"`.
- `README.md`: tools section lists all ten tools.

### Tests
- New `tests/test_pptx_mutations.py` for delete/reorder.
- `tests/test_pptx_ops.py` extended with placeholder/title tests.
- `tests/test_server_e2e.py` extended with title-set, delete, reorder,
  and bad-index error-path tests. Suite is now 39 tests, ~1 s.

### Known limitations
- `delete_slide` leaves the slide part as an orphan in the .pptx
  package. PowerPoint and python-pptx tolerate this; cleanup deferred.
- No run-level formatting (bold/italic/colour) or shape-level edits
  yet — those land in iteration 3.

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
