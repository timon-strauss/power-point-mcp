# Iteration 1 — Plan summary

## Scope

Stand up the minimum viable MCP server bound to one PowerPoint file. The
server reads metadata, lists slides, reads slide details, creates the
target file from a template, and adds a single slide using a named layout
with caller-supplied placeholder content. No deletion, no shape editing,
no images, no styling — those land in later iterations.

## Bind mechanism

Configuration comes from environment variables, frozen into a
`ServerConfig` dataclass at startup:

- `PPTX_TARGET` (required): absolute path to the one .pptx the server may
  read or write.
- `PPTX_TEMPLATE` (optional): absolute path to the .pptx/.potx template
  used by `create_presentation_from_template`.

Every filesystem operation routes through
`power_point_mcp.security.assert_within_target`, which resolves the
requested path and rejects anything that does not equal
`cfg.target_path`. This is the entire enforcement boundary — there is no
other place in the code that opens or writes a PPTX.

## Tools

- `presentation_info()` — slide count, dimensions (EMU + inches), layout
  names, master name, whether slide 0 has a title placeholder.
- `list_slides()` — index, layout, best-effort title, 120-char snippet.
- `read_slide(slide_index)` — shape-level detail: id, name, type,
  placeholder index/type, full text.
- `create_presentation_from_template(overwrite=False)` — copy
  `PPTX_TEMPLATE` to `PPTX_TARGET`. Refuses to overwrite unless asked.
- `add_slide(layout_name, placeholders)` — append a slide using a named
  layout; only fills placeholders whose `.name` is a key in the dict.
  Unknown keys are returned as warnings, not errors. Caller-supplied
  text is the only text written — the server invents nothing.

All tool bodies are wrapped in a small `_safe` decorator that catches
configuration, security, and lookup errors and returns
`{"error": "..."}` so MCP clients see structured failures instead of
crashes.

## Security

- Target binding is enforced at every entry point.
- No network code in the package. A `tests/test_no_network_imports.py`
  test fails the build if anyone smuggles in `requests`, `httpx`, or
  `urllib.request`.
- The server never invents user-facing content. If the caller does not
  supply a placeholder value, the placeholder is left untouched.

## Testing

- Unit tests for `config`, `security`, and `pptx_ops`.
- An end-to-end test that drives the registered FastMCP tools directly
  via `server.call_tool(...)`.
- A static test that scans the source tree for forbidden imports.

The tests live under `tests/` and run against the project's local
`.venv`. The Test agent runs them; the Write agent does not.

## Out of scope (deferred)

- Editing or deleting existing slides or shapes.
- Inserting images, charts, or tables.
- Styling, theming, or master edits.
- Authentication, multi-file workspaces, or remote PPTX sources.
- Anything that reaches the network.

## Acceptance criteria

1. `uv sync --group dev` succeeds in the project's .venv.
2. The five MCP tools listed above are registered on the FastMCP
   instance returned by `create_server(cfg)`.
3. Every tool refuses to operate on any file other than
   `cfg.target_path`.
4. The full pytest suite passes locally.
5. The README documents prerequisites, setup, env-var configuration, an
   MCP client snippet, and the academic-pptx-skill recommendation.
