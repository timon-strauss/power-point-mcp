# Iteration 2 — Plan

## Scope
Iteration 1 covered read + create + append. Iteration 2 adds **modify and structural changes** to existing slides while keeping the single-file-bound, no-hallucination, no-path-arguments contract intact.

Five new MCP tools:
1. `set_slide_placeholder(slide_index, placeholder_name, text)` — by name.
2. `set_slide_placeholder_by_idx(slide_index, placeholder_idx, text)` — by idx.
3. `set_slide_title(slide_index, text)` — convenience wrapper for the title.
4. `delete_slide(slide_index)` — uses `prs.slides._sldIdLst.remove(...)`.
5. `reorder_slide(slide_index, new_index)` — `_sldIdLst` element reordering.

## Function signatures
See actual implementations in `pptx_ops.py` and `server.py`.

## Risks
1. `_sldIdLst` is a private python-pptx API. Pin `python-pptx>=1.0.2,<2`. Isolate to one helper.
2. `delete_slide` leaves the slide part as an orphan in the .pptx package. PowerPoint and python-pptx tolerate this. Cleanup deferred.
3. Concurrency: same as iter 1.

## Testing plan
- Extend `tests/test_pptx_ops.py` with placeholder-set tests.
- New `tests/test_pptx_mutations.py` for delete/reorder.
- Extend `tests/test_server_e2e.py` for end-to-end wiring.

## Out of scope
- Per-shape insert/remove, run-level formatting (bold/italic), speaker notes, theme edits, orphan cleanup, bulk ops, anything outside the bound target.

## Acceptance criteria
- 10 MCP tools registered total.
- Every new tool returns `{"error": ...}` on bad input (no exceptions to MCP client).
- All tests pass via `uv run pytest -q`.
- Version bumped to 0.2.0 in `pyproject.toml` and `src/power_point_mcp/__init__.py`.
- README "Tools" section updated.
