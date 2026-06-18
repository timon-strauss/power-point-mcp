"""End-to-end-ish tests that drive the registered tool functions directly.

We bypass the transport layer because exercising stdio inside pytest is
fragile. The tools are still real FastMCP-registered callables; this test
verifies the wiring (open -> mutate -> save -> reopen) works.
"""

from __future__ import annotations

import asyncio
from typing import Any

import pytest
from pptx import Presentation

from power_point_mcp.config import ServerConfig
from power_point_mcp.server import create_server


def _call(server: Any, name: str, **kwargs: Any) -> Any:
    """Invoke a registered FastMCP tool by name and return its python value."""
    result = asyncio.run(server.call_tool(name, kwargs))
    # FastMCP.call_tool returns either (content_list, structured_value) on
    # newer versions or just content_list. Prefer the structured value.
    if isinstance(result, tuple) and len(result) == 2:
        result = result[1]
    # FastMCP wraps non-object tool returns (and sometimes object returns)
    # in a ``{"result": ...}`` envelope for the structured payload. Unwrap
    # it so callers see the tool's actual return value.
    if isinstance(result, dict) and set(result.keys()) == {"result"}:
        return result["result"]
    return result


def test_full_flow(cfg: ServerConfig) -> None:
    server = create_server(cfg)

    created = _call(server, "create_presentation_from_template", overwrite=True)
    assert isinstance(created, dict)
    assert "error" not in created

    info = _call(server, "presentation_info")
    assert isinstance(info, dict)
    assert info.get("slide_count") == 0

    # Discover the actual title placeholder name from the freshly created file.
    prs = Presentation(str(cfg.target_path))
    layout = next(l for l in prs.slide_layouts if l.name == "Title Slide")
    title_ph_name = next(
        ph.name for ph in layout.placeholders if ph.placeholder_format.idx == 0
    )

    added = _call(
        server,
        "add_slide",
        layout_name="Title Slide",
        placeholders={title_ph_name: "Hi"},
    )
    assert isinstance(added, dict)
    assert added.get("slide_index") == 0

    listed = _call(server, "list_slides")
    assert isinstance(listed, list)
    assert len(listed) == 1
    assert listed[0]["layout_name"] == "Title Slide"


def test_add_slide_before_create_returns_error(cfg: ServerConfig) -> None:
    server = create_server(cfg)
    out = _call(server, "add_slide", layout_name="Title Slide", placeholders={})
    assert isinstance(out, dict)
    assert "error" in out


def test_create_from_template_requires_template(tmp_path: Any) -> None:
    cfg = ServerConfig(
        target_path=(tmp_path / "out.pptx").resolve(), template_path=None
    )
    server = create_server(cfg)
    out = _call(server, "create_presentation_from_template", overwrite=True)
    assert isinstance(out, dict)
    assert "error" in out
