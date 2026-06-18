"""Shared pytest fixtures."""

from __future__ import annotations

from pathlib import Path

import pytest
from pptx import Presentation

from power_point_mcp.config import ServerConfig


@pytest.fixture()
def tmp_pptx_template(tmp_path: Path) -> Path:
    """Build a tiny .pptx with python-pptx's default layouts."""
    template_path = tmp_path / "template.pptx"
    prs = Presentation()  # default layouts ("Title Slide", etc.)
    prs.save(str(template_path))
    return template_path


@pytest.fixture()
def cfg(tmp_path: Path, tmp_pptx_template: Path) -> ServerConfig:
    """ServerConfig pointing at a not-yet-existent target plus a real template."""
    return ServerConfig(
        target_path=(tmp_path / "out.pptx").resolve(),
        template_path=tmp_pptx_template.resolve(),
    )
