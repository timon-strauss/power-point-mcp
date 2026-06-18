"""Tests for power_point_mcp.security.assert_within_target."""

from __future__ import annotations

from pathlib import Path

import pytest

from power_point_mcp.config import ServerConfig
from power_point_mcp.security import SecurityError, assert_within_target


def _cfg(target: Path) -> ServerConfig:
    return ServerConfig(target_path=target.resolve(), template_path=None)


def test_accepts_exact_target(tmp_path: Path) -> None:
    target = tmp_path / "out.pptx"
    cfg = _cfg(target)
    assert assert_within_target(target, cfg) == target.resolve()


def test_rejects_sibling(tmp_path: Path) -> None:
    target = tmp_path / "out.pptx"
    cfg = _cfg(target)
    sibling = tmp_path / "other.pptx"
    with pytest.raises(SecurityError):
        assert_within_target(sibling, cfg)


def test_rejects_parent_dir(tmp_path: Path) -> None:
    target = tmp_path / "out.pptx"
    cfg = _cfg(target)
    with pytest.raises(SecurityError):
        assert_within_target(tmp_path, cfg)


def test_rejects_relative_traversal(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    target = tmp_path / "out.pptx"
    cfg = _cfg(target)
    monkeypatch.chdir(tmp_path)
    with pytest.raises(SecurityError):
        assert_within_target(Path("../foo.pptx"), cfg)


def test_rejects_directory_path(tmp_path: Path) -> None:
    a_dir = tmp_path / "subdir"
    a_dir.mkdir()
    cfg = _cfg(a_dir)
    with pytest.raises(SecurityError):
        assert_within_target(a_dir, cfg)
