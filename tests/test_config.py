"""Tests for power_point_mcp.config.load_config_from_env."""

from __future__ import annotations

from pathlib import Path

import pytest
from pptx import Presentation

from power_point_mcp.config import ConfigError, load_config_from_env


def _make_pptx(path: Path) -> Path:
    Presentation().save(str(path))
    return path


def test_missing_target_raises(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.delenv("PPTX_TARGET", raising=False)
    monkeypatch.delenv("PPTX_TEMPLATE", raising=False)
    with pytest.raises(ConfigError):
        load_config_from_env()


def test_template_missing_file_raises(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    target = tmp_path / "out.pptx"
    monkeypatch.setenv("PPTX_TARGET", str(target))
    monkeypatch.setenv("PPTX_TEMPLATE", str(tmp_path / "no-such.pptx"))
    with pytest.raises(ConfigError):
        load_config_from_env()


def test_valid_target_no_template(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    target = (tmp_path / "out.pptx").resolve()
    monkeypatch.setenv("PPTX_TARGET", str(target))
    monkeypatch.delenv("PPTX_TEMPLATE", raising=False)
    cfg = load_config_from_env()
    assert cfg.target_path == target
    assert cfg.template_path is None
    assert cfg.target_path.is_absolute()


def test_valid_both(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    template = _make_pptx(tmp_path / "tpl.pptx").resolve()
    target = (tmp_path / "out.pptx").resolve()
    monkeypatch.setenv("PPTX_TARGET", str(target))
    monkeypatch.setenv("PPTX_TEMPLATE", str(template))
    cfg = load_config_from_env()
    assert cfg.target_path == target
    assert cfg.template_path == template
    assert cfg.template_path.is_absolute()
