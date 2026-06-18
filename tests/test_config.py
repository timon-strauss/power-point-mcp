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


def test_dotenv_supplements_environ(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    """A .env file in cwd fills in vars; os.environ wins for keys it sets."""
    target_from_dotenv = (tmp_path / "fromdotenv.pptx").resolve()
    target_from_env = (tmp_path / "fromenv.pptx").resolve()
    template = _make_pptx(tmp_path / "tpl.pptx").resolve()

    cwd = tmp_path / "cwd"
    cwd.mkdir()
    dotenv = cwd / ".env"
    dotenv.write_text(
        "# a comment\n"
        f'PPTX_TARGET="{target_from_dotenv}"\n'
        f"PPTX_TEMPLATE={template}\n"
    )
    monkeypatch.chdir(cwd)
    monkeypatch.delenv("PPTX_TARGET", raising=False)
    monkeypatch.delenv("PPTX_TEMPLATE", raising=False)

    cfg = load_config_from_env()
    assert cfg.target_path == target_from_dotenv
    assert cfg.template_path == template

    # os.environ wins over .env values when both define the same key.
    monkeypatch.setenv("PPTX_TARGET", str(target_from_env))
    cfg2 = load_config_from_env()
    assert cfg2.target_path == target_from_env
