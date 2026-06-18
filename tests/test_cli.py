"""Tests for the CLI surface in power_point_mcp.__main__."""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

import pytest


REPO_ROOT = Path(__file__).resolve().parents[1]


def _run_cli(*args: str, env: dict[str, str] | None = None) -> subprocess.CompletedProcess[str]:
    """Invoke ``python -m power_point_mcp`` with the given args."""
    cmd = [sys.executable, "-m", "power_point_mcp", *args]
    return subprocess.run(
        cmd,
        cwd=str(REPO_ROOT),
        env=env,
        capture_output=True,
        text=True,
        timeout=30,
    )


def test_version_flag() -> None:
    result = _run_cli("--version", env={"PATH": ""})
    assert result.returncode == 0, result.stderr
    combined = result.stdout + result.stderr
    assert "power-point-mcp" in combined
    assert "0.3.0" in combined


def test_doctor_no_env(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    # An empty cwd ensures no .env is picked up.
    empty = tmp_path / "empty"
    empty.mkdir()
    env = {"PATH": ""}  # deliberately minimal; no PPTX_TARGET
    result = subprocess.run(
        [sys.executable, "-m", "power_point_mcp", "--doctor"],
        cwd=str(empty),
        env=env,
        capture_output=True,
        text=True,
        timeout=30,
    )
    assert result.returncode == 1
    assert "[fail]" in result.stdout
    assert "PPTX_TARGET" in result.stdout


def test_doctor_happy_path(tmp_path: Path, tmp_pptx_template: Path) -> None:
    target = tmp_path / "out.pptx"
    # Touch a real file so all checks pass — copy the fixture.
    target.write_bytes(tmp_pptx_template.read_bytes())

    empty = tmp_path / "cwd"
    empty.mkdir()
    env = {
        "PATH": "",
        "PPTX_TARGET": str(target),
        "PPTX_TEMPLATE": str(tmp_pptx_template),
    }
    result = subprocess.run(
        [sys.executable, "-m", "power_point_mcp", "--doctor"],
        cwd=str(empty),
        env=env,
        capture_output=True,
        text=True,
        timeout=30,
    )
    assert result.returncode == 0, result.stdout + result.stderr
    assert "[ok]" in result.stdout
    assert "all checks passed" in result.stdout
