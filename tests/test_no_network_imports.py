"""Static check: no networking imports in the package source."""

from __future__ import annotations

from pathlib import Path

FORBIDDEN = (
    "import requests",
    "import httpx",
    "from urllib.request",
    "import urllib.request",
)


def test_no_network_imports() -> None:
    pkg = Path(__file__).resolve().parent.parent / "src" / "power_point_mcp"
    offenders: list[str] = []
    for py in pkg.rglob("*.py"):
        text = py.read_text(encoding="utf-8")
        for needle in FORBIDDEN:
            if needle in text:
                offenders.append(f"{py}: {needle}")
    assert not offenders, f"Networking imports found: {offenders}"
