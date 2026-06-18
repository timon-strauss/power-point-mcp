"""Entry point: ``python -m power_point_mcp`` or the ``power-point-mcp`` script.

Supports:
  (no flags)    Start the MCP server (stdio transport).
  --version     Print version and exit.
  --doctor      Validate environment configuration without starting the
                server. Exits 0 on all-clear, 1 on any failure.
"""

from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path

from . import __version__


_VALID_TEMPLATE_SUFFIXES = {".pptx", ".potx"}


def _print_status(level: str, message: str) -> None:
    """Print a status line prefixed with [ok] / [fail] / [warn]."""
    print(f"[{level}] {message}")


def _module_version(module_name: str) -> str:
    try:
        module = __import__(module_name)
    except Exception:
        return "unknown"
    return str(getattr(module, "__version__", "unknown"))


def _check_target() -> bool:
    """Validate PPTX_TARGET env var. Returns True if OK."""
    raw = os.environ.get("PPTX_TARGET")
    if not raw:
        _print_status("fail", "PPTX_TARGET is not set")
        return False
    _print_status("ok", f"PPTX_TARGET is set: {raw}")
    try:
        resolved = Path(raw).expanduser().resolve(strict=False)
    except Exception as exc:
        _print_status("fail", f"PPTX_TARGET cannot be resolved: {exc}")
        return False
    _print_status("ok", f"PPTX_TARGET resolves to: {resolved}")

    ok = True
    if resolved.suffix.lower() != ".pptx":
        _print_status("fail", f"PPTX_TARGET must end in .pptx (got {resolved.suffix!r})")
        ok = False

    parent = resolved.parent
    if parent.exists() and parent.is_dir():
        _print_status("ok", f"PPTX_TARGET parent directory exists: {parent}")
    else:
        _print_status("fail", f"PPTX_TARGET parent directory missing: {parent}")
        ok = False

    if resolved.exists():
        if resolved.is_file():
            _print_status("ok", f"PPTX_TARGET file exists: {resolved}")
        else:
            _print_status("fail", f"PPTX_TARGET exists but is not a file: {resolved}")
            ok = False
        if os.access(resolved, os.R_OK):
            _print_status("ok", "PPTX_TARGET is readable")
        else:
            _print_status("fail", "PPTX_TARGET exists but is not readable")
            ok = False
    else:
        _print_status(
            "warn",
            "PPTX_TARGET does not exist yet "
            "(call create_presentation_from_template to initialise it)",
        )
    return ok


def _check_template() -> bool:
    """Validate PPTX_TEMPLATE env var. Returns True if OK (or unset)."""
    raw = os.environ.get("PPTX_TEMPLATE")
    if not raw:
        _print_status("warn", "PPTX_TEMPLATE is not set (optional)")
        return True
    _print_status("ok", f"PPTX_TEMPLATE is set: {raw}")
    try:
        resolved = Path(raw).expanduser().resolve(strict=False)
    except Exception as exc:
        _print_status("fail", f"PPTX_TEMPLATE cannot be resolved: {exc}")
        return False

    ok = True
    if not resolved.exists():
        _print_status("fail", f"PPTX_TEMPLATE does not exist: {resolved}")
        return False
    if not resolved.is_file():
        _print_status("fail", f"PPTX_TEMPLATE is not a regular file: {resolved}")
        ok = False
    if resolved.suffix.lower() not in _VALID_TEMPLATE_SUFFIXES:
        _print_status(
            "fail",
            f"PPTX_TEMPLATE must end in .pptx or .potx (got {resolved.suffix!r})",
        )
        ok = False
    if os.access(resolved, os.R_OK):
        _print_status("ok", f"PPTX_TEMPLATE is readable: {resolved}")
    else:
        _print_status("fail", "PPTX_TEMPLATE is not readable")
        ok = False
    return ok


def _doctor() -> int:
    """Run environment checks. Returns process exit code."""
    print(f"power-point-mcp {__version__} doctor")
    print("-" * 40)
    ok_target = _check_target()
    ok_template = _check_template()

    pptx_v = _module_version("pptx")
    if pptx_v == "unknown":
        _print_status("warn", "python-pptx version unknown (import failed)")
    else:
        _print_status("ok", f"python-pptx version: {pptx_v}")

    mcp_v = _module_version("mcp")
    if mcp_v == "unknown":
        _print_status("warn", "mcp version unknown (import failed)")
    else:
        _print_status("ok", f"mcp version: {mcp_v}")

    print("-" * 40)
    if ok_target and ok_template:
        _print_status("ok", "all checks passed")
        return 0
    _print_status("fail", "one or more checks failed")
    return 1


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="power-point-mcp",
        description=(
            "MCP server bound to a single PowerPoint file. "
            "Configure via PPTX_TARGET and (optional) PPTX_TEMPLATE."
        ),
    )
    parser.add_argument(
        "--version",
        action="version",
        version=f"power-point-mcp {__version__}",
    )
    parser.add_argument(
        "--doctor",
        action="store_true",
        help="Validate environment configuration without starting the server.",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    """CLI entry point. Returns process exit code."""
    parser = _build_parser()
    args = parser.parse_args(argv)

    if args.doctor:
        return _doctor()

    # Default: start the server. Imports here to keep --version/--help fast
    # and so --doctor can run even if config would reject the env.
    from .config import load_config_from_env
    from .server import create_server

    cfg = load_config_from_env()
    server = create_server(cfg)
    server.run()  # FastMCP defaults to stdio transport.
    return 0


if __name__ == "__main__":
    sys.exit(main())
