"""Server configuration loaded from environment variables.

The server binds to exactly one PPTX file (PPTX_TARGET) and optionally a
template (PPTX_TEMPLATE). Both are resolved into a frozen ServerConfig at
startup so the rest of the code can rely on absolute paths.
"""

from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path


class ConfigError(RuntimeError):
    """Raised when environment-variable configuration is invalid."""


_VALID_TEMPLATE_SUFFIXES = {".pptx", ".potx"}
_VALID_TARGET_SUFFIXES = {".pptx"}


@dataclass(frozen=True)
class ServerConfig:
    """Immutable runtime configuration for the MCP server."""

    target_path: Path
    template_path: Path | None


def _resolve_target(raw: str) -> Path:
    path = Path(raw).expanduser()
    resolved = path.resolve(strict=False)
    if resolved.suffix.lower() not in _VALID_TARGET_SUFFIXES:
        raise ConfigError(
            f"PPTX_TARGET must point to a .pptx file, got: {resolved}"
        )
    parent = resolved.parent
    if not parent.exists():
        raise ConfigError(
            f"PPTX_TARGET parent directory does not exist: {parent}"
        )
    if not parent.is_dir():
        raise ConfigError(
            f"PPTX_TARGET parent is not a directory: {parent}"
        )
    if resolved.exists():
        if not resolved.is_file():
            raise ConfigError(
                f"PPTX_TARGET exists but is not a regular file: {resolved}"
            )
        if not os.access(resolved, os.R_OK):
            raise ConfigError(
                f"PPTX_TARGET exists but is not readable: {resolved}"
            )
    return resolved


def _resolve_template(raw: str) -> Path:
    path = Path(raw).expanduser()
    try:
        resolved = path.resolve(strict=True)
    except FileNotFoundError as exc:
        raise ConfigError(
            f"PPTX_TEMPLATE points to a non-existent file: {path}"
        ) from exc
    if not resolved.is_file():
        raise ConfigError(
            f"PPTX_TEMPLATE is not a regular file: {resolved}"
        )
    if resolved.suffix.lower() not in _VALID_TEMPLATE_SUFFIXES:
        raise ConfigError(
            f"PPTX_TEMPLATE must end in .pptx or .potx, got: {resolved}"
        )
    return resolved


def load_config_from_env() -> ServerConfig:
    """Read PPTX_TARGET and (optional) PPTX_TEMPLATE from the environment."""
    raw_target = os.environ.get("PPTX_TARGET")
    if not raw_target:
        raise ConfigError(
            "PPTX_TARGET is required. Set it to the absolute path of the "
            ".pptx file the server should be bound to."
        )
    target = _resolve_target(raw_target)

    raw_template = os.environ.get("PPTX_TEMPLATE")
    template = _resolve_template(raw_template) if raw_template else None

    return ServerConfig(target_path=target, template_path=template)
