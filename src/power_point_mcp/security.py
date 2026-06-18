"""Security boundary: every filesystem path goes through here.

The server is only allowed to touch the single file recorded in
ServerConfig.target_path. Any tool that accepts a user-supplied path must
funnel it through ``assert_within_target`` first.
"""

from __future__ import annotations

import os
from pathlib import Path

from .config import ServerConfig


class SecurityError(PermissionError):
    """Raised when an operation tries to escape the bound target file."""


def assert_within_target(requested: Path, cfg: ServerConfig) -> Path:
    """Resolve ``requested`` and verify it equals the bound target path."""
    candidate = Path(requested).expanduser().resolve(strict=False)
    if candidate != cfg.target_path:
        raise SecurityError(
            "Refusing to operate on a path other than the bound target. "
            f"requested={candidate} target={cfg.target_path}"
        )
    if candidate.exists() and candidate.is_dir():
        raise SecurityError(
            f"Bound target resolved to a directory, not a file: {candidate}"
        )
    return candidate


def ensure_target_exists_or_can_create(cfg: ServerConfig) -> bool:
    """Return True if the target file exists, or its parent dir is writable."""
    if cfg.target_path.exists():
        return True
    parent = cfg.target_path.parent
    return parent.is_dir() and os.access(parent, os.W_OK)
