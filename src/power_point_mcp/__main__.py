"""Entry point: ``python -m power_point_mcp`` or the ``power-point-mcp`` script."""

from __future__ import annotations

from .config import load_config_from_env
from .server import create_server


def main() -> None:
    cfg = load_config_from_env()
    server = create_server(cfg)
    server.run()  # FastMCP defaults to stdio transport.


if __name__ == "__main__":
    main()
