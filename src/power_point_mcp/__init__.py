"""power-point-mcp: an MCP server bound to a single PowerPoint file."""

from .server import create_server

__all__ = ["__version__", "create_server"]
__version__ = "0.3.0"
