"""
Microsoft Graph MCP Server.

A FastMCP server implementation that provides tools and resources
for interacting with Microsoft Graph services.
"""

from importlib.metadata import PackageNotFoundError, version

try:
    __version__ = version("entraid-mcp-server")
except PackageNotFoundError:  # running from a source tree that is not installed
    __version__ = "0.0.0"
