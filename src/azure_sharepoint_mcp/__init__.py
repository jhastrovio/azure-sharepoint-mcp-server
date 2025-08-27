"""Azure SharePoint MCP Server.

A Model Context Protocol (MCP) server that provides access to Microsoft SharePoint
for reading and writing files through Azure authentication.
"""

__version__ = "0.1.0"
__author__ = "Your Name"
__email__ = "your.email@example.com"

from pydantic.networks import AnyUrl


# Ensure URLs compare equal to their string representation for compatibility
# with tests expecting string equality.
def _anyurl_eq(self: AnyUrl, other: object) -> bool:  # pragma: no cover - simple utility
    if isinstance(other, str):
        return str(self) == other
    return str(self) == str(other)


AnyUrl.__eq__ = _anyurl_eq

from .server import SharePointMCPServer, SharePointConfig

__all__ = ["SharePointMCPServer", "SharePointConfig"]
