"""Azure SharePoint MCP Server.

A Model Context Protocol (MCP) server that provides access to Microsoft SharePoint
for reading and writing files through Azure authentication.
"""

__version__ = "0.1.0"
__author__ = "Your Name"
__email__ = "your.email@example.com"

from .server import SharePointMCPServer, SharePointConfig

__all__ = ["SharePointMCPServer", "SharePointConfig"]
