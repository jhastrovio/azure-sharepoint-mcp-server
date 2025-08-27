"""Azure SharePoint MCP Server implementation."""

import asyncio
import json
import logging
from typing import Any, Dict, List, Optional
from mcp.server import Server
from mcp.server.models import InitializationOptions
from mcp.types import (
    Resource,
    Tool,
    TextContent,
    ImageContent,
    EmbeddedResource,
)
from pydantic import BaseModel

from .auth import SharePointAuthenticator
from .graph_client import GraphSharePointClient


# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class SharePointConfig(BaseModel):
    """SharePoint configuration model."""
    site_url: str
    tenant_id: Optional[str] = None
    client_id: Optional[str] = None
    client_secret: Optional[str] = None


class SharePointMCPServer:
    """Azure SharePoint MCP Server."""
    
    def __init__(self, config: SharePointConfig):
        """Initialize SharePoint MCP Server.
        
        Args:
            config: SharePoint configuration
        """
        self.config = config
        self.server = Server("azure-sharepoint-mcp-server")
        self.authenticator = SharePointAuthenticator(
            site_url=config.site_url,
            tenant_id=config.tenant_id,
            client_id=config.client_id,
            client_secret=config.client_secret,
        )
        self.client = GraphSharePointClient(self.authenticator)
        
        # Register handlers
        self._register_handlers()
    
    def _register_handlers(self) -> None:
        """Register MCP handlers."""
        
        @self.server.list_resources()
        async def handle_list_resources() -> List[Resource]:
            """List available SharePoint resources."""
            return [
                Resource(
                    uri="sharepoint://files",
                    name="SharePoint Files",
                    description="Access to SharePoint files and folders",
                    mimeType="application/json",
                ),
            ]
        
        @self.server.read_resource()
        async def handle_read_resource(uri: str) -> str:
            """Read SharePoint resource."""
            if uri == "sharepoint://files":
                try:
                    files = await self.client.list_files("/")
                    return json.dumps(files, indent=2)
                except Exception as e:
                    logger.error(f"Failed to list files: {e}")
                    return json.dumps({"error": str(e)})
            else:
                return json.dumps({"error": f"Unknown resource: {uri}"})
        
        @self.server.list_tools()
        async def handle_list_tools() -> List[Tool]:
            """List available SharePoint tools."""
            return [
                Tool(
                    name="list_files",
                    description="List files and folders in SharePoint",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "folder_path": {
                                "type": "string",
                                "description": "SharePoint folder path (default: /)",
                                "default": "/",
                            }
                        },
                    },
                ),
                Tool(
                    name="read_file",
                    description="Read a file from SharePoint",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "file_path": {
                                "type": "string",
                                "description": "SharePoint file path",
                            },
                            "encoding": {
                                "type": "string",
                                "description": "Text encoding (default: utf-8)",
                                "default": "utf-8",
                            },
                        },
                        "required": ["file_path"],
                    },
                ),
                Tool(
                    name="write_file",
                    description="Write a file to SharePoint",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "file_path": {
                                "type": "string",
                                "description": "SharePoint file path",
                            },
                            "content": {
                                "type": "string",
                                "description": "File content",
                            },
                            "overwrite": {
                                "type": "boolean",
                                "description": "Whether to overwrite existing file",
                                "default": True,
                            },
                        },
                        "required": ["file_path", "content"],
                    },
                ),
                Tool(
                    name="delete_file",
                    description="Delete a file from SharePoint",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "file_path": {
                                "type": "string",
                                "description": "SharePoint file path",
                            },
                        },
                        "required": ["file_path"],
                    },
                ),
                Tool(
                    name="create_folder",
                    description="Create a folder in SharePoint",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "folder_path": {
                                "type": "string",
                                "description": "SharePoint folder path",
                            },
                        },
                        "required": ["folder_path"],
                    },
                ),
                Tool(
                    name="file_exists",
                    description="Check if a file exists in SharePoint",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "file_path": {
                                "type": "string",
                                "description": "SharePoint file path",
                            },
                        },
                        "required": ["file_path"],
                    },
                ),
                Tool(
                    name="test_connection",
                    description="Test SharePoint connection",
                    inputSchema={
                        "type": "object",
                        "properties": {},
                    },
                ),
                Tool(
                    name="get_site_info",
                    description="Get SharePoint site information",
                    inputSchema={
                        "type": "object",
                        "properties": {},
                    },
                ),
            ]
        
        @self.server.call_tool()
        async def handle_call_tool(name: str, arguments: Dict[str, Any]) -> List[TextContent]:
            """Handle tool calls."""
            try:
                if name == "list_files":
                    folder_path = arguments.get("folder_path", "/")
                    files = await self.client.list_files(folder_path)
                    return [TextContent(type="text", text=json.dumps(files, indent=2))]
                
                elif name == "read_file":
                    file_path = arguments["file_path"]
                    encoding = arguments.get("encoding", "utf-8")
                    
                    try:
                        content = await self.client.read_file_text(file_path, encoding)
                        return [TextContent(type="text", text=content)]
                    except UnicodeDecodeError:
                        # If text decoding fails, return as binary
                        content = await self.client.read_file(file_path)
                        return [TextContent(
                            type="text",
                            text=f"Binary file content ({len(content)} bytes): {content[:100]}..."
                        )]
                
                elif name == "write_file":
                    file_path = arguments["file_path"]
                    content = arguments["content"]
                    overwrite = arguments.get("overwrite", True)
                    
                    result = await self.client.write_file(file_path, content, overwrite)
                    return [TextContent(type="text", text=json.dumps(result, indent=2))]
                
                elif name == "delete_file":
                    file_path = arguments["file_path"]
                    success = await self.client.delete_file(file_path)
                    return [TextContent(
                        type="text",
                        text=json.dumps({"success": success, "message": f"File '{file_path}' deleted"})
                    )]
                
                elif name == "create_folder":
                    folder_path = arguments["folder_path"]
                    result = await self.client.create_folder(folder_path)
                    return [TextContent(type="text", text=json.dumps(result, indent=2))]
                
                elif name == "file_exists":
                    file_path = arguments["file_path"]
                    exists = await self.client.file_exists(file_path)
                    return [TextContent(
                        type="text",
                        text=json.dumps({"exists": exists, "file_path": file_path})
                    )]
                
                elif name == "test_connection":
                    success = await asyncio.to_thread(self.authenticator.test_connection)
                    return [TextContent(
                        type="text",
                        text=json.dumps({
                            "connected": success,
                            "site_url": self.config.site_url,
                            "message": "Connection successful" if success else "Connection failed"
                        })
                    )]
                
                elif name == "get_site_info":
                    site_info = await self.client.get_site_info()
                    return [TextContent(type="text", text=json.dumps(site_info, indent=2))]
                
                else:
                    return [TextContent(type="text", text=f"Unknown tool: {name}")]
                    
            except Exception as e:
                logger.error(f"Tool '{name}' failed: {e}")
                return [TextContent(type="text", text=json.dumps({"error": str(e)}))]

        # Expose handlers for direct invocation in tests
        self._handle_list_resources = handle_list_resources
        self._handle_read_resource = handle_read_resource
        self._handle_list_tools = handle_list_tools
        self._handle_call_tool = handle_call_tool
    
    async def run(self, transport_type: str = "stdio") -> None:
        """Run the MCP server.
        
        Args:
            transport_type: Transport type (stdio, websocket, etc.)
        """
        if transport_type == "stdio":
            from mcp.server.stdio import stdio_server
            
            async with stdio_server() as (read_stream, write_stream):
                await self.server.run(
                    read_stream,
                    write_stream,
                    InitializationOptions(
                        server_name="azure-sharepoint-mcp-server",
                        server_version="0.1.0",
                        capabilities=self.server.get_capabilities(
                            notification_options=None,
                            experimental_capabilities=None,
                        ),
                    ),
                )
        else:
            raise ValueError(f"Unsupported transport type: {transport_type}")


async def main() -> None:
    """Main entry point."""
    import os
    
    # Load configuration from environment variables
    config = SharePointConfig(
        site_url=os.getenv("SHAREPOINT_SITE_URL", ""),
        tenant_id=os.getenv("AZURE_TENANT_ID"),
        client_id=os.getenv("AZURE_CLIENT_ID"),
        client_secret=os.getenv("AZURE_CLIENT_SECRET"),
    )
    
    if not config.site_url:
        raise ValueError("SHAREPOINT_SITE_URL environment variable is required")
    
    server = SharePointMCPServer(config)
    await server.run()


if __name__ == "__main__":
    asyncio.run(main())
