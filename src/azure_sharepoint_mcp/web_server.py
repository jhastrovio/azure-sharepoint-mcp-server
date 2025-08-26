"""Web server wrapper for Azure SharePoint MCP Server."""

import json
import logging
import os
from typing import Any, Dict, List

from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

from .server import SharePointMCPServer, SharePointConfig

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Create FastAPI app
app = FastAPI(
    title="Azure SharePoint MCP Server",
    description="MCP Server for SharePoint operations via Microsoft Graph API",
    version="1.0.0"
)

# Add CORS middleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Initialize MCP server
mcp_server = None

def initialize_mcp_server():
    """Initialize the MCP server with configuration."""
    global mcp_server
    try:
        config = SharePointConfig(
            sharepoint_site_url=os.getenv("SHAREPOINT_SITE_URL"),
            azure_tenant_id=os.getenv("AZURE_TENANT_ID"),
            azure_client_id=os.getenv("AZURE_CLIENT_ID"),
            azure_client_secret=os.getenv("AZURE_CLIENT_SECRET")
        )
        mcp_server = SharePointMCPServer(config)
        logger.info("MCP Server initialized successfully")
    except Exception as e:
        logger.error(f"Failed to initialize MCP Server: {e}")
        raise

@app.on_event("startup")
async def startup_event():
    """Initialize MCP server on startup."""
    initialize_mcp_server()

@app.get("/")
async def root():
    """Root endpoint."""
    return {
        "message": "Azure SharePoint MCP Server",
        "status": "running",
        "version": "1.0.0"
    }

@app.get("/health")
async def health():
    """Health check endpoint."""
    return {"status": "healthy", "service": "sharepoint-mcp"}

@app.get("/tools")
async def tools():
    """List available MCP tools."""
    if mcp_server:
        return {"tools": [tool.name for tool in mcp_server.list_tools()]}
    return {"tools": ["list_files", "read_file", "write_file", "delete_file", "create_folder", "get_site_info"]}

@app.post("/execute")
async def execute_tool(tool_name: str, params: Dict[str, Any] = None):
    """Execute an MCP tool."""
    if not mcp_server:
        raise HTTPException(status_code=500, detail="MCP Server not initialized")
    
    try:
        result = await mcp_server.call_tool(tool_name, params or {})
        return {"success": True, "result": result}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/site-info")
async def get_site_info():
    """Get SharePoint site information."""
    if not mcp_server:
        raise HTTPException(status_code=500, detail="MCP Server not initialized")
    
    try:
        result = await mcp_server.call_tool("get_site_info", {})
        return {"success": True, "result": result}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/files")
async def list_files():
    """List files in SharePoint."""
    if not mcp_server:
        raise HTTPException(status_code=500, detail="MCP Server not initialized")
    
    try:
        result = await mcp_server.call_tool("list_files", {})
        return {"success": True, "result": result}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
