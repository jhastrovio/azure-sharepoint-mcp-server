"""Web server wrapper for Azure SharePoint MCP Server."""

import json
import logging
import os
from typing import Any, Dict, List

from flask import Flask, request, jsonify
from flask_cors import CORS

from .server import SharePointMCPServer, SharePointConfig

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Create Flask app
app = Flask(__name__)
logger.info("Flask app created.")

# Add CORS support
CORS(app)

# Initialize MCP server
mcp_server = None

def initialize_mcp_server():
    """Initialize the MCP server with configuration."""
    global mcp_server
    try:
        config = SharePointConfig(
            site_url=os.getenv("SHAREPOINT_SITE_URL"),
            tenant_id=os.getenv("AZURE_TENANT_ID"),
            client_id=os.getenv("AZURE_CLIENT_ID"),
            client_secret=os.getenv("AZURE_CLIENT_SECRET"),
        )
        mcp_server = SharePointMCPServer(config)
        logger.info("MCP server initialized successfully.")
    except Exception as e:
        logger.error(f"Failed to initialize MCP Server: {e}")
        raise

# Initialize MCP server on app startup
@app.before_first_request
def startup_event():
    """Initialize MCP server on startup."""
    try:
        initialize_mcp_server()
    except Exception as e:
        logger.error(f"Startup crash: {type(e).__name__} - {e}", exc_info=True)

@app.route("/", methods=["GET"])
def root():
    """Root endpoint."""
    return jsonify({
        "message": "Azure SharePoint MCP Server",
        "status": "running",
        "version": "1.0.0"
    })

@app.route("/health", methods=["GET"])
def health():
    """Health check endpoint."""
    return jsonify({"status": "healthy", "service": "sharepoint-mcp"})

@app.route("/tools", methods=["GET"])
def tools():
    """List available MCP tools."""
    if mcp_server:
        try:
            tools = mcp_server.list_tools()
            return jsonify({"tools": [tool.name for tool in tools]})
        except Exception as e:
            logger.error(f"Error listing tools: {e}")
            return jsonify({"tools": ["list_files", "read_file", "write_file", "delete_file", "create_folder", "get_site_info"]})
    return jsonify({"tools": ["list_files", "read_file", "write_file", "delete_file", "create_folder", "get_site_info"]})

@app.route("/execute", methods=["POST"])
def execute_tool():
    """Execute an MCP tool."""
    if not mcp_server:
        return jsonify({"error": "MCP Server not initialized"}), 500
    
    try:
        data = request.get_json()
        tool_name = data.get("tool_name")
        params = data.get("params", {})
        
        if not tool_name:
            return jsonify({"error": "tool_name is required"}), 400
        
        result = mcp_server.call_tool(tool_name, params)
        # Handle serialization for different result types
        serialized = []
        for item in result:
            if hasattr(item, 'model_dump'):
                serialized.append(item.model_dump())
            else:
                serialized.append(item)
        return jsonify({"success": True, "result": serialized})
    except Exception as e:
        logger.error(f"Error executing tool {tool_name}: {e}")
        return jsonify({"error": str(e)}), 500

@app.route("/site-info", methods=["GET"])
def get_site_info():
    """Get SharePoint site information."""
    if not mcp_server:
        return jsonify({"error": "MCP Server not initialized"}), 500
    
    try:
        result = mcp_server.call_tool("get_site_info", {})
        serialized = []
        for item in result:
            if hasattr(item, 'model_dump'):
                serialized.append(item.model_dump())
            else:
                serialized.append(item)
        return jsonify({"success": True, "result": serialized})
    except Exception as e:
        logger.error(f"Error getting site info: {e}")
        return jsonify({"error": str(e)}), 500

@app.route("/files", methods=["GET"])
def list_files():
    """List files in SharePoint."""
    if not mcp_server:
        return jsonify({"error": "MCP Server not initialized"}), 500
    
    try:
        result = mcp_server.call_tool("list_files", {})
        serialized = []
        for item in result:
            if hasattr(item, 'model_dump'):
                serialized.append(item.model_dump())
            else:
                serialized.append(item)
        return jsonify({"success": True, "result": serialized})
    except Exception as e:
        logger.error(f"Error listing files: {e}")
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8000, debug=False)
