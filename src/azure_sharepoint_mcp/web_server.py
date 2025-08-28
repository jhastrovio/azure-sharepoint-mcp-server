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

# Initialize MCP server (lazy, safe)
mcp_server = None

def initialize_mcp_server():
    """Attempt to initialize the MCP server.

    Returns:
        (SharePointMCPServer|None, str|None): tuple of (server, error_message).
        Do not raise here so Gunicorn workers don't crash on bad config.
    """
    global mcp_server
    if mcp_server:
        return mcp_server, None

    site_url = os.getenv("SHAREPOINT_SITE_URL")
    tenant_id = os.getenv("AZURE_TENANT_ID")
    client_id = os.getenv("AZURE_CLIENT_ID")
    client_secret = os.getenv("AZURE_CLIENT_SECRET")

    missing = [k for k, v in {
        "SHAREPOINT_SITE_URL": site_url,
        "AZURE_TENANT_ID": tenant_id,
        "AZURE_CLIENT_ID": client_id,
        "AZURE_CLIENT_SECRET": client_secret,
    }.items() if not v]

    if missing:
        msg = f"Missing required environment variables: {', '.join(missing)}"
        logger.error(msg)
        return None, msg

    try:
        config = SharePointConfig(
            site_url=site_url,
            tenant_id=tenant_id,
            client_id=client_id,
            client_secret=client_secret,
        )
        mcp_server = SharePointMCPServer(config)
        logger.info("MCP server initialized successfully.")
        return mcp_server, None
    except Exception as e:
        logger.error(f"Failed to initialize MCP Server: {e}", exc_info=True)
        return None, f"Failed to initialize MCP Server: {e}"


# Try to initialize on startup for a warm worker, but don't crash the process
@app.before_first_request
def startup_event():
    server, err = initialize_mcp_server()
    if err:
        logger.warning(f"MCP server not initialized at startup: {err}")

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
    server, err = initialize_mcp_server()
    if err:
        return jsonify({"error": err}), 500

    try:
        tools = server.list_tools()
        return jsonify({"tools": [tool.name for tool in tools]})
    except Exception as e:
        logger.error(f"Error listing tools: {e}")
        return jsonify({"tools": ["list_files", "read_file", "write_file", "delete_file", "create_folder", "get_site_info"]})

@app.route("/execute", methods=["POST"])
def execute_tool():
    """Execute an MCP tool."""
    server, err = initialize_mcp_server()
    if err:
        return jsonify({"error": err}), 500

    try:
        data = request.get_json()
        tool_name = data.get("tool_name")
        params = data.get("params", {})

        if not tool_name:
            return jsonify({"error": "tool_name is required"}), 400

        result = server.call_tool(tool_name, params)
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
    server, err = initialize_mcp_server()
    if err:
        return jsonify({"error": err}), 500

    try:
        result = server.call_tool("get_site_info", {})
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
    server, err = initialize_mcp_server()
    if err:
        return jsonify({"error": err}), 500

    try:
        result = server.call_tool("list_files", {})
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
