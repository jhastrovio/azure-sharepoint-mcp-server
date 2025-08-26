#!/bin/bash
echo "Installing Python dependencies..."
pip install -r requirements.txt
echo "Starting SharePoint MCP Server..."
python -m uvicorn azure_sharepoint_mcp.web_server:app --host 0.0.0.0 --port 8000
