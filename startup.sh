#!/bin/bash
cd /home/site/wwwroot

# Set Python path
export PYTHONPATH=/home/site/wwwroot/src:$PYTHONPATH

# Debug: Show current directory and Python path
echo "Current directory: $(pwd)"
echo "Python path: $PYTHONPATH"
echo "Listing files in current directory:"
ls -la
echo "Listing src directory:"
ls -la src/
echo "Listing azure_sharepoint_mcp directory:"
ls -la src/azure_sharepoint_mcp/

# Install requirements
echo "Installing requirements..."
pip install -r requirements.txt

# Try to start the application using main.py
echo "Starting SharePoint MCP Server using main.py..."
if python main.py; then
    echo "Application started successfully with main.py"
else
    echo "Failed to start with main.py, trying alternative method..."
    # Fallback: try direct uvicorn command
    uvicorn main:app --host 0.0.0.0 --port 8000
fi
