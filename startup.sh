#!/bin/bash
cd /home/site/wwwroot

# Set Python path
export PYTHONPATH=/home/site/wwwroot/src:$PYTHONPATH

# Install dependencies if needed
if [ ! -d "venv" ]; then
    echo "Creating virtual environment..."
    python -m venv venv
fi

# Activate virtual environment
source venv/bin/activate

# Install requirements
echo "Installing requirements..."
pip install -r requirements.txt

# Start the application
echo "Starting SharePoint MCP Server..."
python -m uvicorn azure_sharepoint_mcp.web_server:app --host 0.0.0.0 --port 8000 --log-level info
