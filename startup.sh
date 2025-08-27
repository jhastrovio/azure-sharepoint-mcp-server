#!/bin/bash
cd /home/site/wwwroot

# Set Python path
export PYTHONPATH=/home/site/wwwroot/src:$PYTHONPATH

# Install requirements
echo "Installing requirements..."
pip install -r requirements.txt

# Start the application using Microsoft's recommended format
echo "Starting SharePoint MCP Server..."
uvicorn main:app --host 0.0.0.0 --port 8000
