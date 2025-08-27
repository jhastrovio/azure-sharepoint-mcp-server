#!/bin/bash
cd /home/site/wwwroot

# Set Python path
export PYTHONPATH=/home/site/wwwroot/src:$PYTHONPATH

# Install requirements
echo "Installing requirements..."
pip install -r requirements.txt

# Start the application
echo "Starting SharePoint MCP Server..."
python application.py
