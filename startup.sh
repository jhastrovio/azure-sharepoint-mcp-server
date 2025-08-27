#!/bin/bash

# Change to the application directory
cd /home/site/wwwroot

# Activate the virtual environment
source antenv/bin/activate

# Set the Python path
export PYTHONPATH=/home/site/wwwroot/src:$PYTHONPATH

# Start the FastAPI application using uvicorn
exec uvicorn azure_sharepoint_mcp.web_server:app --host 0.0.0.0 --port 8000
