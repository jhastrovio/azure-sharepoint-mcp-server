"""Gunicorn configuration for Azure Web Apps."""

import os
import sys

# Add src directory to Python path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

# Import the FastAPI app
from azure_sharepoint_mcp.web_server import app

# Gunicorn configuration
bind = "0.0.0.0:8000"
workers = 1
worker_class = "uvicorn.workers.UvicornWorker"
timeout = 120
keepalive = 2
max_requests = 1000
max_requests_jitter = 100
