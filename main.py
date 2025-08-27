"""Main application file for Azure App Service deployment."""

import os
import sys

# Add the src directory to Python path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

# Import the FastAPI app
from azure_sharepoint_mcp.web_server import app

# This is the standard way Azure App Service expects to find the app
# The app variable should be accessible at module level
