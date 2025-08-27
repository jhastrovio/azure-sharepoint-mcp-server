#!/usr/bin/env python3
"""Test script to verify module imports work correctly."""

import sys
import os

# Add src directory to Python path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

print("Python path:", sys.path)
print("Current directory:", os.getcwd())
print("Files in current directory:", os.listdir('.'))
print("Files in src directory:", os.listdir('src'))

try:
    print("Attempting to import azure_sharepoint_mcp...")
    import azure_sharepoint_mcp
    print("✅ Successfully imported azure_sharepoint_mcp")
    
    print("Attempting to import azure_sharepoint_mcp.web_server...")
    from azure_sharepoint_mcp import web_server
    print("✅ Successfully imported web_server")
    
    print("Attempting to access the app...")
    app = web_server.app
    print("✅ Successfully accessed the FastAPI app")
    
    print("App title:", app.title)
    print("App version:", app.version)
    
except ImportError as e:
    print(f"❌ Import error: {e}")
    print(f"Error type: {type(e)}")
    import traceback
    traceback.print_exc()
except Exception as e:
    print(f"❌ Unexpected error: {e}")
    import traceback
    traceback.print_exc()

print("Test completed.")
