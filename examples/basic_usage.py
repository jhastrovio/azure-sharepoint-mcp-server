#!/usr/bin/env python3
"""
Basic usage example for Azure SharePoint MCP Server.

This example demonstrates how to use the SharePoint MCP server
to perform basic file operations.
"""

import asyncio
import os
from azure_sharepoint_mcp import SharePointMCPServer, SharePointConfig


async def main():
    """Demonstrate basic SharePoint operations."""
    
    # Configuration
    config = SharePointConfig(
        site_url=os.getenv("SHAREPOINT_SITE_URL", "https://yourtenant.sharepoint.com/sites/yoursite"),
        tenant_id=os.getenv("AZURE_TENANT_ID"),
        client_id=os.getenv("AZURE_CLIENT_ID"),
        client_secret=os.getenv("AZURE_CLIENT_SECRET"),
    )
    
    print("Azure SharePoint MCP Server Example")
    print("===================================")
    print(f"Site URL: {config.site_url}")
    print()
    
    # Create server instance
    server = SharePointMCPServer(config)
    
    # Test connection
    print("Testing connection...")
    try:
        success = server.authenticator.test_connection()
        if success:
            print("✓ Connection successful!")
        else:
            print("✗ Connection failed!")
            return
    except Exception as e:
        print(f"✗ Connection error: {e}")
        return
    
    print()
    
    # Example operations
    try:
        # List files in root directory
        print("Listing files in root directory:")
        files = server.client.list_files("/")
        for file_info in files[:5]:  # Show first 5 files
            print(f"  {file_info['type']}: {file_info['name']}")
        print()
        
        # Get site information
        print("Getting site information...")
        site_info = server.client.get_site_info()
        print(f"Site Name: {site_info['name']}")
        print(f"Site ID: {site_info['id']}")
        print()
        
        # Create a test file
        print("Creating test file...")
        import time
        test_content = f"Hello from Azure SharePoint MCP Server!\nGenerated at: {time.strftime('%Y-%m-%d %H:%M:%S')}"
        result = server.client.write_file("test_mcp_file.txt", test_content)
        print(f"✓ Created: {result['name']} ({result['size']} bytes)")
        print()
        
        # Read the test file
        print("Reading test file...")
        content = server.client.read_file_text("test_mcp_file.txt")
        print(f"Content:\n{content}")
        print()
        
        # Check if file exists
        print("Checking if file exists...")
        exists = server.client.file_exists("test_mcp_file.txt")
        print(f"File exists: {exists}")
        print()
        
        # Clean up - delete test file
        print("Cleaning up...")
        server.client.delete_file("test_mcp_file.txt")
        print("✓ Test file deleted")
        
    except Exception as e:
        print(f"Error during operations: {e}")


if __name__ == "__main__":
    # Check if required environment variables are set
    if not os.getenv("SHAREPOINT_SITE_URL"):
        print("Error: SHAREPOINT_SITE_URL environment variable is required")
        print("Please set it to your SharePoint site URL, e.g.:")
        print("export SHAREPOINT_SITE_URL='https://yourtenant.sharepoint.com/sites/yoursite'")
        exit(1)
    
    asyncio.run(main())
