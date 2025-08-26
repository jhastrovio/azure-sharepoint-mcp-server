# Azure SharePoint MCP Server

A Model Context Protocol (MCP) server that provides secure access to Microsoft SharePoint for reading and writing files through Azure authentication.

## Features

- **File Operations**: Read, write, and delete files in SharePoint
- **Folder Management**: List contents and create folders
- **Azure Authentication**: Secure authentication using Azure Identity
- **MCP Compliance**: Full integration with MCP-compatible clients
- **Type Safety**: Built with Python type hints and Pydantic models

## Installation

### From Source

```bash
git clone https://github.com/yourusername/azure-sharepoint-mcp-server.git
cd azure-sharepoint-mcp-server
pip install -e .
```

### From PyPI (when published)

```bash
pip install azure-sharepoint-mcp-server
```

## Configuration

### Environment Variables

Set the following environment variables:

```bash
# Required
export SHAREPOINT_SITE_URL="https://yourtenant.sharepoint.com/sites/yoursite"

# Optional (if not using default Azure credential)
export AZURE_TENANT_ID="your-tenant-id"
export AZURE_CLIENT_ID="your-client-id"
export AZURE_CLIENT_SECRET="your-client-secret"
```

### Azure App Registration

1. Create an Azure App Registration in the Azure Portal
2. Add SharePoint permissions:
   - `Sites.Read.All` - Read access to SharePoint sites
   - `Sites.ReadWrite.All` - Read/write access to SharePoint sites
3. Grant admin consent for the permissions
4. Create a client secret (if using service principal authentication)

## Usage

### Running the Server

```bash
# Using environment variables
azure-sharepoint-mcp-server

# Or run directly with Python
python -m azure_sharepoint_mcp.server
```

### Available Tools

The server provides the following MCP tools:

#### `list_files`
List files and folders in a SharePoint directory.

```json
{
  "name": "list_files",
  "arguments": {
    "folder_path": "/Documents"
  }
}
```

#### `read_file`
Read the contents of a file from SharePoint.

```json
{
  "name": "read_file",
  "arguments": {
    "file_path": "/Documents/example.txt",
    "encoding": "utf-8"
  }
}
```

#### `write_file`
Write content to a file in SharePoint.

```json
{
  "name": "write_file",
  "arguments": {
    "file_path": "/Documents/new-file.txt",
    "content": "Hello, SharePoint!",
    "overwrite": true
  }
}
```

#### `delete_file`
Delete a file from SharePoint.

```json
{
  "name": "delete_file",
  "arguments": {
    "file_path": "/Documents/file-to-delete.txt"
  }
}
```

#### `create_folder`
Create a new folder in SharePoint.

```json
{
  "name": "create_folder",
  "arguments": {
    "folder_path": "/Documents/new-folder"
  }
}
```

#### `file_exists`
Check if a file exists in SharePoint.

```json
{
  "name": "file_exists",
  "arguments": {
    "file_path": "/Documents/example.txt"
  }
}
```

#### `test_connection`
Test the SharePoint connection.

```json
{
  "name": "test_connection",
  "arguments": {}
}
```

## Client Integration

### Claude Desktop

Add to your Claude Desktop configuration:

```json
{
  "mcpServers": {
    "azure-sharepoint": {
      "command": "azure-sharepoint-mcp-server",
      "env": {
        "SHAREPOINT_SITE_URL": "https://yourtenant.sharepoint.com/sites/yoursite",
        "AZURE_TENANT_ID": "your-tenant-id",
        "AZURE_CLIENT_ID": "your-client-id",
        "AZURE_CLIENT_SECRET": "your-client-secret"
      }
    }
  }
}
```

### Programmatic Usage

```python
import asyncio
from azure_sharepoint_mcp import SharePointMCPServer, SharePointConfig

async def main():
    config = SharePointConfig(
        site_url="https://yourtenant.sharepoint.com/sites/yoursite",
        tenant_id="your-tenant-id",
        client_id="your-client-id",
        client_secret="your-client-secret"
    )
    
    server = SharePointMCPServer(config)
    await server.run()

if __name__ == "__main__":
    asyncio.run(main())
```

## Development

### Setup Development Environment

```bash
# Clone the repository
git clone https://github.com/yourusername/azure-sharepoint-mcp-server.git
cd azure-sharepoint-mcp-server

# Install with development dependencies
pip install -e ".[dev]"

# Install pre-commit hooks
pre-commit install
```

### Running Tests

```bash
pytest
```

### Code Formatting

```bash
# Format code
black src/ tests/
isort src/ tests/

# Type checking
mypy src/
```

## Authentication Methods

### Default Azure Credential (Recommended)

The server uses Azure's DefaultAzureCredential, which tries multiple authentication methods in order:

1. Environment variables (`AZURE_CLIENT_ID`, `AZURE_CLIENT_SECRET`, `AZURE_TENANT_ID`)
2. Managed Identity
3. Azure CLI (`az login`)
4. Azure PowerShell
5. Interactive browser authentication

### Service Principal

For production environments, use a service principal:

```bash
export AZURE_TENANT_ID="your-tenant-id"
export AZURE_CLIENT_ID="your-client-id"
export AZURE_CLIENT_SECRET="your-client-secret"
```

## Security Considerations

- Store secrets securely using Azure Key Vault or environment variables
- Use managed identity when running in Azure
- Follow the principle of least privilege for SharePoint permissions
- Regularly rotate client secrets

## Troubleshooting

### Common Issues

1. **Authentication Failed**: Ensure your Azure credentials are correctly configured
2. **Permission Denied**: Verify your app registration has the required SharePoint permissions
3. **Site Not Found**: Check that the SharePoint site URL is correct and accessible

### Debug Logging

Enable debug logging by setting the log level:

```python
import logging
logging.basicConfig(level=logging.DEBUG)
```

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests for new functionality
5. Ensure all tests pass
6. Submit a pull request

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Related Projects

- [Model Context Protocol](https://github.com/anthropics/mcp)
- [Azure Identity](https://github.com/Azure/azure-sdk-for-python/tree/main/sdk/identity)
- [Office 365 REST Python Client](https://github.com/vgrem/Office365-REST-Python-Client)
