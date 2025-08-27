"""Tests for SharePoint MCP Server."""

import pytest
from unittest.mock import AsyncMock, Mock, patch
from azure_sharepoint_mcp import SharePointMCPServer, SharePointConfig


@pytest.fixture
def config():
    """Test configuration."""
    return SharePointConfig(
        site_url="https://test.sharepoint.com/sites/test",
        tenant_id="test-tenant-id",
        client_id="test-client-id",
        client_secret="test-client-secret",
    )


@pytest.fixture
def server(config):
    """Test server instance."""
    return SharePointMCPServer(config)


def test_server_initialization(server, config):
    """Test server initialization."""
    assert server.config == config
    assert server.server.name == "azure-sharepoint-mcp-server"
    assert server.authenticator.site_url == config.site_url


@pytest.mark.asyncio
async def test_list_resources(server):
    """Test listing resources."""
    resources = await server._handle_list_resources()
    
    assert len(resources) == 1
    assert str(resources[0].uri) == "sharepoint://files"
    assert resources[0].name == "SharePoint Files"


@pytest.mark.asyncio
async def test_list_tools(server):
    """Test listing tools."""
    tools = await server._handle_list_tools()
    
    tool_names = [tool.name for tool in tools]
    expected_tools = [
        "list_files",
        "read_file", 
        "write_file",
        "delete_file",
        "create_folder",
        "file_exists",
        "test_connection"
    ]
    
    for expected_tool in expected_tools:
        assert expected_tool in tool_names


@pytest.mark.asyncio
async def test_read_resource_files(server):
    """Test reading files resource."""
    with patch.object(server.client, 'list_files', AsyncMock(return_value=[
        {"name": "test.txt", "type": "file", "path": "/test.txt"}
    ])):
        result = await server._handle_read_resource("sharepoint://files")
        assert "test.txt" in result


@pytest.mark.asyncio
async def test_read_resource_unknown(server):
    """Test reading unknown resource."""
    result = await server._handle_read_resource("unknown://resource")
    assert "error" in result.lower()


@pytest.mark.asyncio
async def test_call_tool_list_files(server):
    """Test list_files tool."""
    with patch.object(server.client, 'list_files', AsyncMock(return_value=[
        {"name": "test.txt", "type": "file", "path": "/test.txt"}
    ])):
        result = await server._handle_call_tool("list_files", {"folder_path": "/"})
        assert len(result) == 1
        assert "test.txt" in result[0].text


@pytest.mark.asyncio
async def test_call_tool_test_connection(server):
    """Test test_connection tool."""
    with patch.object(server.authenticator, 'test_connection') as mock_test:
        mock_test.return_value = True
        
        result = await server._handle_call_tool("test_connection", {})
        assert len(result) == 1
        assert "connected" in result[0].text.lower()


@pytest.mark.asyncio
async def test_call_tool_unknown(server):
    """Test unknown tool."""
    result = await server._handle_call_tool("unknown_tool", {})
    assert len(result) == 1
    assert "unknown tool" in result[0].text.lower()
