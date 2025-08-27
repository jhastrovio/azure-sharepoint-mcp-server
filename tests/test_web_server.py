import pytest
from fastapi.testclient import TestClient
from mcp.types import TextContent
from unittest.mock import AsyncMock, patch
from types import SimpleNamespace

from azure_sharepoint_mcp import web_server


def mock_init(mock_server):
    def _init():
        web_server.mcp_server = mock_server
    return _init


def test_execute_returns_json():
    mock_server = SimpleNamespace()
    mock_server.call_tool = AsyncMock(return_value=[TextContent(type="text", text="hello")])

    with patch("azure_sharepoint_mcp.web_server.initialize_mcp_server", mock_init(mock_server)):
        with TestClient(web_server.app) as client:
            response = client.post("/execute", params={"tool_name": "test"}, json={"params": {}})
            assert response.status_code == 200
            data = response.json()
            assert data["result"][0]["type"] == "text"
            assert data["result"][0]["text"] == "hello"
