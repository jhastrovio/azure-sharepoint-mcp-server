import pytest
from unittest.mock import Mock, patch

from azure_sharepoint_mcp.graph_client import GraphSharePointClient


def _make_client():
    auth = Mock()
    auth.get_graph_token.return_value = "token"
    auth.site_url = "https://test.sharepoint.com/sites/test"
    client = GraphSharePointClient(auth)
    client._get_default_drive_id = Mock(return_value="drive123")
    client.file_exists = Mock(return_value=False)
    return client


def test_write_file_overwrite_replace_and_path():
    client = _make_client()
    captured_headers = {}

    def mock_put(url, headers=None, data=None):
        captured_headers.update(headers)
        response = Mock()
        response.raise_for_status = Mock()
        response.json.return_value = {
            "name": "file.txt",
            "size": 4,
            "id": "1"
        }
        return response

    with patch("requests.put", side_effect=mock_put):
        result = client.write_file("/folder/file.txt", b"data", overwrite=True)

    assert captured_headers["@microsoft.graph.conflictBehavior"] == "replace"
    assert result["path"] == "/folder/file.txt"


def test_write_file_no_overwrite_fail_and_path():
    client = _make_client()
    captured_headers = {}

    def mock_put(url, headers=None, data=None):
        captured_headers.update(headers)
        response = Mock()
        response.raise_for_status = Mock()
        response.json.return_value = {
            "name": "file2.txt",
            "size": 4,
            "id": "2"
        }
        return response

    with patch("requests.put", side_effect=mock_put):
        result = client.write_file("/file2.txt", b"data", overwrite=False)

    assert captured_headers["@microsoft.graph.conflictBehavior"] == "fail"
    assert result["path"] == "/file2.txt"
