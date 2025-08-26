"""Microsoft Graph SharePoint client for file operations."""

import io
import json
import requests
from typing import List, Dict, Any, Optional, Union
from urllib.parse import quote

from .auth import SharePointAuthenticator


class GraphSharePointClient:
    """Microsoft Graph client for SharePoint file operations."""
    
    def __init__(self, authenticator: SharePointAuthenticator):
        """Initialize Microsoft Graph SharePoint client.
        
        Args:
            authenticator: SharePoint authenticator instance
        """
        self.authenticator = authenticator
        self.base_url = "https://graph.microsoft.com/v1.0"
        self._site_id: Optional[str] = None
        self._default_drive_id: Optional[str] = None
    
    def _get_headers(self) -> Dict[str, str]:
        """Get authorization headers for Microsoft Graph API."""
        token = self.authenticator.get_graph_token()
        return {
            "Authorization": f"Bearer {token}",
            "Accept": "application/json",
            "Content-Type": "application/json"
        }
    
    def _get_site_id(self) -> str:
        """Get the SharePoint site ID."""
        if self._site_id is not None:
            return self._site_id
            
        # Extract site path from URL
        site_url = self.authenticator.site_url
        if "/sites/" in site_url:
            site_path = site_url.split("/sites/", 1)[1]
            hostname = site_url.split("//", 1)[1].split("/", 1)[0]
        else:
            raise ValueError("Invalid SharePoint site URL format")
        
        # Get site information
        url = f"{self.base_url}/sites/{hostname}:/sites/{site_path}"
        response = requests.get(url, headers=self._get_headers())
        response.raise_for_status()
        
        site_data = response.json()
        self._site_id = site_data["id"]
        return self._site_id
    
    def _get_default_drive_id(self) -> str:
        """Get the default document library drive ID."""
        if self._default_drive_id is not None:
            return self._default_drive_id
            
        site_id = self._get_site_id()
        url = f"{self.base_url}/sites/{site_id}/drives"
        response = requests.get(url, headers=self._get_headers())
        response.raise_for_status()
        
        drives_data = response.json()
        if not drives_data["value"]:
            raise Exception("No document libraries found")
            
        # Use the first drive (usually "Documents")
        self._default_drive_id = drives_data["value"][0]["id"]
        return self._default_drive_id
    
    def list_files(self, folder_path: str = "/") -> List[Dict[str, Any]]:
        """List files in a SharePoint folder.
        
        Args:
            folder_path: SharePoint folder path (default: root)
            
        Returns:
            List of file information dictionaries
        """
        try:
            drive_id = self._get_default_drive_id()
            
            # Build the path
            if folder_path == "/" or folder_path == "":
                url = f"{self.base_url}/drives/{drive_id}/root/children"
            else:
                # Remove leading/trailing slashes and encode path
                clean_path = folder_path.strip("/")
                encoded_path = quote(clean_path, safe="/")
                url = f"{self.base_url}/drives/{drive_id}/root:/{encoded_path}:/children"
            
            response = requests.get(url, headers=self._get_headers())
            response.raise_for_status()
            
            files_data = response.json()
            files_info = []
            
            for item in files_data.get("value", []):
                file_info = {
                    "name": item["name"],
                    "type": "folder" if "folder" in item else "file",
                    "path": f"/{item['name']}" if folder_path == "/" else f"{folder_path.rstrip('/')}/{item['name']}",
                    "size": item.get("size"),
                    "modified": item.get("lastModifiedDateTime"),
                    "created": item.get("createdDateTime"),
                    "id": item["id"]
                }
                
                if "file" in item:
                    file_info["mimeType"] = item["file"].get("mimeType")
                
                files_info.append(file_info)
            
            return files_info
            
        except Exception as e:
            raise Exception(f"Failed to list files: {str(e)}")
    
    def read_file(self, file_path: str) -> bytes:
        """Read a file from SharePoint.
        
        Args:
            file_path: SharePoint file path
            
        Returns:
            File content as bytes
        """
        try:
            drive_id = self._get_default_drive_id()
            
            # Remove leading slash and encode path
            clean_path = file_path.lstrip("/")
            encoded_path = quote(clean_path, safe="/")
            
            # Get download URL
            url = f"{self.base_url}/drives/{drive_id}/root:/{encoded_path}:/content"
            response = requests.get(url, headers=self._get_headers())
            response.raise_for_status()
            
            return response.content
            
        except Exception as e:
            raise Exception(f"Failed to read file '{file_path}': {str(e)}")
    
    def read_file_text(self, file_path: str, encoding: str = "utf-8") -> str:
        """Read a text file from SharePoint.
        
        Args:
            file_path: SharePoint file path
            encoding: Text encoding (default: utf-8)
            
        Returns:
            File content as string
        """
        content = self.read_file(file_path)
        return content.decode(encoding)
    
    def write_file(
        self, 
        file_path: str, 
        content: Union[str, bytes], 
        overwrite: bool = True
    ) -> Dict[str, Any]:
        """Write a file to SharePoint.
        
        Args:
            file_path: SharePoint file path
            content: File content (string or bytes)
            overwrite: Whether to overwrite existing file
            
        Returns:
            File information dictionary
        """
        try:
            drive_id = self._get_default_drive_id()
            
            # Convert string content to bytes
            if isinstance(content, str):
                content_bytes = content.encode("utf-8")
            else:
                content_bytes = content
            
            # Remove leading slash and encode path
            clean_path = file_path.lstrip("/")
            encoded_path = quote(clean_path, safe="/")
            
            # Upload file
            url = f"{self.base_url}/drives/{drive_id}/root:/{encoded_path}:/content"
            
            headers = self._get_headers()
            headers["Content-Type"] = "application/octet-stream"
            
            response = requests.put(url, headers=headers, data=content_bytes)
            response.raise_for_status()
            
            file_data = response.json()
            
            return {
                "name": file_data["name"],
                "path": f"/{file_data['name']}",
                "size": file_data["size"],
                "modified": file_data.get("lastModifiedDateTime"),
                "created": file_data.get("createdDateTime"),
                "id": file_data["id"]
            }
            
        except Exception as e:
            raise Exception(f"Failed to write file '{file_path}': {str(e)}")
    
    def delete_file(self, file_path: str) -> bool:
        """Delete a file from SharePoint.
        
        Args:
            file_path: SharePoint file path
            
        Returns:
            True if successful
        """
        try:
            drive_id = self._get_default_drive_id()
            
            # Remove leading slash and encode path
            clean_path = file_path.lstrip("/")
            encoded_path = quote(clean_path, safe="/")
            
            url = f"{self.base_url}/drives/{drive_id}/root:/{encoded_path}"
            response = requests.delete(url, headers=self._get_headers())
            response.raise_for_status()
            
            return True
            
        except Exception as e:
            raise Exception(f"Failed to delete file '{file_path}': {str(e)}")
    
    def create_folder(self, folder_path: str) -> Dict[str, Any]:
        """Create a folder in SharePoint.
        
        Args:
            folder_path: SharePoint folder path
            
        Returns:
            Folder information dictionary
        """
        try:
            drive_id = self._get_default_drive_id()
            
            # Parse parent path and folder name
            clean_path = folder_path.strip("/")
            path_parts = clean_path.split("/")
            folder_name = path_parts[-1]
            
            if len(path_parts) > 1:
                parent_path = "/".join(path_parts[:-1])
                encoded_parent = quote(parent_path, safe="/")
                url = f"{self.base_url}/drives/{drive_id}/root:/{encoded_parent}:/children"
            else:
                url = f"{self.base_url}/drives/{drive_id}/root/children"
            
            # Create folder
            data = {
                "name": folder_name,
                "folder": {},
                "@microsoft.graph.conflictBehavior": "fail"
            }
            
            response = requests.post(url, headers=self._get_headers(), json=data)
            response.raise_for_status()
            
            folder_data = response.json()
            
            return {
                "name": folder_data["name"],
                "type": "folder",
                "path": f"/{clean_path}",
                "id": folder_data["id"]
            }
            
        except Exception as e:
            raise Exception(f"Failed to create folder '{folder_path}': {str(e)}")
    
    def file_exists(self, file_path: str) -> bool:
        """Check if a file exists in SharePoint.
        
        Args:
            file_path: SharePoint file path
            
        Returns:
            True if file exists, False otherwise
        """
        try:
            drive_id = self._get_default_drive_id()
            
            # Remove leading slash and encode path
            clean_path = file_path.lstrip("/")
            encoded_path = quote(clean_path, safe="/")
            
            url = f"{self.base_url}/drives/{drive_id}/root:/{encoded_path}"
            response = requests.get(url, headers=self._get_headers())
            
            return response.status_code == 200
            
        except Exception:
            return False
    
    def get_site_info(self) -> Dict[str, Any]:
        """Get SharePoint site information.
        
        Returns:
            Site information dictionary
        """
        try:
            site_id = self._get_site_id()
            url = f"{self.base_url}/sites/{site_id}"
            response = requests.get(url, headers=self._get_headers())
            response.raise_for_status()
            
            site_data = response.json()
            return {
                "id": site_data["id"],
                "name": site_data["name"],
                "webUrl": site_data["webUrl"],
                "description": site_data.get("description", "")
            }
            
        except Exception as e:
            raise Exception(f"Failed to get site info: {str(e)}")
