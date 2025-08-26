"""SharePoint client for file operations."""

import io
from typing import List, Dict, Any, Optional, Union
from pathlib import Path
from office365.sharepoint.files.file import File
from office365.sharepoint.folders.folder import Folder
from office365.sharepoint.client_context import ClientContext

from .auth import SharePointAuthenticator


class SharePointClient:
    """Client for SharePoint file operations."""
    
    def __init__(self, authenticator: SharePointAuthenticator):
        """Initialize SharePoint client.
        
        Args:
            authenticator: SharePoint authenticator instance
        """
        self.authenticator = authenticator
        self._context: Optional[ClientContext] = None
    
    @property
    def context(self) -> ClientContext:
        """Get SharePoint client context."""
        if self._context is None:
            self._context = self.authenticator.get_context()
        return self._context
    
    def list_files(self, folder_path: str = "/") -> List[Dict[str, Any]]:
        """List files in a SharePoint folder.
        
        Args:
            folder_path: SharePoint folder path (default: root)
            
        Returns:
            List of file information dictionaries
        """
        try:
            # Get the folder
            if folder_path == "/":
                folder = self.context.web.get_default_document_library().root_folder
            else:
                folder = self.context.web.get_folder_by_server_relative_url(folder_path)
            
            # Load files and folders
            self.context.load(folder)
            self.context.load(folder.files)
            self.context.load(folder.folders)
            self.context.execute_query()
            
            files_info = []
            
            # Add files
            for file in folder.files:
                files_info.append({
                    "name": file.name,
                    "type": "file",
                    "path": file.server_relative_url,
                    "size": file.length,
                    "modified": file.time_last_modified.isoformat() if file.time_last_modified else None,
                    "created": file.time_created.isoformat() if file.time_created else None,
                })
            
            # Add folders
            for subfolder in folder.folders:
                if not subfolder.name.startswith("_"):  # Skip system folders
                    files_info.append({
                        "name": subfolder.name,
                        "type": "folder",
                        "path": subfolder.server_relative_url,
                        "size": None,
                        "modified": None,
                        "created": None,
                    })
            
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
            file = self.context.web.get_file_by_server_relative_url(file_path)
            self.context.load(file)
            self.context.execute_query()
            
            # Download file content
            content = file.read()
            return content
            
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
            # Convert string content to bytes
            if isinstance(content, str):
                content = content.encode("utf-8")
            
            # Get the target folder path and filename
            path_parts = file_path.strip("/").split("/")
            filename = path_parts[-1]
            folder_path = "/" + "/".join(path_parts[:-1]) if len(path_parts) > 1 else "/"
            
            # Get the target folder
            if folder_path == "/":
                target_folder = self.context.web.get_default_document_library().root_folder
            else:
                target_folder = self.context.web.get_folder_by_server_relative_url(folder_path)
            
            # Upload the file
            file_stream = io.BytesIO(content)
            uploaded_file = target_folder.upload_file(filename, file_stream, overwrite)
            self.context.execute_query()
            
            # Load file properties
            self.context.load(uploaded_file)
            self.context.execute_query()
            
            return {
                "name": uploaded_file.name,
                "path": uploaded_file.server_relative_url,
                "size": uploaded_file.length,
                "modified": uploaded_file.time_last_modified.isoformat() if uploaded_file.time_last_modified else None,
                "created": uploaded_file.time_created.isoformat() if uploaded_file.time_created else None,
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
            file = self.context.web.get_file_by_server_relative_url(file_path)
            file.delete_object()
            self.context.execute_query()
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
            # Get parent folder and new folder name
            path_parts = folder_path.strip("/").split("/")
            folder_name = path_parts[-1]
            parent_path = "/" + "/".join(path_parts[:-1]) if len(path_parts) > 1 else "/"
            
            # Get parent folder
            if parent_path == "/":
                parent_folder = self.context.web.get_default_document_library().root_folder
            else:
                parent_folder = self.context.web.get_folder_by_server_relative_url(parent_path)
            
            # Create new folder
            new_folder = parent_folder.folders.add(folder_name)
            self.context.execute_query()
            
            return {
                "name": new_folder.name,
                "type": "folder",
                "path": new_folder.server_relative_url,
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
            file = self.context.web.get_file_by_server_relative_url(file_path)
            self.context.load(file)
            self.context.execute_query()
            return file.exists
        except Exception:
            return False
