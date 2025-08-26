"""Azure authentication utilities for SharePoint access."""

import os
import requests
from typing import Optional
from azure.identity import DefaultAzureCredential, ClientSecretCredential


class SharePointAuthenticator:
    """Handles Azure authentication for SharePoint access."""
    
    def __init__(
        self,
        site_url: str,
        tenant_id: Optional[str] = None,
        client_id: Optional[str] = None,
        client_secret: Optional[str] = None,
    ):
        """Initialize SharePoint authenticator.
        
        Args:
            site_url: SharePoint site URL
            tenant_id: Azure tenant ID (optional, can be set via env var)
            client_id: Azure client ID (optional, can be set via env var)
            client_secret: Azure client secret (optional, can be set via env var)
        """
        self.site_url = site_url
        self.tenant_id = tenant_id or os.getenv("AZURE_TENANT_ID")
        self.client_id = client_id or os.getenv("AZURE_CLIENT_ID")
        self.client_secret = client_secret or os.getenv("AZURE_CLIENT_SECRET")
        self._credential: Optional[ClientSecretCredential] = None
    
    def get_credential(self) -> ClientSecretCredential:
        """Get Azure credential.
        
        Returns:
            Azure credential instance
            
        Raises:
            ValueError: If required authentication parameters are missing
        """
        if self._credential is not None:
            return self._credential
            
        if self.client_id and self.client_secret and self.tenant_id:
            # Use service principal authentication
            self._credential = ClientSecretCredential(
                tenant_id=self.tenant_id,
                client_id=self.client_id,
                client_secret=self.client_secret
            )
        else:
            # Use default Azure credential (managed identity, Azure CLI, etc.)
            self._credential = DefaultAzureCredential()
        
        return self._credential
    
    def get_graph_token(self) -> str:
        """Get access token for Microsoft Graph API.
        
        Returns:
            Access token string
        """
        credential = self.get_credential()
        token = credential.get_token("https://graph.microsoft.com/.default")
        return token.token
    
    def test_connection(self) -> bool:
        """Test SharePoint connection via Microsoft Graph.
        
        Returns:
            True if connection is successful, False otherwise
        """
        try:
            token = self.get_graph_token()
            
            # Extract site path from URL
            if "/sites/" in self.site_url:
                site_path = self.site_url.split("/sites/", 1)[1]
                hostname = self.site_url.split("//", 1)[1].split("/", 1)[0]
            else:
                return False
            
            # Test by getting site information
            headers = {
                "Authorization": f"Bearer {token}",
                "Accept": "application/json"
            }
            
            url = f"https://graph.microsoft.com/v1.0/sites/{hostname}:/sites/{site_path}"
            response = requests.get(url, headers=headers)
            
            return response.status_code == 200
        except Exception:
            return False
