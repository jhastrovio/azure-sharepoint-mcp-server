"""Azure authentication utilities for SharePoint access."""

import os
from typing import Optional
from azure.identity import DefaultAzureCredential, ClientSecretCredential
from office365.sharepoint.client_context import ClientContext


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
        self._context: Optional[ClientContext] = None
    
    def get_context(self) -> ClientContext:
        """Get authenticated SharePoint client context.
        
        Returns:
            Authenticated SharePoint ClientContext
            
        Raises:
            ValueError: If required authentication parameters are missing
        """
        if self._context is not None:
            return self._context
            
        if self.client_id and self.client_secret and self.tenant_id:
            # Use service principal authentication
            credential = ClientSecretCredential(
                tenant_id=self.tenant_id,
                client_id=self.client_id,
                client_secret=self.client_secret
            )
        else:
            # Use default Azure credential (managed identity, Azure CLI, etc.)
            credential = DefaultAzureCredential()
        
        # Create SharePoint context with Azure credential
        self._context = ClientContext(self.site_url).with_credentials(credential)
        return self._context
    
    def test_connection(self) -> bool:
        """Test SharePoint connection.
        
        Returns:
            True if connection is successful, False otherwise
        """
        try:
            ctx = self.get_context()
            # Test by getting the web properties
            web = ctx.web
            ctx.load(web)
            ctx.execute_query()
            return True
        except Exception:
            return False
