"""Microsoft Graph client utility.

This module provides a utility class for making requests to the Microsoft Graph API.
"""

import logging

from msgraph import GraphServiceClient
from msgraph_beta import GraphServiceClient as GraphBetaServiceClient

from ..auth.graph_auth import GraphAuthManager


class GraphClient:
    """Core client utility for Microsoft Graph API interactions.

    This class is responsible for:
    1. Initializing and managing the Microsoft Graph client (v1.0)
    2. Initializing and managing the Microsoft Graph beta client (for Intune/beta features)
    3. Providing core API functionality
    4. Handling shared request configurations
    """

    def __init__(self, auth_manager: GraphAuthManager):
        """Initialize the GraphClient.

        Args:
            auth_manager: GraphAuthManager instance for authentication
        """
        self.auth_manager = auth_manager
        self._client = None
        self._beta_client = None
        self.logger = logging.getLogger(__name__)

    def get_client(self) -> GraphServiceClient:
        """Get or create a Graph client (v1.0 endpoint).

        Returns:
            Initialized GraphServiceClient
        """
        if self._client is None:
            self._client = self.auth_manager.get_graph_client()
            self.logger.info("Graph client initialized")
        return self._client

    def get_beta_client(self) -> GraphBetaServiceClient:
        """Get or create a Graph beta client (beta endpoint).

        The beta client targets ``https://graph.microsoft.com/beta`` and is
        used for Intune device-management features and other APIs that are
        only available on the beta endpoint.

        Returns:
            Initialized beta GraphServiceClient
        """
        if self._beta_client is None:
            self._beta_client = self.auth_manager.get_beta_graph_client()
            self.logger.info("Graph beta client initialized")
        return self._beta_client
