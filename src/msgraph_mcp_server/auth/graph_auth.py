"""Microsoft Graph authentication module.

This module provides authentication for the Microsoft Graph API using
Azure Identity credentials.

By default the server uses :class:`azure.identity.DefaultAzureCredential`,
which automatically picks up the developer's current Azure login context
(``az login``/``azd auth login``, Visual Studio / VS Code, environment
variables, managed identity, workload identity, etc.). This means no
client id / tenant id / secret configuration is required for normal
interactive development.

For non-interactive / CI scenarios the manager still supports explicit
client-secret or certificate credentials when the relevant environment
variables (or constructor arguments) are provided.
"""

from __future__ import annotations

import logging
import os
from pathlib import Path
from typing import List, Optional

from azure.core.credentials import TokenCredential
from azure.identity import (
    CertificateCredential,
    ClientSecretCredential,
    DefaultAzureCredential,
)
from dotenv import load_dotenv
from msgraph import GraphServiceClient

logger = logging.getLogger(__name__)

# Default scope for Microsoft Graph when using a TokenCredential.
DEFAULT_GRAPH_SCOPES: List[str] = ["https://graph.microsoft.com/.default"]


def _load_env_files() -> None:
    """Load the first ``.env`` file found in the well-known locations.

    This is best-effort: if no file is found the caller's process
    environment is used as-is (which is the expected path when running
    under ``az login`` with no secrets).
    """
    candidates = [
        Path(__file__).resolve().parent.parent.parent.parent / "config" / ".env",
        Path.cwd() / "config" / ".env",
        Path.home() / ".entraid" / ".env",
        Path("/etc/entraid/.env"),
    ]
    for path in candidates:
        try:
            if path.is_file():
                load_dotenv(path, override=False)
                logger.info("Loaded environment variables from %s", path)
                return
        except OSError:
            # Unreadable path – just skip it.
            continue
    logger.debug("No .env file found; relying on the current process environment")


# Load env files once at import time so subclasses/consumers see the values.
_load_env_files()


class AuthenticationError(Exception):
    """Raised when Microsoft Graph authentication cannot be configured."""


class GraphAuthManager:
    """Build an authenticated :class:`GraphServiceClient`.

    Resolution order:

    1. An explicit :class:`~azure.core.credentials.TokenCredential` passed
       via ``credential``.
    2. Client-secret credential if ``tenant_id``/``client_id``/``client_secret``
       are all supplied (or available via ``TENANT_ID``/``CLIENT_ID``/
       ``CLIENT_SECRET`` environment variables).
    3. Certificate credential if ``tenant_id``/``client_id``/
       ``certificate_path`` are all supplied (with an optional password).
    4. :class:`~azure.identity.DefaultAzureCredential` – the preferred
       default which leverages the developer's current Azure login.
    """

    def __init__(
        self,
        *,
        credential: Optional[TokenCredential] = None,
        tenant_id: Optional[str] = None,
        client_id: Optional[str] = None,
        client_secret: Optional[str] = None,
        certificate_path: Optional[str] = None,
        certificate_password: Optional[str] = None,
        scopes: Optional[List[str]] = None,
    ) -> None:
        self._explicit_credential = credential

        self.tenant_id = tenant_id or os.environ.get("TENANT_ID") or os.environ.get("AZURE_TENANT_ID")
        self.client_id = client_id or os.environ.get("CLIENT_ID") or os.environ.get("AZURE_CLIENT_ID")
        self.client_secret = (
            client_secret or os.environ.get("CLIENT_SECRET") or os.environ.get("AZURE_CLIENT_SECRET")
        )
        self.certificate_path = (
            certificate_path
            or os.environ.get("CERTIFICATE_PATH")
            or os.environ.get("AZURE_CLIENT_CERTIFICATE_PATH")
        )
        self.certificate_password = (
            certificate_password
            or os.environ.get("CERTIFICATE_PWD")
            or os.environ.get("AZURE_CLIENT_CERTIFICATE_PASSWORD")
        )

        self.scopes = scopes or list(DEFAULT_GRAPH_SCOPES)
        self._graph_client: Optional[GraphServiceClient] = None
        self._credential: Optional[TokenCredential] = None

    # ------------------------------------------------------------------
    # Credential selection
    # ------------------------------------------------------------------
    def _build_credential(self) -> TokenCredential:
        if self._explicit_credential is not None:
            logger.info("Using caller-supplied TokenCredential")
            return self._explicit_credential

        if self.tenant_id and self.client_id and self.client_secret:
            logger.info("Using ClientSecretCredential (tenant_id/client_id/client_secret)")
            return ClientSecretCredential(
                tenant_id=self.tenant_id,
                client_id=self.client_id,
                client_secret=self.client_secret,
            )

        if self.tenant_id and self.client_id and self.certificate_path:
            logger.info("Using CertificateCredential (tenant_id/client_id/certificate_path)")
            return CertificateCredential(
                tenant_id=self.tenant_id,
                client_id=self.client_id,
                certificate_path=self.certificate_path,
                password=self.certificate_password,
            )

        logger.info(
            "Using DefaultAzureCredential – falling back to the current Azure login context"
        )
        try:
            return DefaultAzureCredential()
        except Exception as exc:  # pragma: no cover - defensive
            raise AuthenticationError(
                f"Failed to initialise DefaultAzureCredential: {exc}"
            ) from exc

    def get_credential(self) -> TokenCredential:
        """Return (and cache) the underlying :class:`TokenCredential`."""
        if self._credential is None:
            self._credential = self._build_credential()
        return self._credential

    # ------------------------------------------------------------------
    # Graph client
    # ------------------------------------------------------------------
    def get_graph_client(self) -> GraphServiceClient:
        """Return (and cache) an authenticated :class:`GraphServiceClient`."""
        if self._graph_client is not None:
            return self._graph_client

        try:
            credential = self.get_credential()
            self._graph_client = GraphServiceClient(
                credentials=credential,
                scopes=self.scopes,
            )
            logger.info("Successfully created Microsoft Graph client")
            return self._graph_client
        except AuthenticationError:
            raise
        except Exception as exc:
            raise AuthenticationError(f"Failed to create Graph client: {exc}") from exc
