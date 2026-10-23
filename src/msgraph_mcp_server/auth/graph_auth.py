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

from azure.core.credentials import TokenCredential
from azure.identity import (
    CertificateCredential,
    ClientSecretCredential,
    DefaultAzureCredential,
)
from dotenv import load_dotenv
from msgraph import GraphServiceClient
from msgraph_beta import GraphServiceClient as GraphBetaServiceClient

logger = logging.getLogger(__name__)

# Default scope for Microsoft Graph when using a TokenCredential.
DEFAULT_GRAPH_SCOPES: list[str] = ["https://graph.microsoft.com/.default"]


def load_env_files() -> None:
    """Load the first ``.env`` file found in the well-known locations.

    This is best-effort: if no file is found the caller's process
    environment is used as-is (which is the expected path when running
    under ``az login`` with no secrets). A variable already set in the
    process environment wins over the file. The server's ``main`` calls
    this at start-up; importing the package loads nothing.
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


class AuthenticationError(Exception):
    """Raised when Microsoft Graph authentication cannot be configured."""


class GraphAuthManager:
    """Build an authenticated :class:`GraphServiceClient`.

    Resolution order:

    1. An explicit :class:`~azure.core.credentials.TokenCredential` passed
       via ``credential``.
    2. Client-secret credential if ``tenant_id``/``client_id``/``client_secret``
       are all supplied (or available via ``TENANT_ID``/``CLIENT_ID``/
       ``CLIENT_SECRET`` environment variables, read when the credential is
       first built rather than at construction).
    3. Certificate credential if ``tenant_id``/``client_id``/
       ``certificate_path`` are all supplied (with an optional password).
    4. :class:`~azure.identity.DefaultAzureCredential` – the preferred
       default which leverages the developer's current Azure login.
    """

    def __init__(
        self,
        *,
        credential: TokenCredential | None = None,
        tenant_id: str | None = None,
        client_id: str | None = None,
        client_secret: str | None = None,
        certificate_path: str | None = None,
        certificate_password: str | None = None,
        scopes: list[str] | None = None,
    ) -> None:
        self._explicit_credential = credential

        self.tenant_id = tenant_id
        self.client_id = client_id
        self.client_secret = client_secret
        self.certificate_path = certificate_path
        self.certificate_password = certificate_password

        self.scopes = scopes or list(DEFAULT_GRAPH_SCOPES)
        self._graph_client: GraphServiceClient | None = None
        self._beta_graph_client: GraphBetaServiceClient | None = None
        self._credential: TokenCredential | None = None

    # ------------------------------------------------------------------
    # Credential selection
    # ------------------------------------------------------------------
    def _build_credential(self) -> TokenCredential:
        if self._explicit_credential is not None:
            logger.info("Using caller-supplied TokenCredential")
            return self._explicit_credential

        # Read the environment only now, after start-up has loaded the settings file.
        env = os.environ.get
        tenant_id = self.tenant_id or env("TENANT_ID") or env("AZURE_TENANT_ID")
        client_id = self.client_id or env("CLIENT_ID") or env("AZURE_CLIENT_ID")
        client_secret = (
            self.client_secret or env("CLIENT_SECRET") or env("AZURE_CLIENT_SECRET")
        )
        certificate_path = (
            self.certificate_path
            or env("CERTIFICATE_PATH")
            or env("AZURE_CLIENT_CERTIFICATE_PATH")
        )
        certificate_password = (
            self.certificate_password
            or env("CERTIFICATE_PWD")
            or env("AZURE_CLIENT_CERTIFICATE_PASSWORD")
        )

        if tenant_id and client_id and client_secret:
            logger.info(
                "Using ClientSecretCredential (tenant_id/client_id/client_secret)"
            )
            return ClientSecretCredential(
                tenant_id=tenant_id,
                client_id=client_id,
                client_secret=client_secret,
            )

        if tenant_id and client_id and certificate_path:
            logger.info(
                "Using CertificateCredential (tenant_id/client_id/certificate_path)"
            )
            return CertificateCredential(
                tenant_id=tenant_id,
                client_id=client_id,
                certificate_path=certificate_path,
                password=certificate_password,
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

    def get_beta_graph_client(self) -> GraphBetaServiceClient:
        """Return (and cache) an authenticated beta :class:`GraphServiceClient`.

        The beta client targets ``https://graph.microsoft.com/beta`` and is
        required for Intune device-management features that are not yet
        exposed on the v1.0 endpoint.
        """
        if self._beta_graph_client is not None:
            return self._beta_graph_client

        try:
            credential = self.get_credential()
            self._beta_graph_client = GraphBetaServiceClient(
                credentials=credential,
                scopes=self.scopes,
            )
            logger.info("Successfully created Microsoft Graph beta client")
            return self._beta_graph_client
        except AuthenticationError:
            raise
        except Exception as exc:
            raise AuthenticationError(
                f"Failed to create Graph beta client: {exc}"
            ) from exc
