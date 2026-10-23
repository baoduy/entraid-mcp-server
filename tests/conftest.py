"""Shared test harness: an in-memory Microsoft Graph and an MCP Context that record every call.

The fake mirrors the msgraph-sdk request-builder tree the resources use
(``client.groups.by_group_id(id).members.with_url(url).get(...)`` and so on).
Every ``get`` keeps the SDK signature ``get(request_configuration=None)``, so a
call the real SDK would reject is rejected here too. Collections serve pages
whose size follows ``$top`` (or the endpoint's default) and link them with
``odata_next_link``. Entities are real msgraph model objects.
"""

from dataclasses import dataclass
from datetime import datetime, timezone
from types import SimpleNamespace
from typing import Any, List, Optional
from urllib.parse import parse_qs, urlparse

import pytest

from msgraph.generated.models.group import Group
from msgraph.generated.models.user import User

from msgraph_mcp_server import server

GRAPH_ROOT = "https://graph.microsoft.com/v1.0"
DEFAULT_PAGE_SIZE = 100
CREATED = datetime(2026, 1, 15, 9, 30, tzinfo=timezone.utc)


class GraphError(Exception):
    """An error response from the fake Microsoft Graph."""


@dataclass
class Request:
    """One request the fake Graph received."""

    method: str  # "get", "with_url" or "patch"
    path: str
    filter: Optional[str] = None
    search: Optional[List[str]] = None
    top: Optional[int] = None
    url: Optional[str] = None
    body: Any = None


class Collection:
    """A Graph collection such as ``/groups`` or ``/groups/{id}/members``."""

    def __init__(self, graph: "FakeGraph", path: str, max_top: Optional[int] = None):
        self._graph = graph
        self.path = path
        self.max_top = max_top  # Graph rejects a larger $top on this endpoint
        self.items: List[Any] = []
        self.page_size = DEFAULT_PAGE_SIZE
        self.respond_none = False  # Graph answers with no body
        self._entities = {}

    async def get(self, request_configuration=None):
        query = getattr(request_configuration, "query_parameters", None)
        top = getattr(query, "top", None)
        self._graph.requests.append(
            Request(
                "get",
                self.path,
                filter=getattr(query, "filter", None),
                search=getattr(query, "search", None),
                top=top,
            )
        )
        if self.max_top is not None and top is not None and top > self.max_top:
            raise GraphError(f"Invalid page size specified: '{top}'. Must be between 1 and {self.max_top}.")
        if self.respond_none:
            return None
        size = top if top and top > 0 else self.page_size
        return self._page(0, size)

    def with_url(self, url: str) -> "_NextPage":
        return _NextPage(self, url)

    def _page(self, start: int, size: int):
        end = start + size
        next_link = f"{GRAPH_ROOT}/{self.path}?$top={size}&$skiptoken={end}" if end < len(self.items) else None
        return SimpleNamespace(value=self.items[start:end], odata_next_link=next_link)

    def by_id(self, entity_id: str) -> "Entity":
        if entity_id not in self._entities:
            self._entities[entity_id] = Entity(self._graph, f"{self.path}/{entity_id}")
        return self._entities[entity_id]

    by_user_id = by_group_id = by_application_id = by_service_principal_id = by_id


class _NextPage:
    """The request builder ``with_url(odata_next_link)`` returns."""

    def __init__(self, collection: Collection, url: str):
        self._collection = collection
        self._url = url

    async def get(self, request_configuration=None):
        self._collection._graph.requests.append(Request("with_url", self._collection.path, url=self._url))
        query = parse_qs(urlparse(self._url).query)
        return self._collection._page(int(query["$skiptoken"][0]), int(query["$top"][0]))


class Entity:
    """A single Graph object such as ``/users/{id}``, with its navigation collections."""

    def __init__(self, graph: "FakeGraph", path: str):
        self._graph = graph
        self.path = path
        self.value: Any = None
        self.error: Optional[Exception] = None
        self.refuse_patch = False
        self.members = Collection(graph, f"{path}/members", max_top=999)
        self.app_role_assignments = Collection(graph, f"{path}/appRoleAssignments")
        self.oauth2_permission_grants = Collection(graph, f"{path}/oauth2PermissionGrants")
        self.authentication = SimpleNamespace(methods=Collection(graph, f"{path}/authentication/methods"))

    async def get(self, request_configuration=None):
        self._graph.requests.append(Request("get", self.path))
        if self.error is not None:
            raise self.error
        if self.value is None:
            raise GraphError(f"Request_ResourceNotFound: Resource '{self.path}' does not exist.")
        return self.value

    async def patch(self, body):
        self._graph.requests.append(Request("patch", self.path, body=body))
        if self.refuse_patch:
            password = body.password_profile.password
            raise GraphError(f"Request_BadRequest: The password '{password}' does not comply with the password policy.")
        return None


class FakeGraph:
    """Stands in for ``GraphClient``: serves both the v1.0 and the beta client."""

    def __init__(self):
        self.requests: List[Request] = []
        self.get_client_calls = 0
        self.users = Collection(self, "users")
        self.groups = Collection(self, "groups", max_top=999)
        self.applications = Collection(self, "applications")
        self.service_principals = Collection(self, "servicePrincipals")
        self.audit_logs = SimpleNamespace(
            directory_audits=Collection(self, "auditLogs/directoryAudits"),
            sign_ins=Collection(self, "auditLogs/signIns"),
        )
        self.device_management = SimpleNamespace(
            managed_devices=Collection(self, "deviceManagement/managedDevices"),
        )

    def get_client(self) -> "FakeGraph":
        self.get_client_calls += 1
        return self

    def get_beta_client(self) -> "FakeGraph":
        self.get_client_calls += 1
        return self

    # -- what the tests read back ------------------------------------------------
    def requests_to(self, path: str) -> List[Request]:
        return [r for r in self.requests if r.path == path]

    def requests_under(self, path: str) -> List[Request]:
        return [r for r in self.requests if r.path == path or r.path.startswith(path + "/")]

    def password_changes(self, user_id: str) -> List[str]:
        return [r.body.password_profile.password for r in self.requests_to(f"users/{user_id}") if r.method == "patch"]

    # -- tenant data ---------------------------------------------------------------
    def add_user(self, user_id: str, readable: bool = True) -> User:
        """A user the MFA report can look up: profile plus one phone method."""
        user = User(
            id=user_id,
            display_name=f"User {user_id}",
            user_principal_name=user_id,
            mail=user_id,
            company_name="Contoso",
        )
        entity = self.users.by_user_id(user_id)
        entity.value = user
        entity.authentication.methods.items = [
            SimpleNamespace(odata_type="#microsoft.graph.phoneAuthenticationMethod")
        ]
        if not readable:
            entity.error = GraphError(f"Authorization_RequestDenied: cannot read methods of {user_id}")
        return user


def make_groups(count: int, prefix: str = "Group") -> List[Group]:
    return [
        Group(
            id=f"group-{i:04d}",
            display_name=f"{prefix} {i:04d}",
            mail=f"group{i:04d}@contoso.com",
            mail_nickname=f"group{i:04d}",
            description=f"{prefix} number {i}",
            group_types=["Unified"],
            security_enabled=False,
            mail_enabled=True,
            visibility="Private",
            created_date_time=CREATED,
        )
        for i in range(count)
    ]


def make_users(count: int, prefix: str = "user") -> List[User]:
    return [
        User(
            id=f"{prefix}-{i:04d}",
            display_name=f"User {i:04d}",
            mail=f"{prefix}{i:04d}@contoso.com",
            user_principal_name=f"{prefix}{i:04d}@contoso.com",
            given_name="Given",
            surname=f"Surname{i:04d}",
            job_title="Engineer",
            office_location="Dublin",
            business_phones=["+353 1 555 0100"],
            mobile_phone="+353 87 555 0100",
        )
        for i in range(count)
    ]


class FakeContext:
    """Records what a tool sends to the AI client."""

    def __init__(self):
        self.infos: List[str] = []
        self.warnings: List[str] = []
        self.errors: List[str] = []
        self.progress: List[tuple] = []

    async def info(self, message, *args, **kwargs):
        self.infos.append(message)

    async def debug(self, message, *args, **kwargs):
        self.infos.append(message)

    async def warning(self, message, *args, **kwargs):
        self.warnings.append(message)

    async def error(self, message, *args, **kwargs):
        self.errors.append(message)

    async def report_progress(self, progress, total=None, message=None):
        self.progress.append((progress, total))


@pytest.fixture
def graph(monkeypatch) -> FakeGraph:
    """An empty contoso tenant wired into the server in place of the real Graph client."""
    fake = FakeGraph()
    monkeypatch.setattr(server, "graph_client", fake)
    return fake


@pytest.fixture
def ctx() -> FakeContext:
    return FakeContext()


class _FrozenDatetime(datetime):
    @classmethod
    def now(cls, tz=None):
        return datetime(2026, 10, 23, 8, 0, 0, tzinfo=timezone.utc)


@pytest.fixture
def frozen_now(monkeypatch):
    """Pins "now" to 2026-10-23T08:00:00Z for the modules that build date-range filters."""
    from msgraph_mcp_server.resources import audit_logs, signin_logs

    monkeypatch.setattr(audit_logs, "datetime", _FrozenDatetime)
    monkeypatch.setattr(signin_logs, "datetime", _FrozenDatetime)
