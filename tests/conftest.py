"""Shared test harness: an in-memory Microsoft Graph and an MCP Context that record every call.

The fake mirrors the msgraph-sdk request-builder tree the resources use
(``client.groups.by_group_id(id).members.with_url(url).get(...)`` and so on).
Every ``get`` keeps the SDK signature ``get(request_configuration=None)``, so a
call the real SDK would reject is rejected here too. Collections serve pages
whose size follows ``$top`` (or the endpoint's default) and link them with
``odata_next_link``. Entities are real msgraph model objects.
"""

import re
from dataclasses import dataclass
from datetime import datetime, timezone
from types import SimpleNamespace
from typing import Any
from urllib.parse import parse_qs, urlparse

import pytest
from msgraph.generated.models.directory_object import DirectoryObject
from msgraph.generated.models.group import Group
from msgraph.generated.models.reference_create import ReferenceCreate
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

    method: str  # "get", "with_url", "patch", "post" or "delete"
    path: str
    filter: str | None = None
    search: list[str] | None = None
    top: int | None = None
    url: str | None = None
    body: Any = None


class Collection:
    """A Graph collection such as ``/groups`` or ``/groups/{id}/members``."""

    def __init__(self, graph: "FakeGraph", path: str, max_top: int | None = None):
        self._graph = graph
        self.path = path
        self.max_top = max_top  # Graph rejects a larger $top on this endpoint
        self.items: list[Any] = []
        self.page_size = DEFAULT_PAGE_SIZE
        self.respond_none = False  # Graph answers with no body
        self.error: Exception | None = (
            None  # Graph refuses every read of this collection
        )
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
        if self.error is not None:
            raise self.error
        if self.max_top is not None and top is not None and top > self.max_top:
            raise GraphError(
                f"Invalid page size specified: '{top}'. Must be between 1 and {self.max_top}."
            )
        if self.respond_none:
            return None
        size = top if top and top > 0 else self.page_size
        return self._page(0, size)

    def with_url(self, url: str) -> "_NextPage":
        return _NextPage(self, url)

    def _page(self, start: int, size: int):
        end = start + size
        next_link = (
            f"{GRAPH_ROOT}/{self.path}?$top={size}&$skiptoken={end}"
            if end < len(self.items)
            else None
        )
        return SimpleNamespace(value=self.items[start:end], odata_next_link=next_link)

    async def post(self, body, request_configuration=None):
        """Creates the object: Graph answers with it, under a new id when it has none."""
        self._graph.requests.append(Request("post", self.path, body=body))
        if body.id is None:
            body.id = f"new-{len(self.items) + 1:04d}"
        self.items.append(body)
        self.by_id(body.id).value = body
        return body

    def by_id(self, entity_id: str) -> "Entity":
        if entity_id not in self._entities:
            self._entities[entity_id] = Entity(self._graph, f"{self.path}/{entity_id}")
        return self._entities[entity_id]

    by_user_id = by_group_id = by_application_id = by_service_principal_id = by_id
    by_directory_role_id = by_conditional_access_policy_id = by_managed_device_id = (
        by_id
    )
    by_password_authentication_method_id = by_id


class _NextPage:
    """The request builder ``with_url(odata_next_link)`` returns."""

    def __init__(self, collection: Collection, url: str):
        self._collection = collection
        self._url = url

    async def get(self, request_configuration=None):
        self._collection._graph.requests.append(
            Request("with_url", self._collection.path, url=self._url)
        )
        query = parse_qs(urlparse(self._url).query)
        return self._collection._page(
            int(query["$skiptoken"][0]), int(query["$top"][0])
        )


class Members(Collection):
    """``/groups/{id}/members``: the member list plus the ``$ref`` calls msgraph-sdk offers on it.

    As in msgraph-sdk, ``by_directory_object_id(id)`` returns a builder with ``ref`` only: it has
    no ``get``. ``ref.post`` takes the body msgraph-sdk types it with and Graph documents for
    ``POST /groups/{id}/members/$ref``: a ``ReferenceCreate`` whose ``@odata.id`` names
    ``https://graph.microsoft.com/v1.0/directoryObjects/{id}``; any other body is refused. Graph's refusals are kept: adding a member twice and removing a non-member fail.
    ``$filter`` is evaluated for ``id eq '<id>'`` only; any other filter is refused, never ignored.
    """

    _ID_FILTER = re.compile(r"^id eq '([^']*)'$")

    def __init__(self, graph: "FakeGraph", path: str):
        super().__init__(graph, path, max_top=999)
        self.ref = _MembersRef(self)

    async def get(self, request_configuration=None):
        query = getattr(request_configuration, "query_parameters", None)
        member_filter = getattr(query, "filter", None)
        if member_filter is None:
            return await super().get(request_configuration)
        self._graph.requests.append(Request("get", self.path, filter=member_filter))
        match = self._ID_FILTER.match(member_filter)
        if match is None:
            raise GraphError(
                f"Request_UnsupportedQuery: the fake Graph does not evaluate $filter={member_filter}"
            )
        return SimpleNamespace(
            value=[m for m in self.items if m.id == match.group(1)],
            odata_next_link=None,
        )

    def by_directory_object_id(self, member_id: str) -> "_MemberItem":
        return _MemberItem(self, member_id)

    def ids(self) -> list[str]:
        return [m.id for m in self.items]


class _MembersRef:
    """``/groups/{id}/members/$ref``."""

    def __init__(self, members: Members):
        self._members = members

    async def get(self, request_configuration=None):
        self._members._graph.requests.append(
            Request("get", f"{self._members.path}/$ref")
        )
        return SimpleNamespace(
            value=[
                f"{GRAPH_ROOT}/directoryObjects/{member_id}"
                for member_id in self._members.ids()
            ],
            odata_next_link=None,
        )

    async def post(self, body, request_configuration=None):
        self._members._graph.requests.append(
            Request("post", f"{self._members.path}/$ref", body=body)
        )
        prefix = f"{GRAPH_ROOT}/directoryObjects/"
        if not isinstance(body, ReferenceCreate) or not (
            body.odata_id or ""
        ).startswith(prefix):
            raise GraphError(
                "Request_BadRequest: the request body must be a reference with an "
                f"'@odata.id' of the form '{prefix}{{id}}'."
            )
        member_id = body.odata_id[len(prefix) :]
        if member_id in self._members.ids():
            raise GraphError(
                "Request_BadRequest: One or more added object references already exist "
                "for the following modified properties: 'members'."
            )
        self._members.items.append(DirectoryObject(id=member_id))


class _MemberItem:
    """``/groups/{id}/members/{member-id}``: msgraph-sdk offers ``ref`` here and no ``get``."""

    def __init__(self, members: Members, member_id: str):
        self.ref = _MemberRef(members, member_id)


class _MemberRef:
    """``/groups/{id}/members/{member-id}/$ref``."""

    def __init__(self, members: Members, member_id: str):
        self._members = members
        self._member_id = member_id

    async def delete(self, request_configuration=None):
        path = f"{self._members.path}/{self._member_id}/$ref"
        self._members._graph.requests.append(Request("delete", path))
        if self._member_id not in self._members.ids():
            raise GraphError(
                f"Request_ResourceNotFound: Resource '{self._member_id}' does not exist "
                "or one of its queried reference-property objects are not present."
            )
        self._members.items = [
            m for m in self._members.items if m.id != self._member_id
        ]


class Owners(Collection):
    """``/groups/{id}/owners``: records owner ``$ref`` calls; owner bodies are not checked (outside DRK-2429)."""

    def __init__(self, graph: "FakeGraph", path: str):
        super().__init__(graph, path)
        self.ref = SimpleNamespace(post=self._post_ref)

    async def _post_ref(self, body, request_configuration=None):
        self._graph.requests.append(Request("post", f"{self.path}/$ref", body=body))

    def by_directory_object_id(self, owner_id: str):
        async def delete(request_configuration=None):
            self._graph.requests.append(
                Request("delete", f"{self.path}/{owner_id}/$ref")
            )

        return SimpleNamespace(ref=SimpleNamespace(delete=delete))


class Entity:
    """A single Graph object such as ``/users/{id}``, with its navigation collections."""

    def __init__(self, graph: "FakeGraph", path: str):
        self._graph = graph
        self.path = path
        self.value: Any = None
        self.error: Exception | None = None
        self.refuse_patch = False
        self.members = Members(graph, f"{path}/members")
        self.app_role_assignments = Collection(graph, f"{path}/appRoleAssignments")
        self.oauth2_permission_grants = Collection(
            graph, f"{path}/oauth2PermissionGrants"
        )
        self.owners = Owners(graph, f"{path}/owners")
        self.member_of = Collection(graph, f"{path}/memberOf")
        self.transitive_member_of = Collection(graph, f"{path}/transitiveMemberOf")
        self.authentication = SimpleNamespace(
            methods=Collection(graph, f"{path}/authentication/methods"),
            password_methods=Collection(
                graph, f"{path}/authentication/passwordMethods"
            ),
        )
        # Intune managed device relationships (beta).
        self.detected_apps = Collection(graph, f"{path}/detectedApps")
        self.device_compliance_policy_states = Collection(
            graph, f"{path}/deviceCompliancePolicyStates"
        )
        self.device_configuration_states = Collection(
            graph, f"{path}/deviceConfigurationStates"
        )

    async def get(self, request_configuration=None):
        self._graph.requests.append(Request("get", self.path))
        if self.error is not None:
            raise self.error
        if self.value is None:
            raise GraphError(
                f"Request_ResourceNotFound: Resource '{self.path}' does not exist."
            )
        return self.value

    async def patch(self, body):
        self._graph.requests.append(Request("patch", self.path, body=body))
        if self.refuse_patch:
            password = body.password_profile.password
            raise GraphError(
                f"Request_BadRequest: The password '{password}' does not comply with the password policy."
            )

    async def delete(self, request_configuration=None):
        self._graph.requests.append(Request("delete", self.path))
        self.value = None


class FakeGraph:
    """Stands in for ``GraphClient``: serves both the v1.0 and the beta client."""

    def __init__(self):
        self.requests: list[Request] = []
        self.get_client_calls = 0
        self.failure: Exception | None = (
            None  # every request fails with this before it reaches Graph
        )
        self.users = Collection(self, "users")
        self.groups = Collection(self, "groups", max_top=999)
        self.applications = Collection(self, "applications")
        self.service_principals = Collection(self, "servicePrincipals")
        self.audit_logs = SimpleNamespace(
            directory_audits=Collection(self, "auditLogs/directoryAudits"),
            sign_ins=Collection(self, "auditLogs/signIns"),
        )
        self.directory_roles = Collection(self, "directoryRoles")
        self.identity = SimpleNamespace(
            conditional_access=SimpleNamespace(
                policies=Collection(self, "identity/conditionalAccess/policies"),
            ),
        )
        self.device_management = SimpleNamespace(
            managed_devices=Collection(self, "deviceManagement/managedDevices"),
            device_compliance_policies=Collection(
                self, "deviceManagement/deviceCompliancePolicies"
            ),
            device_configurations=Collection(
                self, "deviceManagement/deviceConfigurations"
            ),
            device_categories=Collection(self, "deviceManagement/deviceCategories"),
        )

    def get_client(self) -> "FakeGraph":
        self.get_client_calls += 1
        if self.failure is not None:
            raise self.failure
        return self

    get_beta_client = get_client

    # -- what the tests read back ------------------------------------------------
    def requests_to(self, path: str) -> list[Request]:
        return [r for r in self.requests if r.path == path]

    def requests_under(self, path: str) -> list[Request]:
        return [
            r for r in self.requests if r.path == path or r.path.startswith(path + "/")
        ]

    def password_changes(self, user_id: str) -> list[str]:
        return [
            r.body.password_profile.password
            for r in self.requests_to(f"users/{user_id}")
            if r.method == "patch"
        ]

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
            entity.error = GraphError(
                f"Authorization_RequestDenied: cannot read methods of {user_id}"
            )
        return user


def make_groups(count: int, prefix: str = "Group") -> list[Group]:
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


def make_users(count: int, prefix: str = "user") -> list[User]:
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


def input_schema(tool) -> dict:
    """A listed tool's input schema: ``input_schema`` on mcp 2.x, ``inputSchema`` before it."""
    schema = getattr(tool, "input_schema", None)
    return schema if schema is not None else tool.inputSchema


class FakeContext:
    """Records what a tool sends to the AI client."""

    def __init__(self):
        self.infos: list[str] = []
        self.warnings: list[str] = []
        self.errors: list[str] = []
        self.progress: list[tuple] = []

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
