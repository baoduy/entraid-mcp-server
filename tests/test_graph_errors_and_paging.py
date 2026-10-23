"""Changed resource functions: Graph errors propagate, and every page is read (DRK-2425 §8 coverage unit)."""

import pytest
from conftest import GraphError, make_users
from msgraph.generated.models.directory_audit import DirectoryAudit
from test_empty_pages import Pages

from msgraph_mcp_server.resources import (
    audit_logs,
    managed_devices,
    service_principals,
    users,
)


def graph_fails(monkeypatch, collection):
    async def get(request_configuration=None):
        raise GraphError("ServiceUnavailable: try again later")

    monkeypatch.setattr(collection, "get", get)


@pytest.mark.parametrize(
    "collection, call",
    [
        (lambda g: g.users, lambda g: users.search_users(g, "Aoife")),
        (
            lambda g: g.device_management.managed_devices,
            lambda g: managed_devices.get_managed_devices_by_user(g, "7f3c9a"),
        ),
        (
            lambda g: g.audit_logs.directory_audits,
            lambda g: audit_logs.get_user_audit_logs(g, "7f3c9a"),
        ),
        (
            lambda g: g.service_principals,
            lambda g: service_principals.get_service_principal_by_app_id(g, "5d2e41"),
        ),
    ],
    ids=[
        "user search",
        "managed devices by user",
        "audit logs by user",
        "service principal by app id",
    ],
)
async def test_a_graph_error_reaches_the_caller(graph, monkeypatch, collection, call):
    graph_fails(monkeypatch, collection(graph))

    with pytest.raises(GraphError, match="^ServiceUnavailable: try again later$"):
        await call(graph)


async def test_user_search_reads_the_next_page_until_the_limit(graph):
    graph.users = Pages(make_users(2), make_users(3, prefix="more"))

    result = await users.search_users(graph, "Aoife", limit=10)

    assert [u["id"] for u in result] == [
        "user-0000",
        "user-0001",
        "more-0000",
        "more-0001",
        "more-0002",
    ]


async def test_audit_logs_by_user_cover_every_page(graph):
    graph.audit_logs.directory_audits.items = [
        DirectoryAudit(id=f"audit-{i:04d}") for i in range(1200)
    ]

    logs = await audit_logs.get_user_audit_logs(graph, "7f3c9a")

    assert [log["id"] for log in logs] == [f"audit-{i:04d}" for i in range(1200)]
    assert len(graph.requests_to("auditLogs/directoryAudits")) == 2
