"""A failed tool call keeps both error texts and is logged once (DRK-2429 §5, DRK-2432 brief §6 R1).

``fixtures/tool_errors_7b4fb3f.json`` is a literal capture from ``src/`` at 7b4fb3f, before this
change: for each of the 48 tools, called once with the arguments in ``tool_calls.py`` while the
work behind it fails with "Service unavailable", the error notifications the AI client received
and the text of the call's error result.
"""

import json
import logging
from pathlib import Path

import pytest
from conftest import GraphError
from fastmcp import Client
from tool_calls import AOIFE, TOOL_CALLS

from msgraph_mcp_server import server
from msgraph_mcp_server.auth.graph_auth import AuthenticationError
from msgraph_mcp_server.resources import permissions_helper

BASELINE = json.loads(
    (Path(__file__).parent / "fixtures" / "tool_errors_7b4fb3f.json").read_text()
)


async def call_once(name, arguments):
    """Calls one tool through an MCP client; returns its error notifications and its error result text."""
    notifications = []

    async def on_log(message):
        if message.level == "error":
            data = message.data
            notifications.append(data["msg"] if isinstance(data, dict) else data)

    async with Client(server.mcp, log_handler=on_log) as client:
        result = await client.call_tool(name, arguments, raise_on_error=False)
    assert result.is_error
    [content] = result.content
    return notifications, content.text


@pytest.fixture
def service_unavailable(graph, monkeypatch):
    """The work behind every tool fails: each Graph request, and the two tools that read no Graph data."""
    graph.failure = GraphError("Service unavailable")

    async def unavailable(*args, **kwargs):
        raise GraphError("Service unavailable")

    monkeypatch.setattr(permissions_helper, "suggest_permissions_for_task", unavailable)
    monkeypatch.setattr(
        permissions_helper, "list_available_categories_and_tasks", unavailable
    )


ALL_TOOLS = sorted(TOOL_CALLS)


def test_the_baseline_covers_every_tool():
    assert len(ALL_TOOLS) == 48
    assert sorted(BASELINE) == ALL_TOOLS


@pytest.mark.parametrize("name", ALL_TOOLS)
async def test_every_tool_keeps_its_error_notification(service_unavailable, name):
    notifications, _ = await call_once(name, TOOL_CALLS[name])

    assert notifications == BASELINE[name]["notifications"]


@pytest.mark.parametrize("name", ALL_TOOLS)
async def test_every_tool_keeps_its_error_result(service_unavailable, name):
    _, error_result = await call_once(name, TOOL_CALLS[name])

    assert error_result == BASELINE[name]["error_result"]


@pytest.mark.parametrize("name", ALL_TOOLS)
async def test_every_failed_tool_call_is_logged_once(service_unavailable, caplog, name):
    caplog.set_level(logging.DEBUG)

    await call_once(name, TOOL_CALLS[name])

    errors = [
        record.getMessage()
        for record in caplog.records
        if record.levelno >= logging.ERROR
    ]
    assert errors == [f"Error calling tool '{name}'"]


# -- Scenario Outline: A failed tool call keeps both error texts --------------------------------


def the_contoso_tenant_is_reachable(graph):
    pass


def sign_in_fails_with_no_azure_login_found(graph):
    graph.failure = AuthenticationError("no Azure login found")


def graph_refuses_the_sign_in_logs(graph):
    graph.audit_logs.sign_ins.error = GraphError(
        "Authorization_RequestDenied: Insufficient privileges"
    )


def graph_fails_with_connection_refused(graph):
    graph.failure = ConnectionError("Connection refused")


def graph_refuses_the_lookup_of_group_g_404(graph):
    graph.groups.by_group_id("g-404").error = GraphError("Request_ResourceNotFound")


FAILED_CALLS = [
    pytest.param(
        "create_group",
        {"group_data": {"mailNickname": "allstaff"}},
        the_contoso_tenant_is_reachable,
        "Error creating group: displayName is required for creating a group",
        "displayName is required for creating a group",
        id="group creation with no display name",
    ),
    pytest.param(
        "search_users",
        {"query": "Aoife"},
        sign_in_fails_with_no_azure_login_found,
        "Authentication error: no Azure login found",
        "no Azure login found",
        id='user search for "Aoife"',
    ),
    pytest.param(
        "get_user_sign_ins",
        {"user_id": AOIFE},
        graph_refuses_the_sign_in_logs,
        "Error fetching sign-in logs for aoife@contoso.com: Authorization_RequestDenied: Insufficient privileges"
        " (Ensure the application has AuditLog.Read.All permission)",
        "Authorization_RequestDenied: Insufficient privileges",
        id="sign-in log list for aoife@contoso.com",
    ),
    pytest.param(
        "get_privileged_users",
        {},
        graph_fails_with_connection_refused,
        "Error fetching privileged users: Connection refused",
        "Connection refused",
        id="privileged user list",
    ),
    pytest.param(
        "add_group_member",
        {"group_id": "g-404", "member_id": AOIFE},
        graph_refuses_the_lookup_of_group_g_404,
        "Error adding member aoife@contoso.com to group g-404: Request_ResourceNotFound",
        "Request_ResourceNotFound",
        id="member add of aoife@contoso.com to group g-404",
    ),
]


@pytest.mark.parametrize(
    "name, arguments, situation, notification, reason", FAILED_CALLS
)
async def test_a_failed_tool_call_keeps_both_error_texts(
    graph, name, arguments, situation, notification, reason
):
    situation(graph)

    notifications, error_result = await call_once(name, arguments)

    assert notifications == [notification]
    assert error_result == f"Error calling tool '{name}': {reason}"
