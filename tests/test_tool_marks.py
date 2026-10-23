"""Read-only and destructive marks on every tool (DRK-2434 §5, §3a; DRK-2433 note 1).

An AI client reads the marks from each listed tool's ``annotations``. Every check reads
them in their wire form (``readOnlyHint``, ``destructiveHint``), the way the client gets
them. The expected marks are literals from the spec, one row per spec name.
"""

import json
from pathlib import Path

import pytest
from fastmcp import Client
from server_process import Machine
from test_tool_list import list_tools

from msgraph_mcp_server import server

DESCRIPTIONS_7949FB4 = json.loads(
    (Path(__file__).parent / "fixtures" / "tool_descriptions_7949fb4.json").read_text()
)


def marks(annotations):
    """A listed tool's annotations as the AI client gets them; no annotations is no marks."""
    if annotations is None:
        return {}
    if isinstance(annotations, dict):
        return {name: value for name, value in annotations.items() if value is not None}
    return annotations.model_dump(mode="json", by_alias=True, exclude_none=True)


async def listed_tools():
    async with Client(server.mcp) as client:
        return await client.list_tools()


async def marks_by_tool():
    return {tool.name: marks(tool.annotations) for tool in await listed_tools()}


@pytest.mark.parametrize(
    "tool",
    [
        pytest.param("search_users", id="user search · search_users"),
        pytest.param("get_user_by_id", id="user details · get_user_by_id"),
        pytest.param("get_user_sign_ins", id="user sign-in list · get_user_sign_ins"),
        pytest.param("get_user_mfa_status", id="user MFA status · get_user_mfa_status"),
        pytest.param(
            "get_group_mfa_status", id="group MFA status · get_group_mfa_status"
        ),
        pytest.param(
            "get_privileged_users", id="privileged user list · get_privileged_users"
        ),
        pytest.param(
            "get_conditional_access_policies",
            id="conditional access policy list · get_conditional_access_policies",
        ),
        pytest.param(
            "get_conditional_access_policy_by_id",
            id="conditional access policy details · get_conditional_access_policy_by_id",
        ),
        pytest.param("get_all_groups", id="group list · get_all_groups"),
        pytest.param("get_group_by_id", id="group details · get_group_by_id"),
        pytest.param(
            "search_groups_by_name", id="group search by name · search_groups_by_name"
        ),
        pytest.param("get_group_members", id="group member list · get_group_members"),
        pytest.param("get_user_groups", id="user group list · get_user_groups"),
        pytest.param("get_user_roles", id="user role list · get_user_roles"),
        pytest.param(
            "get_all_managed_devices",
            id="managed device list · get_all_managed_devices",
        ),
        pytest.param(
            "get_managed_devices_by_user",
            id="managed device list of a user · get_managed_devices_by_user",
        ),
        pytest.param(
            "get_managed_device_by_id",
            id="managed device details · get_managed_device_by_id",
        ),
        pytest.param(
            "get_detected_apps_for_device",
            id="detected app list of a device · get_detected_apps_for_device",
        ),
        pytest.param(
            "get_device_compliance_policy_states",
            id="compliance policy state list of a device · get_device_compliance_policy_states",
        ),
        pytest.param(
            "get_device_configuration_states",
            id="configuration state list of a device · get_device_configuration_states",
        ),
        pytest.param(
            "get_device_compliance_policies",
            id="device compliance policy list · get_device_compliance_policies",
        ),
        pytest.param(
            "get_device_configurations",
            id="device configuration list · get_device_configurations",
        ),
        pytest.param(
            "get_device_categories", id="device category list · get_device_categories"
        ),
        pytest.param(
            "get_user_audit_logs", id="user audit log list · get_user_audit_logs"
        ),
        pytest.param(
            "list_user_password_methods",
            id="user password method list · list_user_password_methods",
        ),
        pytest.param(
            "get_user_password_method",
            id="user password method details · get_user_password_method",
        ),
        pytest.param(
            "suggest_permissions_for_task",
            id="permission suggestion for a task · suggest_permissions_for_task",
        ),
        pytest.param(
            "list_permission_categories_and_tasks",
            id="permission category and task list · list_permission_categories_and_tasks",
        ),
        pytest.param(
            "get_all_graph_permissions",
            id="Microsoft Graph permission list · get_all_graph_permissions",
        ),
        pytest.param("search_permissions", id="permission search · search_permissions"),
        pytest.param("list_applications", id="application list · list_applications"),
        pytest.param(
            "get_application_by_id", id="application details · get_application_by_id"
        ),
        pytest.param(
            "list_service_principals",
            id="service principal list · list_service_principals",
        ),
        pytest.param(
            "get_service_principal_by_id",
            id="service principal details · get_service_principal_by_id",
        ),
    ],
)
async def test_tools_that_only_read_state_that_they_are_read_only(tool):
    tool_marks = (await marks_by_tool())[tool]

    assert tool_marks.get("readOnlyHint") is True
    # MCP ignores the destructive mark on a read-only tool, so none is set (DRK-2434 §3a).
    assert "destructiveHint" not in tool_marks


CHANGING_TOOLS = [
    pytest.param("delete_group", id="group deletion · delete_group"),
    pytest.param("delete_application", id="application deletion · delete_application"),
    pytest.param(
        "delete_service_principal",
        id="service principal deletion · delete_service_principal",
    ),
    pytest.param("remove_group_member", id="member removal · remove_group_member"),
    pytest.param("remove_group_owner", id="owner removal · remove_group_owner"),
    pytest.param(
        "reset_user_password_direct", id="password reset · reset_user_password_direct"
    ),
    pytest.param("update_group", id="group update · update_group"),
    pytest.param("update_application", id="application update · update_application"),
    pytest.param(
        "update_service_principal",
        id="service principal update · update_service_principal",
    ),
    pytest.param("create_group", id="group creation · create_group"),
    pytest.param("create_application", id="application creation · create_application"),
    pytest.param(
        "create_service_principal",
        id="service principal creation · create_service_principal",
    ),
    pytest.param("add_group_member", id="member add · add_group_member"),
    pytest.param("add_group_owner", id="owner add · add_group_owner"),
]


@pytest.mark.parametrize("tool", CHANGING_TOOLS)
async def test_tools_that_change_the_tenant_state_that_they_are_not_read_only(tool):
    tool_marks = (await marks_by_tool())[tool]

    assert tool_marks.get("readOnlyHint") is False


@pytest.mark.parametrize(
    "tool, destructive",
    [
        pytest.param("delete_group", True, id="group deletion · delete_group"),
        pytest.param(
            "delete_application", True, id="application deletion · delete_application"
        ),
        pytest.param(
            "delete_service_principal",
            True,
            id="service principal deletion · delete_service_principal",
        ),
        pytest.param(
            "remove_group_member", True, id="member removal · remove_group_member"
        ),
        pytest.param(
            "remove_group_owner", True, id="owner removal · remove_group_owner"
        ),
        pytest.param(
            "reset_user_password_direct",
            True,
            id="password reset · reset_user_password_direct",
        ),
        pytest.param("update_group", True, id="group update · update_group"),
        pytest.param(
            "update_application", True, id="application update · update_application"
        ),
        pytest.param(
            "update_service_principal",
            True,
            id="service principal update · update_service_principal",
        ),
        pytest.param("create_group", False, id="group creation · create_group"),
        pytest.param(
            "create_application", False, id="application creation · create_application"
        ),
        pytest.param(
            "create_service_principal",
            False,
            id="service principal creation · create_service_principal",
        ),
        pytest.param("add_group_member", False, id="member add · add_group_member"),
        pytest.param("add_group_owner", False, id="owner add · add_group_owner"),
    ],
)
async def test_tools_that_change_the_tenant_state_whether_the_change_is_destructive(
    tool, destructive
):
    tool_marks = (await marks_by_tool())[tool]

    assert tool_marks.get("destructiveHint") is destructive


async def wire_marks(transport, tmp_path):
    """Each listed tool's marks, through the in-process client or raw over stdio."""
    if transport == "in-process client":
        return list((await marks_by_tool()).values())
    response = await list_tools(transport, Machine(tmp_path))
    assert "error" not in response
    return [marks(tool.get("annotations")) for tool in response["result"]["tools"]]


@pytest.mark.parametrize("transport", ["in-process client", "2025-11-25", "2026-07-28"])
async def test_every_tool_states_whether_it_is_read_only(transport, tmp_path):
    every_tool = await wire_marks(transport, tmp_path)

    assert len(every_tool) == 48
    for tool_marks in every_tool:
        assert isinstance(tool_marks.get("readOnlyHint"), bool)
        # No MCP hint other than read-only and destructive (DRK-2434 §3a).
        assert "idempotentHint" not in tool_marks
        assert "openWorldHint" not in tool_marks
        assert "title" not in tool_marks


async def test_every_tool_that_is_not_read_only_states_whether_it_is_destructive():
    not_read_only = [
        tool_marks
        for tool_marks in (await marks_by_tool()).values()
        if tool_marks.get("readOnlyHint") is False
    ]

    # Guard: with no tool marked "not read-only" the check below would pass on nothing.
    assert not_read_only
    for tool_marks in not_read_only:
        assert isinstance(tool_marks.get("destructiveHint"), bool)


async def test_tool_descriptions_do_not_change():
    listed = {tool.name: tool.description for tool in await listed_tools()}

    assert listed == DESCRIPTIONS_7949FB4
