"""The tool list and the result shapes stay as in release 0.3.0 (DRK-2421 §5, DRK-2420 finding 2).

The baselines in ``tests/fixtures`` are literal files written from release 0.3.0
(``src/`` at base 3ad92a2 equals tag ``v0.3.0``). A result's shape is its outline:
every dict key, nested dicts and the first list item expanded, leaf values dropped.
"""

import json
from datetime import datetime, timezone
from pathlib import Path

import pytest
from conftest import make_groups, make_users
from fastmcp import Client
from msgraph.generated.models.app_identity import AppIdentity
from msgraph.generated.models.app_role_assignment import AppRoleAssignment
from msgraph.generated.models.application import Application
from msgraph.generated.models.audit_activity_initiator import AuditActivityInitiator
from msgraph.generated.models.device_detail import DeviceDetail
from msgraph.generated.models.directory_audit import DirectoryAudit
from msgraph.generated.models.geo_coordinates import GeoCoordinates
from msgraph.generated.models.key_value import KeyValue
from msgraph.generated.models.modified_property import ModifiedProperty
from msgraph.generated.models.o_auth2_permission_grant import OAuth2PermissionGrant
from msgraph.generated.models.service_principal import ServicePrincipal
from msgraph.generated.models.sign_in import SignIn
from msgraph.generated.models.sign_in_location import SignInLocation
from msgraph.generated.models.sign_in_status import SignInStatus
from msgraph.generated.models.target_resource import TargetResource
from msgraph.generated.models.user import User
from msgraph.generated.models.user_identity import UserIdentity
from msgraph_beta.generated.models.hardware_information import HardwareInformation
from msgraph_beta.generated.models.managed_device import ManagedDevice

from msgraph_mcp_server import server

FIXTURES = Path(__file__).parent / "fixtures"
WHEN = datetime(2026, 10, 20, 7, 45, tzinfo=timezone.utc)


def baseline(name):
    return json.loads((FIXTURES / name).read_text())


def outline(value):
    if isinstance(value, dict):
        return {key: outline(item) for key, item in value.items()}
    if isinstance(value, list) and value and isinstance(value[0], (dict, list)):
        return [outline(value[0])]
    return None


async def test_the_tool_list_does_not_change():
    async with Client(server.mcp) as client:
        tools = await client.list_tools()

    listed = {
        tool.name: {
            name: {
                "required": name in tool.inputSchema.get("required", []),
                "default": schema.get("default"),
            }
            for name, schema in tool.inputSchema["properties"].items()
        }
        for tool in tools
    }
    assert len(listed) == 48
    assert listed == baseline("tools_v0_3_0.json")


# -- tenant data with every optional branch filled, so each outline is complete -------------


def a_managed_device(user_id="7f3c9a"):
    return ManagedDevice(
        id="device-1",
        device_name="LAPTOP-0042",
        user_id=user_id,
        operating_system="Windows",
        enrolled_date_time=WHEN,
        last_sync_date_time=WHEN,
        hardware_information=HardwareInformation(
            serial_number="SN-0042", manufacturer="Contoso"
        ),
    )


def an_audit_entry():
    return DirectoryAudit(
        id="audit-1",
        activity_date_time=WHEN,
        activity_display_name="Update user",
        additional_details=[KeyValue(key="UserType", value="Member")],
        initiated_by=AuditActivityInitiator(
            user=UserIdentity(
                id="7f3c9a", display_name="Mai", user_principal_name="mai@contoso.com"
            ),
            app=AppIdentity(app_id="app-1", display_name="Portal"),
        ),
        target_resources=[
            TargetResource(
                id="aoife",
                display_name="Aoife",
                modified_properties=[
                    ModifiedProperty(
                        display_name="JobTitle", old_value="a", new_value="b"
                    )
                ],
            )
        ],
    )


def a_sign_in():
    return SignIn(
        id="sign-in-1",
        created_date_time=WHEN,
        user_id="7f3c9a",
        status=SignInStatus(error_code=0),
        device_detail=DeviceDetail(device_id="device-1", display_name="LAPTOP-0042"),
        location=SignInLocation(
            city="Dublin",
            country_or_region="IE",
            geo_coordinates=GeoCoordinates(latitude=53.35, longitude=-6.26),
        ),
    )


async def password_reset_with_a_generated_password(graph, ctx):
    graph.add_user("aoife@contoso.com")
    return await server.reset_user_password_direct(
        "aoife@contoso.com", ctx, generate_password=True
    )


async def user_search(graph, ctx):
    graph.users.items = make_users(3)
    return await server.search_users("aoife", ctx)


async def group_list(graph, ctx):
    graph.groups.items = make_groups(3)
    return await server.get_all_groups(ctx)


async def group_search(graph, ctx):
    graph.groups.items = make_groups(3, prefix="Fin")
    return await server.search_groups_by_name("Fin", ctx)


async def group_member_list(graph, ctx):
    graph.groups.by_group_id("all-staff").members.items = make_users(3)
    return await server.get_group_members("all-staff", ctx)


async def group_creation_when_the_name_already_exists(graph, ctx):
    graph.groups.items = make_groups(1)
    return await server.create_group(
        ctx, {"displayName": "Group 0000", "mailNickname": "group0000"}
    )


async def group_mfa_report(graph, ctx):
    graph.groups.by_group_id("all-staff").members.items = [
        User(id=graph.add_user("aoife@contoso.com").id)
    ]
    return await server.get_group_mfa_status("all-staff", ctx)


async def managed_device_list(graph, ctx):
    graph.device_management.managed_devices.items = [a_managed_device()]
    return await server.get_all_managed_devices(ctx)


async def managed_devices_by_user(graph, ctx):
    graph.device_management.managed_devices.items = [a_managed_device()]
    return await server.get_managed_devices_by_user("7f3c9a", ctx)


async def audit_logs_by_user(graph, ctx):
    graph.audit_logs.directory_audits.items = [an_audit_entry()]
    return await server.get_user_audit_logs("7f3c9a", ctx)


async def sign_in_logs_by_user(graph, ctx):
    graph.audit_logs.sign_ins.items = [a_sign_in()]
    return await server.get_user_sign_ins("7f3c9a", ctx)


async def application_details(graph, ctx):
    graph.applications.by_application_id("payroll-sync-object").value = Application(
        id="payroll-sync-object",
        app_id="payroll-sync-app",
        display_name="Payroll Sync",
        created_date_time=WHEN,
    )
    graph.service_principals.items = [
        ServicePrincipal(id="payroll-sync-sp", app_id="payroll-sync-app")
    ]
    sp = graph.service_principals.by_service_principal_id("payroll-sync-sp")
    sp.app_role_assignments.items = [AppRoleAssignment(id="assignment-1")]
    sp.oauth2_permission_grants.items = [OAuth2PermissionGrant(id="grant-1")]
    return await server.get_application_by_id("payroll-sync-object", ctx)


TOOLS = [
    pytest.param(
        password_reset_with_a_generated_password,
        id="password reset with a generated password",
    ),
    pytest.param(user_search, id="user search"),
    pytest.param(group_list, id="group list"),
    pytest.param(group_search, id="group search"),
    pytest.param(group_member_list, id="group member list"),
    pytest.param(
        group_creation_when_the_name_already_exists,
        id="group creation when the name already exists",
    ),
    pytest.param(group_mfa_report, id="group MFA report"),
    pytest.param(managed_device_list, id="managed device list"),
    pytest.param(managed_devices_by_user, id="managed devices by user"),
    pytest.param(audit_logs_by_user, id="audit logs by user"),
    pytest.param(sign_in_logs_by_user, id="sign-in logs by user"),
    pytest.param(application_details, id="application details"),
]


@pytest.mark.parametrize("use_tool", TOOLS)
async def test_results_keep_their_shape(graph, ctx, request, use_tool):
    result = await use_tool(graph, ctx)

    assert result
    assert (
        outline(result)
        == baseline("result_shapes_v0_3_0.json")[request.node.callspec.id]
    )
