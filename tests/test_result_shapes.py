"""Results keep their shape for the 36 tools the 0.3.0 baseline leaves out (DRK-2429 §5).

``test_tool_list_and_shapes.py`` covers 12 tools against ``result_shapes_v0_3_0.json``. This file
covers the other 36, plus a second use of two covered tools: a new group (``create_group``) and a
supplied password (``reset_user_password_direct``). The
baseline here, ``fixtures/result_shapes_7b4fb3f.json``, is a literal capture from ``src/`` at
7b4fb3f, with each tool given tenant data that fills every list and nested object it returns.
``add_group_member`` was captured with a ``$ref`` fake that also took the body 7b4fb3f posts;
its result is the one the tool returns at 7b4fb3f when the add succeeds.
"""

import json
from datetime import datetime, timezone
from pathlib import Path

import pytest
from conftest import MICROSOFT_GRAPH_APP_ID
from msgraph.generated.models.app_role import AppRole
from msgraph.generated.models.app_role_assignment import AppRoleAssignment
from msgraph.generated.models.application import Application
from msgraph.generated.models.authentication_strength_policy import (
    AuthenticationStrengthPolicy,
)
from msgraph.generated.models.conditional_access_applications import (
    ConditionalAccessApplications,
)
from msgraph.generated.models.conditional_access_condition_set import (
    ConditionalAccessConditionSet,
)
from msgraph.generated.models.conditional_access_grant_controls import (
    ConditionalAccessGrantControls,
)
from msgraph.generated.models.conditional_access_policy import ConditionalAccessPolicy
from msgraph.generated.models.conditional_access_policy_state import (
    ConditionalAccessPolicyState,
)
from msgraph.generated.models.conditional_access_session_controls import (
    ConditionalAccessSessionControls,
)
from msgraph.generated.models.conditional_access_users import ConditionalAccessUsers
from msgraph.generated.models.directory_role import DirectoryRole
from msgraph.generated.models.group import Group
from msgraph.generated.models.o_auth2_permission_grant import OAuth2PermissionGrant
from msgraph.generated.models.password_authentication_method import (
    PasswordAuthenticationMethod,
)
from msgraph.generated.models.permission_scope import PermissionScope
from msgraph.generated.models.service_principal import ServicePrincipal
from msgraph.generated.models.sign_in_frequency_session_control import (
    SignInFrequencySessionControl,
)
from msgraph.generated.models.user import User
from msgraph_beta.generated.models.detected_app import DetectedApp
from msgraph_beta.generated.models.device_category import DeviceCategory
from msgraph_beta.generated.models.device_compliance_policy_state import (
    DeviceCompliancePolicyState,
)
from msgraph_beta.generated.models.device_configuration_state import (
    DeviceConfigurationState,
)
from msgraph_beta.generated.models.windows10_compliance_policy import (
    Windows10CompliancePolicy,
)
from msgraph_beta.generated.models.windows10_general_configuration import (
    Windows10GeneralConfiguration,
)
from test_tool_list_and_shapes import TOOLS as TOOLS_0_3_0
from test_tool_list_and_shapes import a_managed_device, outline
from tool_calls import TOOL_CALLS

from msgraph_mcp_server import server

BASELINE = json.loads(
    (Path(__file__).parent / "fixtures" / "result_shapes_7b4fb3f.json").read_text()
)
WHEN = datetime(2026, 10, 20, 7, 45, tzinfo=timezone.utc)
AOIFE = "aoife@contoso.com"
ALL_STAFF = "all-staff"
PAYROLL_SYNC = "payroll-sync-object"
PAYROLL_SYNC_SP = "payroll-sync-sp"
MICROSOFT_GRAPH_SP = "sp-msgraph-object"


# -- tenant data ----------------------------------------------------------------------------


def all_staff(graph):
    group = Group(
        id=ALL_STAFF,
        display_name="All Staff",
        mail="allstaff@contoso.com",
        mail_nickname="allstaff",
        description="Everyone at Contoso",
        group_types=["Unified"],
        security_enabled=False,
        mail_enabled=True,
        visibility="Private",
        created_date_time=WHEN,
    )
    graph.groups.by_group_id(ALL_STAFF).value = group
    return group


def aoife():
    return User(
        id=AOIFE,
        odata_type="#microsoft.graph.user",
        display_name="Aoife Byrne",
        mail=AOIFE,
        user_principal_name=AOIFE,
        given_name="Aoife",
        surname="Byrne",
        job_title="Engineer",
        office_location="Dublin",
        business_phones=["+353 1 555 0100"],
        mobile_phone="+353 87 555 0100",
    )


def global_administrator():
    return DirectoryRole(
        id="role-1",
        odata_type="#microsoft.graph.directoryRole",
        display_name="Global Administrator",
        description="Can manage all aspects of Entra ID",
        role_template_id="62e90394-69f5-4237-9190-012177145e10",
    )


def payroll_sync_app(graph):
    app = Application(
        id=PAYROLL_SYNC,
        app_id="payroll-sync-app",
        display_name="Payroll Sync",
        created_date_time=WHEN,
        sign_in_audience="AzureADMyOrg",
        publisher_domain="contoso.com",
        tags=["payroll"],
    )
    graph.applications.items = [app]
    graph.applications.by_application_id(PAYROLL_SYNC).value = app
    return app


def payroll_sync_sp(graph):
    sp = ServicePrincipal(
        id=PAYROLL_SYNC_SP,
        app_id="payroll-sync-app",
        display_name="Payroll Sync",
        account_enabled=True,
        app_owner_organization_id="contoso-tenant",
        tags=["payroll"],
    )
    graph.service_principals.items = [sp]
    entity = graph.service_principals.by_service_principal_id(PAYROLL_SYNC_SP)
    entity.value = sp
    entity.app_role_assignments.items = [
        AppRoleAssignment(id="assignment-1", created_date_time=WHEN)
    ]
    entity.oauth2_permission_grants.items = [
        OAuth2PermissionGrant(id="grant-1", scope="User.Read")
    ]
    return sp


def microsoft_graph_permissions(graph):
    """Microsoft Graph's service principal: its object id differs from its app id, as in a tenant."""
    sp = ServicePrincipal(
        id=MICROSOFT_GRAPH_SP,
        app_id=MICROSOFT_GRAPH_APP_ID,
        oauth2_permission_scopes=[
            PermissionScope(
                value="User.Read",
                admin_consent_display_name="Sign in and read user profile",
                admin_consent_description="Allows users to sign in and read their profile.",
                user_consent_display_name="Sign you in and read your profile",
                user_consent_description="Allows you to sign in and read your profile.",
                is_enabled=True,
            )
        ],
        app_roles=[
            AppRole(
                value="User.Read.All",
                display_name="Read all users' full profiles",
                description="Allows the app to read user profiles.",
                is_enabled=True,
            )
        ],
    )
    graph.service_principals.items = [sp]
    graph.service_principals.by_service_principal_id(MICROSOFT_GRAPH_SP).value = sp
    return sp


def mfa_policy():
    return ConditionalAccessPolicy(
        id="policy-1",
        display_name="Require MFA for All Staff",
        state=ConditionalAccessPolicyState.Enabled,
        created_date_time=WHEN,
        modified_date_time=WHEN,
        conditions=ConditionalAccessConditionSet(
            users=ConditionalAccessUsers(
                include_groups=[ALL_STAFF], exclude_users=["break-glass"]
            ),
            applications=ConditionalAccessApplications(include_applications=["All"]),
            client_app_types=[],
        ),
        grant_controls=ConditionalAccessGrantControls(
            operator="OR",
            built_in_controls=[],
            authentication_strength=AuthenticationStrengthPolicy(
                id="strength-1", display_name="Multifactor authentication"
            ),
        ),
        session_controls=ConditionalAccessSessionControls(
            sign_in_frequency=SignInFrequencySessionControl(value=8, is_enabled=True)
        ),
    )


def laptop_0042(graph):
    device = graph.device_management.managed_devices.by_managed_device_id("device-1")
    device.value = a_managed_device()
    return device


# -- one use per tool -----------------------------------------------------------------------


async def use(graph, ctx, name):
    """Calls the tool with its arguments from ``tool_calls.py``, as an AI client would."""
    return await getattr(server, name)(ctx=ctx, **TOOL_CALLS[name])


async def member_add(graph, ctx):
    all_staff(graph)
    return await use(graph, ctx, "add_group_member")


async def owner_add(graph, ctx):
    return await use(graph, ctx, "add_group_owner")


async def application_creation(graph, ctx):
    return await use(graph, ctx, "create_application")


async def group_creation(graph, ctx):
    return await use(graph, ctx, "create_group")


async def service_principal_creation(graph, ctx):
    return await use(graph, ctx, "create_service_principal")


async def application_deletion(graph, ctx):
    payroll_sync_app(graph)
    return await use(graph, ctx, "delete_application")


async def group_deletion(graph, ctx):
    all_staff(graph)
    return await use(graph, ctx, "delete_group")


async def service_principal_deletion(graph, ctx):
    payroll_sync_sp(graph)
    return await use(graph, ctx, "delete_service_principal")


async def graph_permission_list(graph, ctx):
    microsoft_graph_permissions(graph)
    return await use(graph, ctx, "get_all_graph_permissions")


async def conditional_access_policy_list(graph, ctx):
    all_staff(graph)
    graph.identity.conditional_access.policies.items = [mfa_policy()]
    return await use(graph, ctx, "get_conditional_access_policies")


async def conditional_access_policy_details(graph, ctx):
    all_staff(graph)
    policies = graph.identity.conditional_access.policies
    policies.by_conditional_access_policy_id("policy-1").value = mfa_policy()
    return await use(graph, ctx, "get_conditional_access_policy_by_id")


async def detected_apps(graph, ctx):
    laptop_0042(graph).detected_apps.items = [
        DetectedApp(id="app-1", display_name="Contoso VPN", version="2.1.0")
    ]
    return await use(graph, ctx, "get_detected_apps_for_device")


async def device_category_list(graph, ctx):
    graph.device_management.device_categories.items = [
        DeviceCategory(id="category-1", display_name="Corporate")
    ]
    return await use(graph, ctx, "get_device_categories")


async def device_compliance_policy_list(graph, ctx):
    graph.device_management.device_compliance_policies.items = [
        Windows10CompliancePolicy(id="compliance-1", display_name="Windows baseline")
    ]
    return await use(graph, ctx, "get_device_compliance_policies")


async def device_compliance_policy_states(graph, ctx):
    laptop_0042(graph).device_compliance_policy_states.items = [
        DeviceCompliancePolicyState(id="state-1", display_name="Windows baseline")
    ]
    return await use(graph, ctx, "get_device_compliance_policy_states")


async def device_configuration_states(graph, ctx):
    laptop_0042(graph).device_configuration_states.items = [
        DeviceConfigurationState(id="state-1", display_name="Wi-Fi profile")
    ]
    return await use(graph, ctx, "get_device_configuration_states")


async def device_configuration_list(graph, ctx):
    graph.device_management.device_configurations.items = [
        Windows10GeneralConfiguration(
            id="configuration-1", display_name="Wi-Fi profile"
        )
    ]
    return await use(graph, ctx, "get_device_configurations")


async def group_details(graph, ctx):
    all_staff(graph)
    return await use(graph, ctx, "get_group_by_id")


async def managed_device_details(graph, ctx):
    laptop_0042(graph)
    return await use(graph, ctx, "get_managed_device_by_id")


async def privileged_user_list(graph, ctx):
    graph.directory_roles.items = [global_administrator()]
    graph.directory_roles.by_directory_role_id("role-1").members.items = [aoife()]
    return await use(graph, ctx, "get_privileged_users")


async def service_principal_details(graph, ctx):
    payroll_sync_sp(graph)
    return await use(graph, ctx, "get_service_principal_by_id")


async def user_details(graph, ctx):
    graph.users.by_user_id(AOIFE).value = aoife()
    return await use(graph, ctx, "get_user_by_id")


async def user_groups(graph, ctx):
    all_staff(graph)
    graph.users.by_user_id(AOIFE).transitive_member_of.items = [
        Group(id=ALL_STAFF, odata_type="#microsoft.graph.group")
    ]
    return await use(graph, ctx, "get_user_groups")


async def user_mfa_status(graph, ctx):
    graph.add_user(AOIFE)
    return await use(graph, ctx, "get_user_mfa_status")


async def password_method_details(graph, ctx):
    methods = graph.users.by_user_id(AOIFE).authentication.password_methods
    methods.by_password_authentication_method_id(
        "password-1"
    ).value = PasswordAuthenticationMethod(id="password-1", created_date_time=WHEN)
    return await use(graph, ctx, "get_user_password_method")


async def user_roles(graph, ctx):
    graph.users.by_user_id(AOIFE).member_of.items = [global_administrator()]
    graph.directory_roles.by_directory_role_id("role-1").value = global_administrator()
    return await use(graph, ctx, "get_user_roles")


async def application_list(graph, ctx):
    payroll_sync_app(graph)
    return await use(graph, ctx, "list_applications")


async def permission_categories_and_tasks(graph, ctx):
    return await use(graph, ctx, "list_permission_categories_and_tasks")


async def service_principal_list(graph, ctx):
    payroll_sync_sp(graph)
    return await use(graph, ctx, "list_service_principals")


async def password_method_list(graph, ctx):
    graph.users.by_user_id(AOIFE).authentication.password_methods.items = [
        PasswordAuthenticationMethod(id="password-1", created_date_time=WHEN)
    ]
    return await use(graph, ctx, "list_user_password_methods")


async def member_removal(graph, ctx):
    all_staff(graph)
    graph.groups.by_group_id(ALL_STAFF).members.items = [User(id=AOIFE)]
    return await use(graph, ctx, "remove_group_member")


async def owner_removal(graph, ctx):
    return await use(graph, ctx, "remove_group_owner")


async def password_reset_with_a_supplied_password(graph, ctx):
    graph.add_user(AOIFE)
    return await use(graph, ctx, "reset_user_password_direct")


async def permission_search(graph, ctx):
    microsoft_graph_permissions(graph)
    return await use(graph, ctx, "search_permissions")


async def permission_suggestion(graph, ctx):
    return await use(graph, ctx, "suggest_permissions_for_task")


async def application_update(graph, ctx):
    payroll_sync_app(graph)
    payroll_sync_sp(graph)
    return await use(graph, ctx, "update_application")


async def group_update(graph, ctx):
    all_staff(graph)
    return await use(graph, ctx, "update_group")


async def service_principal_update(graph, ctx):
    payroll_sync_sp(graph)
    return await use(graph, ctx, "update_service_principal")


TOOLS = {
    "add_group_member": member_add,
    "add_group_owner": owner_add,
    "create_application": application_creation,
    "create_group": group_creation,
    "create_service_principal": service_principal_creation,
    "delete_application": application_deletion,
    "delete_group": group_deletion,
    "delete_service_principal": service_principal_deletion,
    "get_all_graph_permissions": graph_permission_list,
    "get_conditional_access_policies": conditional_access_policy_list,
    "get_conditional_access_policy_by_id": conditional_access_policy_details,
    "get_detected_apps_for_device": detected_apps,
    "get_device_categories": device_category_list,
    "get_device_compliance_policies": device_compliance_policy_list,
    "get_device_compliance_policy_states": device_compliance_policy_states,
    "get_device_configuration_states": device_configuration_states,
    "get_device_configurations": device_configuration_list,
    "get_group_by_id": group_details,
    "get_managed_device_by_id": managed_device_details,
    "get_privileged_users": privileged_user_list,
    "get_service_principal_by_id": service_principal_details,
    "get_user_by_id": user_details,
    "get_user_groups": user_groups,
    "get_user_mfa_status": user_mfa_status,
    "get_user_password_method": password_method_details,
    "get_user_roles": user_roles,
    "list_applications": application_list,
    "list_permission_categories_and_tasks": permission_categories_and_tasks,
    "list_service_principals": service_principal_list,
    "list_user_password_methods": password_method_list,
    "remove_group_member": member_removal,
    "remove_group_owner": owner_removal,
    "reset_user_password_direct": password_reset_with_a_supplied_password,
    "search_permissions": permission_search,
    "suggest_permissions_for_task": permission_suggestion,
    "update_application": application_update,
    "update_group": group_update,
    "update_service_principal": service_principal_update,
}

TOOLS_IN_0_3_0_BASELINE = {
    "reset_user_password_direct",
    "search_users",
    "get_all_groups",
    "search_groups_by_name",
    "get_group_members",
    "create_group",
    "get_group_mfa_status",
    "get_all_managed_devices",
    "get_managed_devices_by_user",
    "get_user_audit_logs",
    "get_user_sign_ins",
    "get_application_by_id",
}


def test_the_two_baselines_cover_all_48_tools():
    assert len(TOOLS_0_3_0) == 12
    assert sorted(TOOLS_IN_0_3_0_BASELINE | set(TOOLS)) == sorted(TOOL_CALLS)
    assert sorted(BASELINE) == sorted(TOOLS)


@pytest.mark.parametrize("name", sorted(TOOLS))
async def test_results_keep_their_shape(graph, ctx, name):
    result = await TOOLS[name](graph, ctx)

    assert result
    assert outline(result) == BASELINE[name]
