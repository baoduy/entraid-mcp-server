"""Application details and sign-in logs (DRK-2421 §5, DRK-2425 §7 and §6a D4–D5)."""

from msgraph.generated.models.app_role_assignment import AppRoleAssignment
from msgraph.generated.models.application import Application
from msgraph.generated.models.o_auth2_permission_grant import OAuth2PermissionGrant
from msgraph.generated.models.service_principal import ServicePrincipal
from msgraph.generated.models.sign_in import SignIn

from msgraph_mcp_server import server

PAYROLL_SYNC = "payroll-sync-object"
PAYROLL_SYNC_APP_ID = "5d2e0b41-0000-4000-8000-000000000001"


def payroll_sync_exists(graph):
    graph.applications.by_application_id(PAYROLL_SYNC).value = Application(
        id=PAYROLL_SYNC, app_id=PAYROLL_SYNC_APP_ID, display_name="Payroll Sync"
    )


async def test_application_details_include_the_service_principals_assignments_and_grants(graph, ctx):
    payroll_sync_exists(graph)
    graph.service_principals.items = [ServicePrincipal(id="payroll-sync-sp", app_id=PAYROLL_SYNC_APP_ID)]
    sp = graph.service_principals.by_service_principal_id("payroll-sync-sp")
    sp.app_role_assignments.items = [AppRoleAssignment(id="assignment-1"), AppRoleAssignment(id="assignment-2")]
    sp.oauth2_permission_grants.items = [OAuth2PermissionGrant(id="grant-1")]

    details = await server.get_application_by_id(PAYROLL_SYNC, ctx)

    assert [a["id"] for a in details["appRoleAssignments"]] == ["assignment-1", "assignment-2"]
    assert [g["id"] for g in details["oauth2PermissionGrants"]] == ["grant-1"]


async def test_application_details_without_a_service_principal_have_empty_assignments_and_grants(graph, ctx):
    payroll_sync_exists(graph)

    details = await server.get_application_by_id(PAYROLL_SYNC, ctx)

    assert details["appRoleAssignments"] == []
    assert details["oauth2PermissionGrants"] == []


def sign_ins(count):
    return [SignIn(id=f"sign-in-{i:04d}", user_id="aoife@contoso.com") for i in range(count)]


async def test_sign_in_logs_cover_every_page_within_the_asked_days(graph, ctx):
    graph.audit_logs.sign_ins.items = sign_ins(2400)

    logs = await server.get_user_sign_ins("aoife@contoso.com", ctx, days=7)

    assert len(logs) == 2400


async def test_sign_in_logs_on_one_page_come_back_whole(graph, ctx):
    graph.audit_logs.sign_ins.items = sign_ins(40)

    logs = await server.get_user_sign_ins("aoife@contoso.com", ctx, days=7)

    assert [log["id"] for log in logs] == [f"sign-in-{i:04d}" for i in range(40)]
