"""Details come back when the permission grants cannot be read (DRK-2429 §5; brief §3 row 5 keeps the warning)."""

import logging

import pytest
from conftest import GraphError
from msgraph.generated.models.application import Application
from msgraph.generated.models.service_principal import ServicePrincipal

from msgraph_mcp_server import server

PAYROLL_SYNC_APP_ID = "5d2e0b41-0000-4000-8000-000000000001"
PAYROLL_SYNC_SP = "payroll-sync-sp"


def grants_cannot_be_read(graph):
    sp = graph.service_principals.by_service_principal_id(PAYROLL_SYNC_SP)
    sp.value = ServicePrincipal(
        id=PAYROLL_SYNC_SP, app_id=PAYROLL_SYNC_APP_ID, display_name="Payroll Sync"
    )
    sp.oauth2_permission_grants.error = GraphError(
        "Authorization_RequestDenied: Insufficient privileges"
    )
    graph.service_principals.items = [sp.value]


async def application_details(graph, ctx):
    graph.applications.by_application_id("payroll-sync-object").value = Application(
        id="payroll-sync-object",
        app_id=PAYROLL_SYNC_APP_ID,
        display_name="Payroll Sync",
    )
    return await server.get_application_by_id("payroll-sync-object", ctx)


async def service_principal_details(graph, ctx):
    return await server.get_service_principal_by_id(PAYROLL_SYNC_SP, ctx)


@pytest.mark.parametrize(
    "ask_for_details",
    [
        pytest.param(application_details, id="application"),
        pytest.param(service_principal_details, id="service principal"),
    ],
)
async def test_details_still_come_back_when_the_permission_grants_cannot_be_read(
    graph, ctx, caplog, ask_for_details
):
    caplog.set_level(logging.DEBUG)
    grants_cannot_be_read(graph)

    details = await ask_for_details(graph, ctx)

    assert details["displayName"] == "Payroll Sync"
    assert details["oauth2PermissionGrants"] == []
    warnings = [r.getMessage() for r in caplog.records if r.levelno == logging.WARNING]
    assert warnings == [
        (
            f"Error fetching oauth2PermissionGrants for service principal {PAYROLL_SYNC_SP}: "
            "Authorization_RequestDenied: Insufficient privileges"
        )
    ]
