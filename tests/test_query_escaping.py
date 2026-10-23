"""User text in Microsoft Graph filters and searches (DRK-2421 §5, DRK-2425 §7 and §6a D1–D2)."""

import pytest
from conftest import make_groups
from msgraph.generated.models.application import Application
from msgraph.generated.models.service_principal import ServicePrincipal

from msgraph_mcp_server import server


def only_filter(graph, path):
    [request] = [r for r in graph.requests_to(path) if r.method == "get"]
    return request.filter


async def test_group_search_by_name_keeps_a_quote_in_the_value(graph, ctx):
    await server.search_groups_by_name("O'Brien Team", ctx)

    assert only_filter(graph, "groups") == "startswith(displayName,'O''Brien Team')"


async def test_the_duplicate_name_check_when_creating_a_group_keeps_a_quote_in_the_value(
    graph, ctx
):
    graph.groups.items = make_groups(1)

    await server.create_group(
        ctx, {"displayName": "O'Brien Team", "mailNickname": "obrienteam"}
    )

    assert only_filter(graph, "groups") == "displayName eq 'O''Brien Team'"


async def test_devices_by_operating_system_keep_a_quote_in_the_value(graph, ctx):
    await server.get_all_managed_devices(ctx, filter_os="Windows'11")

    assert (
        only_filter(graph, "deviceManagement/managedDevices")
        == "operatingSystem eq 'Windows''11'"
    )


async def test_devices_by_user_keep_a_quote_in_the_value(graph, ctx):
    await server.get_managed_devices_by_user("7f3c'9a", ctx)

    assert (
        only_filter(graph, "deviceManagement/managedDevices") == "userId eq '7f3c''9a'"
    )


async def test_audit_logs_by_user_keep_a_quote_in_the_value(graph, ctx, frozen_now):
    await server.get_user_audit_logs("7f3c'9a", ctx)

    assert only_filter(graph, "auditLogs/directoryAudits") == (
        "initiatedBy/user/id eq '7f3c''9a'"
        " and activityDateTime ge 2026-09-23T08:00:00Z"
        " and activityDateTime le 2026-10-23T08:00:00Z"
    )


async def test_sign_in_logs_by_user_keep_a_quote_in_the_value(graph, ctx, frozen_now):
    await server.get_user_sign_ins("7f3c'9a", ctx)

    assert only_filter(graph, "auditLogs/signIns") == (
        "createdDateTime ge 2026-10-16T08:00:00Z"
        " and createdDateTime le 2026-10-23T08:00:00Z"
        " and userId eq '7f3c''9a'"
    )


async def test_the_service_principal_lookup_in_application_details_keeps_a_quote_in_the_value(
    graph, ctx
):
    graph.applications.by_application_id("app-object-1").value = Application(
        id="app-object-1", app_id="5d2e'41", display_name="Payroll Sync"
    )
    graph.service_principals.items = [ServicePrincipal(id="sp-1", app_id="5d2e'41")]

    await server.get_application_by_id("app-object-1", ctx)

    assert only_filter(graph, "servicePrincipals") == "appId eq '5d2e''41'"


async def test_crafted_text_cannot_add_a_condition_to_a_group_search(graph, ctx):
    await server.search_groups_by_name("Finance' or 1 eq 1 or 'a' eq 'a", ctx)

    assert (
        only_filter(graph, "groups")
        == "startswith(displayName,'Finance'' or 1 eq 1 or ''a'' eq ''a')"
    )


@pytest.mark.parametrize(
    "text, escaped",
    [
        ('Bob "The Builder"', 'Bob \\"The Builder\\"'),
        ("CONTOSO\\aoife", "CONTOSO\\\\aoife"),
    ],
    ids=["double quotes", "backslash"],
)
async def test_user_search_keeps_quotes_and_backslashes_as_typed(
    graph, ctx, text, escaped
):
    await server.search_users(text, ctx)

    [request] = graph.requests_to("users")
    assert request.search == [
        (
            f'("displayName:{escaped}" OR "mail:{escaped}" OR "userPrincipalName:{escaped}"'
            f' OR "givenName:{escaped}" OR "surName:{escaped}" OR "otherMails:{escaped}")'
        )
    ]


# -- the helpers every filter and search goes through (DRK-2425 §3 row 1) ---------------


@pytest.mark.parametrize(
    "value, literal",
    [
        ("", "''"),
        ("Finance", "'Finance'"),
        ("O'Brien", "'O''Brien'"),
        ("a'b'c", "'a''b''c'"),
        ("'x'", "'''x'''"),
        ("Finance' or 1 eq 1 or 'a' eq 'a", "'Finance'' or 1 eq 1 or ''a'' eq ''a'"),
    ],
    ids=[
        "empty",
        "no quote",
        "one quote",
        "several quotes",
        "quotes at both ends",
        "crafted condition",
    ],
)
def test_odata_string_literal_wraps_the_value_once_and_doubles_each_quote(
    value, literal
):
    from msgraph_mcp_server.utils.odata import odata_string_literal

    assert odata_string_literal(value) == literal


@pytest.mark.parametrize(
    "value, escaped",
    [
        ('"', '\\"'),
        ("\\", "\\\\"),
        ('a\\"b', 'a\\\\\\"b'),
        ("Aoife Byrne", "Aoife Byrne"),
    ],
    ids=["double quote", "backslash", "both", "none"],
)
def test_graph_search_escape_prefixes_each_double_quote_and_backslash(value, escaped):
    from msgraph_mcp_server.utils.odata import graph_search_escape

    assert graph_search_escape(value) == escaped
