"""User search schema, Microsoft Graph's service principal by app id, and the MFA mail (DRK-2451 items 1-3).

User search runs over the MCP protocol with fastmcp's in-memory ``Client``: the client checks
each structured result against the output schema the server lists, as an AI client does.
"""

import pytest
from conftest import MICROSOFT_GRAPH_APP_ID
from fastmcp import Client
from msgraph.generated.models.user import User
from test_result_shapes import AOIFE, microsoft_graph_permissions

from msgraph_mcp_server import server

USER_READ = {
    "id": None,
    "value": "User.Read",
    "type": "delegated",
    "adminConsentDisplayName": "Sign in and read user profile",
    "adminConsentDescription": "Allows users to sign in and read their profile.",
    "userConsentDisplayName": "Sign you in and read your profile",
    "userConsentDescription": "Allows you to sign in and read your profile.",
    "isEnabled": True,
}
USER_READ_ALL = {
    "id": None,
    "value": "User.Read.All",
    "type": "application",
    "displayName": "Read all users' full profiles",
    "description": "Allows the app to read user profiles.",
    "isEnabled": True,
}


async def search_users_as_a_client(arguments):
    async with Client(server.mcp) as client:
        return await client.call_tool("search_users", arguments)


# -- search_users results pass its published output schema -----------------------------------


async def test_user_search_returns_null_fields_and_phone_lists_to_a_client(graph):
    graph.users.items = [
        User(
            id="niamh-0001",
            display_name="Niamh Kelly",
            mail=None,
            user_principal_name="niamh@contoso.com",
            given_name="Niamh",
            surname="Kelly",
            job_title=None,
            office_location="Cork",
            business_phones=["+353 21 555 0100"],
            mobile_phone=None,
        ),
        User(
            id="cian-0002",
            display_name="Cian Walsh",
            mail="cian@contoso.com",
            user_principal_name="cian@contoso.com",
            given_name="Cian",
            surname="Walsh",
            job_title="Engineer",
            office_location="Dublin",
            business_phones=[],
            mobile_phone="+353 87 555 0101",
        ),
    ]

    result = await search_users_as_a_client({"query": "contoso"})

    assert result.structured_content == {
        "result": [
            {
                "id": "niamh-0001",
                "displayName": "Niamh Kelly",
                "mail": None,
                "userPrincipalName": "niamh@contoso.com",
                "givenName": "Niamh",
                "surname": "Kelly",
                "jobTitle": None,
                "officeLocation": "Cork",
                "businessPhones": ["+353 21 555 0100"],
                "mobilePhone": None,
            },
            {
                "id": "cian-0002",
                "displayName": "Cian Walsh",
                "mail": "cian@contoso.com",
                "userPrincipalName": "cian@contoso.com",
                "givenName": "Cian",
                "surname": "Walsh",
                "jobTitle": "Engineer",
                "officeLocation": "Dublin",
                "businessPhones": [],
                "mobilePhone": "+353 87 555 0101",
            },
        ]
    }


async def test_a_user_search_with_no_match_returns_an_empty_list_to_a_client(graph):
    result = await search_users_as_a_client({"query": "nobody"})

    assert result.structured_content == {"result": []}


# -- Graph permissions found by app id -------------------------------------------------------


async def test_graph_permissions_are_found_by_app_id(graph, ctx):
    microsoft_graph_permissions(graph)

    result = await server.get_all_graph_permissions(ctx)

    assert result == {
        "status": "success",
        "delegated_permissions": [USER_READ],
        "application_permissions": [USER_READ_ALL],
    }
    assert [(r.method, r.filter) for r in graph.requests_to("servicePrincipals")] == [
        ("get", "appId eq '00000003-0000-0000-c000-000000000000'")
    ]
    assert graph.requests_under(f"servicePrincipals/{MICROSOFT_GRAPH_APP_ID}") == []


@pytest.mark.parametrize(
    ("permission_type", "delegated", "application"),
    [
        (None, [USER_READ], [USER_READ_ALL]),
        ("delegated", [USER_READ], []),
        ("application", [], [USER_READ_ALL]),
    ],
)
async def test_permission_search_finds_graph_permissions_by_app_id(
    graph, ctx, permission_type, delegated, application
):
    microsoft_graph_permissions(graph)

    result = await server.search_permissions("User.Read", ctx, permission_type)

    assert result == {
        "status": "success",
        "search_term": "user.read",
        "matching_delegated_permissions": delegated,
        "matching_application_permissions": application,
        "total_matches": len(delegated) + len(application),
    }
    assert [(r.method, r.filter) for r in graph.requests_to("servicePrincipals")] == [
        ("get", "appId eq '00000003-0000-0000-c000-000000000000'")
    ]
    assert graph.requests_under(f"servicePrincipals/{MICROSOFT_GRAPH_APP_ID}") == []


async def test_graph_permissions_report_a_missing_graph_service_principal(graph, ctx):
    result = await server.get_all_graph_permissions(ctx)

    assert result == {
        "status": "error",
        "message": "Microsoft Graph service principal not found",
    }


# -- MFA result carries the user's mail ------------------------------------------------------


async def test_the_user_mfa_status_carries_the_users_mail(graph, ctx):
    graph.add_user(AOIFE)

    result = await server.get_user_mfa_status(AOIFE, ctx)

    assert result["mail"] == "aoife@contoso.com"
    assert result["companyName"] == "Contoso"
    profile_reads = [r for r in graph.requests_to(f"users/{AOIFE}") if r.select]
    assert [sorted(r.select) for r in profile_reads] == [["companyName", "mail"]]


async def test_the_group_mfa_report_carries_each_members_mail(graph, ctx):
    graph.add_user(AOIFE)
    graph.groups.by_group_id("all-staff").members.items = [User(id=AOIFE)]

    report = await server.get_group_mfa_status("all-staff", ctx)

    assert [(m["userPrincipalName"], m["mail"]) for m in report] == [
        ("aoife@contoso.com", "aoife@contoso.com")
    ]


async def test_the_user_mfa_status_of_a_user_without_mail_has_a_null_mail(graph, ctx):
    graph.add_user(AOIFE).mail = None

    result = await server.get_user_mfa_status(AOIFE, ctx)

    assert result["mail"] is None
    assert result["companyName"] == "Contoso"
