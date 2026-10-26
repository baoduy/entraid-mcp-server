"""Permission search over Graph permissions with a null searched field (DRK-2462).

Microsoft Graph can return a permission scope or an app role without a value, display name
or description. The search counts that field as no match, still searches the row's other
fields, and returns a matching row with its null field intact.
"""

import pytest
from msgraph.generated.models.app_role import AppRole
from msgraph.generated.models.permission_scope import PermissionScope
from test_result_shapes import microsoft_graph_permissions
from test_schema_graph_sp_mfa_mail import USER_READ, USER_READ_ALL

from msgraph_mcp_server import server

# Every searched field of these rows holds "user.read", so each still matches with one null.
USER_READ_BASIC_SCOPE = {
    "value": "User.ReadBasic.All",
    "admin_consent_display_name": "User.Read for basic profiles",
    "admin_consent_description": "Allows User.Read of basic profiles.",
    "is_enabled": True,
}
USER_READ_BASIC = {
    "id": None,
    "value": "User.ReadBasic.All",
    "type": "delegated",
    "adminConsentDisplayName": "User.Read for basic profiles",
    "adminConsentDescription": "Allows User.Read of basic profiles.",
    "userConsentDisplayName": None,
    "userConsentDescription": None,
    "isEnabled": True,
}
USER_READ_BASIC_ROLE = {
    "value": "User.ReadBasic.All",
    "display_name": "User.Read for all basic profiles",
    "description": "Allows User.Read of all basic profiles.",
    "is_enabled": True,
}
USER_READ_BASIC_ALL = {
    "id": None,
    "value": "User.ReadBasic.All",
    "type": "application",
    "displayName": "User.Read for all basic profiles",
    "description": "Allows User.Read of all basic profiles.",
    "isEnabled": True,
}
# No searched field of these rows holds "user.read", so the search reads every one of them.
MAIL_SEND_SCOPE = {
    "value": "Mail.Send",
    "admin_consent_display_name": "Send mail as a user",
    "admin_consent_description": "Allows the app to send mail as users.",
    "is_enabled": True,
}
MAIL_SEND_ROLE = {
    "value": "Mail.Send",
    "display_name": "Send mail as any user",
    "description": "Allows the app to send mail as any user.",
    "is_enabled": True,
}


def search_result(term, delegated, application):
    return {
        "status": "success",
        "search_term": term,
        "matching_delegated_permissions": delegated,
        "matching_application_permissions": application,
        "total_matches": len(delegated) + len(application),
    }


@pytest.mark.parametrize(
    ("field", "attribute"),
    [
        ("value", "value"),
        ("adminConsentDisplayName", "admin_consent_display_name"),
        ("adminConsentDescription", "admin_consent_description"),
    ],
)
async def test_delegated_search_matches_a_row_on_the_fields_that_are_not_null(
    graph, ctx, field, attribute
):
    sp = microsoft_graph_permissions(graph)
    sp.oauth2_permission_scopes += [
        PermissionScope(**{**USER_READ_BASIC_SCOPE, attribute: None}),
        PermissionScope(**{**MAIL_SEND_SCOPE, attribute: None}),
    ]

    result = await server.search_permissions("User.Read", ctx, "delegated")

    assert result == search_result(
        "user.read", [USER_READ, {**USER_READ_BASIC, field: None}], []
    )


@pytest.mark.parametrize(
    ("field", "attribute"),
    [
        ("value", "value"),
        ("displayName", "display_name"),
        ("description", "description"),
    ],
)
async def test_application_search_matches_a_row_on_the_fields_that_are_not_null(
    graph, ctx, field, attribute
):
    sp = microsoft_graph_permissions(graph)
    sp.app_roles += [
        AppRole(**{**USER_READ_BASIC_ROLE, attribute: None}),
        AppRole(**{**MAIL_SEND_ROLE, attribute: None}),
    ]

    result = await server.search_permissions("User.Read", ctx, "application")

    assert result == search_result(
        "user.read", [], [USER_READ_ALL, {**USER_READ_BASIC_ALL, field: None}]
    )


@pytest.mark.parametrize(
    ("term", "search_term", "delegated", "application"),
    [
        ("User.Read", "user.read", [USER_READ], [USER_READ_ALL]),
        ("None", "none", [], []),
    ],
)
async def test_search_never_matches_a_row_whose_searched_fields_are_all_null(
    graph, ctx, term, search_term, delegated, application
):
    sp = microsoft_graph_permissions(graph)
    sp.oauth2_permission_scopes.append(PermissionScope(is_enabled=True))
    sp.app_roles.append(AppRole(is_enabled=True))

    result = await server.search_permissions(term, ctx)

    assert result == search_result(search_term, delegated, application)


@pytest.mark.parametrize(
    ("permission_type", "delegated", "application"),
    [
        (
            None,
            [USER_READ, {**USER_READ_BASIC, "value": None}],
            [USER_READ_ALL, {**USER_READ_BASIC_ALL, "value": None}],
        ),
        ("delegated", [USER_READ, {**USER_READ_BASIC, "value": None}], []),
        ("application", [], [USER_READ_ALL, {**USER_READ_BASIC_ALL, "value": None}]),
    ],
)
async def test_search_with_null_rows_present_searches_only_the_requested_lists(
    graph, ctx, permission_type, delegated, application
):
    sp = microsoft_graph_permissions(graph)
    sp.oauth2_permission_scopes.append(
        PermissionScope(**{**USER_READ_BASIC_SCOPE, "value": None})
    )
    sp.app_roles.append(AppRole(**{**USER_READ_BASIC_ROLE, "value": None}))

    result = await server.search_permissions("User.Read", ctx, permission_type)

    assert result == search_result("user.read", delegated, application)
