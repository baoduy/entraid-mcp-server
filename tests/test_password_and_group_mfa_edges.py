"""Edges of the password reset and group MFA tools the acceptance tests do not reach (DRK-2424 §3 rows 2, 4 and 5)."""

import pytest
from conftest import GraphError
from msgraph.generated.models.device import Device
from msgraph.generated.models.user import User

from msgraph_mcp_server import server

AOIFE = "aoife@contoso.com"
ALL_STAFF = "all-staff"


async def test_a_reset_with_no_password_and_no_generation_is_refused_before_any_change(
    graph, ctx
):
    graph.add_user(AOIFE)

    with pytest.raises(ValueError) as refusal:
        await server.reset_user_password_direct(AOIFE, ctx)

    assert (
        str(refusal.value)
        == "Password must be provided or generate_password must be set to True"
    )
    assert ctx.errors == [
        (
            f"Error directly resetting password for user {AOIFE}: "
            "Password must be provided or generate_password must be set to True"
        )
    ]
    assert graph.password_changes(AOIFE) == []


@pytest.mark.parametrize(
    "arrange",
    [
        pytest.param(lambda members: None, id="no members"),
        pytest.param(
            lambda members: setattr(members, "respond_none", True),
            id="Graph answers with no body",
        ),
        pytest.param(
            lambda members: setattr(
                members, "items", [Device(id="laptop-0042", display_name="LAPTOP-0042")]
            ),
            id="only a device",
        ),
    ],
)
async def test_a_group_with_no_user_members_gets_only_the_empty_report_warning(
    graph, ctx, caplog, arrange
):
    members = graph.groups.by_group_id(ALL_STAFF).members
    arrange(members)

    report = await server.get_group_mfa_status(ALL_STAFF, ctx)

    assert report == []
    assert ctx.warnings == [f"No MFA data found for group {ALL_STAFF}"]
    assert graph.requests_under("users") == []
    assert [r.getMessage() for r in caplog.records if r.levelname == "WARNING"] == [
        f"No user members found in group {ALL_STAFF}"
    ]


async def test_a_refused_member_list_reaches_the_client_as_an_error(
    graph, ctx, caplog, monkeypatch
):
    members = graph.groups.by_group_id(ALL_STAFF).members

    async def refuse(request_configuration=None):
        raise GraphError("Authorization_RequestDenied: Insufficient privileges")

    monkeypatch.setattr(members, "get", refuse)

    with pytest.raises(
        GraphError, match="^Authorization_RequestDenied: Insufficient privileges$"
    ):
        await server.get_group_mfa_status(ALL_STAFF, ctx)

    assert ctx.errors == [
        f"Error fetching group MFA status for {ALL_STAFF}: Authorization_RequestDenied: Insufficient privileges"
    ]
    assert (
        f"Error fetching group MFA status for group {ALL_STAFF}: Authorization_RequestDenied: Insufficient privileges"
        in [r.getMessage() for r in caplog.records]
    )


async def test_an_unreadable_member_is_logged_with_the_graph_error(graph, ctx, caplog):
    graph.add_user(AOIFE, readable=False)
    graph.groups.by_group_id(ALL_STAFF).members.items = [User(id=AOIFE)]

    await server.get_group_mfa_status(ALL_STAFF, ctx)

    assert (
        f"Error processing member {AOIFE}: Authorization_RequestDenied: cannot read methods of {AOIFE}"
        in [r.getMessage() for r in caplog.records]
    )


async def test_a_refused_reset_logs_the_error_type_only(graph, ctx, caplog):
    graph.add_user(AOIFE)
    graph.users.by_user_id(AOIFE).refuse_patch = True

    with pytest.raises(RuntimeError):
        await server.reset_user_password_direct(
            AOIFE, ctx, password="Sp!ce-Harbor-2026"
        )

    assert f"Error directly resetting password for user {AOIFE}: GraphError" in [
        r.getMessage() for r in caplog.records
    ]
