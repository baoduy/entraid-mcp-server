"""Add member and remove member do what they report (DRK-2429 §5, DRK-2432 brief §3 rows 10–11).

The fake Graph keeps each group's member list and only the member calls msgraph-sdk offers
(``conftest.Members``), with Graph's own refusals for a duplicate add and a missing member.
"""

import json

import pytest
from fastmcp import Client
from msgraph.generated.models.group import Group
from msgraph.generated.models.user import User

from msgraph_mcp_server import server

AOIFE = "aoife@contoso.com"
ALL_STAFF = "all-staff"
CONTRACTORS = "contractors"


def all_staff(graph, members):
    graph.groups.by_group_id(ALL_STAFF).value = Group(
        id=ALL_STAFF, display_name="All Staff", group_types=["Unified"]
    )
    group_members = graph.groups.by_group_id(ALL_STAFF).members
    group_members.items = [User(id=member_id) for member_id in members]
    return group_members


async def use_tool(name, arguments):
    async with Client(server.mcp) as client:
        result = await client.call_tool(name, arguments, raise_on_error=False)
    [content] = result.content
    # An error result carries plain text, not a result dict.
    return result.is_error, content.text if result.is_error else json.loads(
        content.text
    )


async def test_removing_a_member_removes_them(graph):
    members = all_staff(graph, ["user-0001", AOIFE])
    assert members.ids() == ["user-0001", AOIFE]

    is_error, result = await use_tool(
        "remove_group_member", {"group_id": ALL_STAFF, "member_id": AOIFE}
    )

    assert members.ids() == ["user-0001"]
    assert is_error is False
    assert result == {
        "status": "success",
        "message": f"Member {AOIFE} was removed from group {ALL_STAFF}",
    }


async def test_adding_a_new_member_adds_them(graph):
    members = all_staff(graph, ["user-0001"])
    assert members.ids() == ["user-0001"]

    is_error, result = await use_tool(
        "add_group_member", {"group_id": ALL_STAFF, "member_id": AOIFE}
    )

    assert members.ids() == ["user-0001", AOIFE]
    assert is_error is False
    assert result == {
        "status": "success",
        "message": f"Member {AOIFE} was added to group {ALL_STAFF}",
    }


@pytest.mark.parametrize(
    "members_before, tool, result",
    [
        pytest.param(
            ["user-0001", AOIFE],
            "add_group_member",
            {
                "status": "already_exists",
                "message": f"Member {AOIFE} is already in group {ALL_STAFF}",
            },
            id="is already in · adds aoife@contoso.com to · is already in",
        ),
        pytest.param(
            ["user-0001"],
            "remove_group_member",
            {
                "status": "not_found",
                "message": f"Member {AOIFE} is not in group {ALL_STAFF}",
            },
            id="is not in · removes aoife@contoso.com from · is not in",
        ),
    ],
)
async def test_known_member_cases_come_back_as_normal_results(
    graph, members_before, tool, result
):
    members = all_staff(graph, members_before)

    is_error, outcome = await use_tool(
        tool, {"group_id": ALL_STAFF, "member_id": AOIFE}
    )

    assert outcome == result
    assert is_error is False
    assert members.ids() == members_before


@pytest.mark.parametrize(
    "tool, message",
    [
        pytest.param(
            "add_group_member",
            "Cannot add members to a dynamic membership group. Members are determined by the membership rule.",
            id="adds aoife@contoso.com to",
        ),
        pytest.param(
            "remove_group_member",
            "Cannot remove members from a dynamic membership group. Members are determined by the membership rule.",
            id="removes aoife@contoso.com from",
        ),
    ],
)
async def test_dynamic_membership_groups_keep_their_result(graph, tool, message):
    graph.groups.by_group_id(CONTRACTORS).value = Group(
        id=CONTRACTORS, display_name="Contractors", group_types=["DynamicMembership"]
    )

    is_error, result = await use_tool(
        tool, {"group_id": CONTRACTORS, "member_id": AOIFE}
    )

    assert result == {
        "status": "error",
        "message": message,
        "groupId": CONTRACTORS,
        "memberId": AOIFE,
        "isDynamicGroup": True,
    }
    assert is_error is False
    assert graph.groups.by_group_id(CONTRACTORS).members.ids() == []
