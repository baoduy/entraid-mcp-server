"""create_group past its duplicate-name check (DRK-2425 §3 row 2, §8 coverage unit)."""

import pytest

from msgraph_mcp_server.resources import groups


@pytest.fixture
def created(graph, monkeypatch):
    """Graph accepts every new group as ``new-group``; owner and member additions are recorded."""
    record = {"posted": [], "owners": [], "members": []}

    async def post(group):
        record["posted"].append(group)
        group.id = "new-group"
        return group

    async def add_owner(graph_client, group_id, owner_id):
        record["owners"].append((group_id, owner_id))

    async def add_member(graph_client, group_id, member_id):
        record["members"].append((group_id, member_id))

    monkeypatch.setattr(graph.groups, "post", post, raising=False)
    monkeypatch.setattr(groups, "add_group_owner", add_owner)
    monkeypatch.setattr(groups, "add_group_member", add_member)
    return record


async def test_a_new_group_is_created_with_its_owners_and_members(graph, created):
    result = await groups.create_group(
        graph,
        {
            "displayName": "O'Brien Team",
            "mailNickname": "obrienteam",
            "description": "Team of O'Brien",
            "groupTypes": ["Unified"],
            "mailEnabled": True,
            "securityEnabled": False,
            "visibility": "Private",
            "owners": ["owner-1"],
            "members": ["user-1", "user-2"],
        },
    )

    assert result == {
        "id": "new-group",
        "displayName": "O'Brien Team",
        "mail": None,
        "mailNickname": "obrienteam",
        "description": "Team of O'Brien",
        "groupTypes": ["Unified"],
        "securityEnabled": False,
        "mailEnabled": True,
        "visibility": "Private",
        "createdDateTime": None,
    }
    assert created["owners"] == [("new-group", "owner-1")]
    assert created["members"] == [("new-group", "user-1"), ("new-group", "user-2")]


async def test_a_dynamic_group_keeps_its_rule_and_gets_no_members_added(graph, created):
    result = await groups.create_group(
        graph,
        {
            "displayName": "Finance staff",
            "mailNickname": "financestaff",
            "groupTypes": ["DynamicMembership"],
            "membershipRule": 'user.department -eq "Finance"',
            "members": ["user-1"],
        },
    )

    assert result["membershipRule"] == 'user.department -eq "Finance"'
    assert result["membershipRuleProcessingState"] == "On"
    assert created["members"] == []


@pytest.mark.parametrize(
    "group_data, message",
    [
        ({"mailNickname": "team"}, "displayName is required for creating a group"),
        ({"displayName": "Team"}, "mailNickname is required for creating a group"),
        (
            {"displayName": "Team", "mailNickname": "team", "groupTypes": ["DynamicMembership"]},
            "membershipRule is required for dynamic membership groups",
        ),
        ({"displayName": "Team", "mailNickname": "team", "owners": "owner-1"}, "owners must be a list of user IDs"),
        ({"displayName": "Team", "mailNickname": "team", "members": "user-1"}, "members must be a list of user IDs"),
    ],
    ids=["no display name", "no mail nickname", "dynamic without rule", "owners not a list", "members not a list"],
)
async def test_invalid_group_data_is_refused_before_anything_is_created(graph, created, group_data, message):
    with pytest.raises(ValueError, match=f"^{message}$"):
        await groups.create_group(graph, group_data)

    assert created["posted"] == []


async def test_no_group_back_from_microsoft_graph_is_an_error(graph, monkeypatch):
    async def post(group):
        return None

    monkeypatch.setattr(graph.groups, "post", post, raising=False)

    with pytest.raises(Exception, match="^Failed to create group$"):
        await groups.create_group(graph, {"displayName": "Team", "mailNickname": "team"})
