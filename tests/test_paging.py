"""List limits in the group list, group search and group member list (DRK-2421 §5, DRK-2425 §7 and §6a D3)."""

import pytest
from conftest import make_groups, make_users

from msgraph_mcp_server import server

ALL_STAFF = "all-staff"


async def group_list(ctx, limit):
    return await server.get_all_groups(ctx, limit=limit)


async def group_search(ctx, limit):
    return await server.search_groups_by_name("Fin", ctx, limit=limit)


async def group_member_list(ctx, limit):
    return await server.get_group_members(ALL_STAFF, ctx, limit=limit)


def groups_named_fin(graph, count):
    graph.groups.items = make_groups(count, prefix="Fin")
    return "groups"


def groups(graph, count):
    graph.groups.items = make_groups(count)
    return "groups"


def members_of_all_staff(graph, count):
    graph.groups.by_group_id(ALL_STAFF).members.items = make_users(count)
    return f"groups/{ALL_STAFF}/members"


TOOLS = [
    pytest.param(group_list, groups, id="group list"),
    pytest.param(group_search, groups_named_fin, id='group search for "Fin"'),
    pytest.param(
        group_member_list, members_of_all_staff, id='group member list for "All Staff"'
    ),
]


@pytest.mark.parametrize("tool, tenant_has", TOOLS)
async def test_paging_stops_once_the_limit_is_reached(graph, ctx, tool, tenant_has):
    path = tenant_has(graph, 300)

    result = await tool(ctx, 5)

    assert len(result) == 5
    assert len(graph.requests_to(path)) == 1


@pytest.mark.parametrize("tool, tenant_has", TOOLS)
async def test_a_limit_above_999_still_returns_up_to_the_limit(
    graph, ctx, tool, tenant_has
):
    path = tenant_has(graph, 2000)

    result = await tool(ctx, 1500)

    assert len(result) == 1500
    [first] = [r for r in graph.requests_to(path) if r.method == "get"]
    assert first.top == 999


@pytest.mark.parametrize("limit", [0, -1])
@pytest.mark.parametrize("tool, tenant_has", TOOLS)
async def test_a_limit_of_0_or_less_returns_nothing_without_asking_microsoft_graph(
    graph, ctx, tool, tenant_has, limit
):
    tenant_has(graph, 300)

    result = await tool(ctx, limit)

    assert result == []
    assert graph.requests == []
    assert graph.get_client_calls == 0


@pytest.mark.parametrize("limit", [999, 1000])
@pytest.mark.parametrize("tool, tenant_has", TOOLS)
async def test_the_page_size_asked_of_microsoft_graph_is_never_above_999(
    graph, ctx, tool, tenant_has, limit
):
    path = tenant_has(graph, 1200)

    await tool(ctx, limit)

    [first] = [r for r in graph.requests_to(path) if r.method == "get"]
    assert first.top == 999
