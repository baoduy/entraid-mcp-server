"""Paging edges the acceptance tests leave open (DRK-2425 §3 rows 5 and 6): empty pages, no body, a limit of 1."""

from types import SimpleNamespace

import pytest
from conftest import make_groups
from test_application_details_and_sign_ins import sign_ins
from test_paging import TOOLS

from msgraph_mcp_server import server


class Pages:
    """A collection that serves the given pages in order, each linking to the next."""

    def __init__(self, *pages):
        self._pages = [
            SimpleNamespace(
                value=list(items),
                odata_next_link=f"page-{i + 1}" if i + 1 < len(pages) else None,
            )
            for i, items in enumerate(pages)
        ]

    async def get(self, request_configuration=None):
        return self._pages[0]

    def with_url(self, url):
        page = self._pages[int(url.removeprefix("page-"))]

        async def get(request_configuration=None):
            return page

        return SimpleNamespace(get=get)


async def test_a_group_list_reads_on_past_an_empty_page(graph, ctx):
    first, second, third = make_groups(3)
    graph.groups = Pages([first], [], [second, third])

    result = await server.get_all_groups(ctx, limit=5)

    assert [g["id"] for g in result] == ["group-0000", "group-0001", "group-0002"]


async def test_sign_in_logs_read_on_past_an_empty_page(graph, ctx):
    first, second, third = sign_ins(3)
    graph.audit_logs.sign_ins = Pages([first], [], [second, third])

    logs = await server.get_user_sign_ins("aoife@contoso.com", ctx, days=7)

    assert [log["id"] for log in logs] == [
        "sign-in-0000",
        "sign-in-0001",
        "sign-in-0002",
    ]


async def test_a_group_list_with_no_body_from_microsoft_graph_is_empty(graph, ctx):
    graph.groups.respond_none = True

    assert await server.get_all_groups(ctx, limit=5) == []


async def test_sign_in_logs_with_no_body_from_microsoft_graph_are_empty(graph, ctx):
    graph.audit_logs.sign_ins.respond_none = True

    assert await server.get_user_sign_ins("aoife@contoso.com", ctx, days=7) == []


@pytest.mark.parametrize("tool, tenant_has", TOOLS)
async def test_a_limit_of_1_returns_one_item(graph, ctx, tool, tenant_has):
    tenant_has(graph, 300)

    result = await tool(ctx, 1)

    assert len(result) == 1


async def test_a_policy_list_stops_at_a_next_page_with_no_body(graph, ctx):
    # Imported here: that module imports ``Pages`` from this one.
    from test_conditional_access_paging import policies as make_policies

    first, second = make_policies("policy-1", "policy-2")
    policies = graph.identity.conditional_access.policies
    policies.items = [first, second]
    policies.page_size = 1

    async def no_body(request_configuration=None):
        return None

    policies.with_url = lambda url: SimpleNamespace(get=no_body)

    result = await server.get_conditional_access_policies(ctx)

    assert [p["id"] for p in result] == ["policy-1"]
