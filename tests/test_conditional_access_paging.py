"""The conditional access policy list reads every page Microsoft Graph serves (DRK-2451 item 4)."""

from types import SimpleNamespace

import pytest
from conftest import GraphError
from msgraph.generated.models.conditional_access_condition_set import (
    ConditionalAccessConditionSet,
)
from msgraph.generated.models.conditional_access_policy import ConditionalAccessPolicy
from test_empty_pages import Pages

from msgraph_mcp_server import server
from msgraph_mcp_server.resources import conditional_access

POLICIES = "identity/conditionalAccess/policies"


def policies(*ids):
    return [
        ConditionalAccessPolicy(
            id=i, display_name=f"Policy {i}", conditions=ConditionalAccessConditionSet()
        )
        for i in ids
    ]


def tenant_with_two_pages_of_policies(graph):
    collection = graph.identity.conditional_access.policies
    collection.items = policies("policy-1", "policy-2", "policy-3")
    collection.page_size = 2


async def test_the_policy_list_returns_the_policies_of_every_page(graph):
    tenant_with_two_pages_of_policies(graph)

    result = await conditional_access.get_conditional_access_policies(graph)

    assert [p["id"] for p in result] == ["policy-1", "policy-2", "policy-3"]
    assert [r.method for r in graph.requests_to(POLICIES)] == ["get", "with_url"]


async def test_the_policy_list_tool_returns_the_policies_of_every_page(graph, ctx):
    tenant_with_two_pages_of_policies(graph)

    result = await server.get_conditional_access_policies(ctx)

    assert [p["id"] for p in result] == ["policy-1", "policy-2", "policy-3"]
    assert [r.method for r in graph.requests_to(POLICIES)] == ["get", "with_url"]


async def test_a_single_page_of_policies_is_read_once(graph):
    graph.identity.conditional_access.policies.items = policies("policy-1", "policy-2")

    result = await conditional_access.get_conditional_access_policies(graph)

    assert [p["id"] for p in result] == ["policy-1", "policy-2"]
    assert [r.method for r in graph.requests_to(POLICIES)] == ["get"]


async def test_the_policy_list_reads_on_past_an_empty_page(graph):
    first, second, third = policies("policy-1", "policy-2", "policy-3")
    graph.identity.conditional_access.policies = Pages([first], [], [second, third])

    result = await conditional_access.get_conditional_access_policies(graph)

    assert [p["id"] for p in result] == ["policy-1", "policy-2", "policy-3"]


async def test_a_tenant_with_no_policies_returns_an_empty_list(graph):
    result = await conditional_access.get_conditional_access_policies(graph)

    assert result == []
    assert [r.method for r in graph.requests_to(POLICIES)] == ["get"]


async def test_a_policy_list_with_no_body_from_microsoft_graph_is_empty(graph):
    graph.identity.conditional_access.policies.respond_none = True

    assert await conditional_access.get_conditional_access_policies(graph) == []


async def test_a_failed_next_page_read_reaches_the_ai_client(graph, ctx):
    tenant_with_two_pages_of_policies(graph)
    refusal = GraphError("Authorization_RequestDenied: no access")

    async def refuse(request_configuration=None):
        raise refusal

    graph.identity.conditional_access.policies.with_url = lambda url: SimpleNamespace(
        get=refuse
    )

    with pytest.raises(GraphError) as raised:
        await server.get_conditional_access_policies(ctx)

    assert raised.value is refusal
    assert ctx.errors == [
        "Error fetching conditional access policies: Authorization_RequestDenied: no access"
    ]
