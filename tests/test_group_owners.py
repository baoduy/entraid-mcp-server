"""Owner add posts a reference body to ``/groups/{id}/owners/$ref`` (DRK-2435, DRK-2444 brief §3 rows 1–2).

The fake Graph refuses an owner ``$ref`` body that is not a ``ReferenceCreate`` pointing at
``directoryObjects/{id}`` (``conftest.Owners``), as real Graph does.
"""

from conftest import GRAPH_ROOT

from msgraph_mcp_server.resources import groups


async def test_adding_an_owner_posts_a_reference_to_the_owner(graph):
    added = await groups.add_group_owner(graph, "all-staff", "user-0001")

    assert added is True
    [post] = graph.requests_to("groups/all-staff/owners/$ref")
    assert post.method == "post"
    assert (
        post.body.odata_id
        == "https://graph.microsoft.com/v1.0/directoryObjects/user-0001"
    )


async def test_a_new_group_posts_a_reference_for_each_listed_owner(graph):
    result = await groups.create_group(
        graph,
        {
            "displayName": "All Staff",
            "mailNickname": "allstaff",
            "owners": ["user-0001", "user-0002"],
        },
    )

    assert result["id"] == "new-0001"
    posts = graph.requests_to("groups/new-0001/owners/$ref")
    assert [(p.method, p.body.odata_id) for p in posts] == [
        ("post", f"{GRAPH_ROOT}/directoryObjects/user-0001"),
        ("post", f"{GRAPH_ROOT}/directoryObjects/user-0002"),
    ]


async def test_a_new_group_with_no_owners_posts_no_owner_reference(graph):
    await groups.create_group(
        graph, {"displayName": "All Staff", "mailNickname": "allstaff", "owners": []}
    )

    assert graph.requests_to("groups/new-0001/owners/$ref") == []
