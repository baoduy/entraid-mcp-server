"""A page with no items that still links to the next page (DRK-2425 §3 rows 5 and 6)."""

from types import SimpleNamespace

from conftest import make_groups
from msgraph_mcp_server import server
from test_application_details_and_sign_ins import sign_ins


class Pages:
    """A collection that serves the given pages in order, each linking to the next."""

    def __init__(self, *pages):
        self._pages = [
            SimpleNamespace(value=list(items), odata_next_link=f"page-{i + 1}" if i + 1 < len(pages) else None)
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
    graph.groups = Pages([], make_groups(3))

    result = await server.get_all_groups(ctx, limit=5)

    assert [g["id"] for g in result] == ["group-0000", "group-0001", "group-0002"]


async def test_sign_in_logs_read_on_past_an_empty_page(graph, ctx):
    graph.audit_logs.sign_ins = Pages([], sign_ins(3))

    logs = await server.get_user_sign_ins("aoife@contoso.com", ctx, days=7)

    assert [log["id"] for log in logs] == ["sign-in-0000", "sign-in-0001", "sign-in-0002"]
