"""The group MFA report through the `get_group_mfa_status` tool (DRK-2421 §5, DRK-2424 §7 and §6a D2–D5, DRK-2420 findings 1 and 3)."""

import pytest
from msgraph.generated.models.device import Device
from msgraph.generated.models.directory_object import DirectoryObject
from msgraph.generated.models.group import Group
from msgraph.generated.models.service_principal import ServicePrincipal
from msgraph.generated.models.user import User

from msgraph_mcp_server import server

ALL_STAFF = "all-staff"
COUNT_WARNING_START = "MFA status could not be read for"


def all_staff_has(graph, members):
    graph.groups.by_group_id(ALL_STAFF).members.items = members


def user_members(graph, count, prefix="staff"):
    """User members whose MFA status can be read."""
    return [User(id=graph.add_user(f"{prefix}{i:03d}@contoso.com").id) for i in range(count)]


def count_warnings(ctx):
    return [w for w in ctx.warnings if w.startswith(COUNT_WARNING_START)]


def looked_up(graph, member_id):
    return graph.requests_under(f"users/{member_id}")


async def test_the_group_mfa_report_covers_every_page_of_members(graph, ctx):
    all_staff_has(graph, user_members(graph, 250))

    report = await server.get_group_mfa_status(ALL_STAFF, ctx)

    assert len(report) == 250


def contractors(graph):
    nested = Group(id="contractors", display_name="Contractors")
    graph.groups.by_group_id("contractors").members.items = user_members(graph, 4, prefix="contractor")
    return nested


@pytest.mark.parametrize(
    "non_user",
    [
        pytest.param(lambda graph: Device(id="laptop-0042", display_name="LAPTOP-0042"), id="device"),
        pytest.param(lambda graph: ServicePrincipal(id="payroll-sync", display_name="Payroll Sync"), id="service principal"),
        pytest.param(contractors, id="nested group with 4 user members"),
        pytest.param(lambda graph: DirectoryObject(id="unknown-object", odata_type=None), id="no odata type"),
    ],
)
async def test_members_that_are_not_users_are_skipped_without_a_lookup(graph, ctx, non_user):
    member = non_user(graph)
    all_staff_has(graph, user_members(graph, 2) + [member])

    report = await server.get_group_mfa_status(ALL_STAFF, ctx)

    assert len(report) == 2
    assert looked_up(graph, member.id) == []


async def test_unreadable_mfa_statuses_are_reported_in_one_warning(graph, ctx):
    members = user_members(graph, 10)
    graph.users.by_user_id(members[0].id).error = Exception("Authorization_RequestDenied")
    graph.users.by_user_id(members[1].id).authentication.methods.respond_none = True
    all_staff_has(graph, members)

    report = await server.get_group_mfa_status(ALL_STAFF, ctx)

    assert len(report) == 8
    assert count_warnings(ctx) == ["MFA status could not be read for 2 of 10 members"]


async def test_the_count_warning_counts_user_members_only(graph, ctx):
    members = user_members(graph, 10)
    graph.users.by_user_id(members[3].id).error = Exception("Authorization_RequestDenied")
    graph.users.by_user_id(members[7].id).error = Exception("Authorization_RequestDenied")
    all_staff_has(graph, members + [Device(id="laptop-0042", display_name="LAPTOP-0042")])

    report = await server.get_group_mfa_status(ALL_STAFF, ctx)

    assert len(report) == 8
    assert count_warnings(ctx) == ["MFA status could not be read for 2 of 10 members"]


async def test_a_group_whose_statuses_are_all_readable_gets_no_count_warning(graph, ctx):
    all_staff_has(graph, user_members(graph, 10))

    report = await server.get_group_mfa_status(ALL_STAFF, ctx)

    assert len(report) == 10
    assert count_warnings(ctx) == []


async def test_a_group_whose_statuses_are_all_unreadable_gets_the_count_and_the_empty_report_warning(graph, ctx):
    members = user_members(graph, 3)
    for member in members:
        graph.users.by_user_id(member.id).error = Exception("Authorization_RequestDenied")
    all_staff_has(graph, members)

    report = await server.get_group_mfa_status(ALL_STAFF, ctx)

    assert report == []
    assert sorted(ctx.warnings) == [
        "MFA status could not be read for 3 of 3 members",
        f"No MFA data found for group {ALL_STAFF}",
    ]
