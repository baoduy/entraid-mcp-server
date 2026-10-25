# Groups

Manage groups: find and read them, create and change them, and control who is a member or an owner.

## 📖 Overview

- Find groups by the start of their display name, list them all, or read one, with paging up to the limit you ask for.
- Create a group with its owners and members in one call. A group whose display name already exists is returned as it is instead of a second one being created; the check is a read just before the create, so two calls at the same moment can still both create one.
- Add or remove a member: someone already in the group is reported as already there, someone not in it as not there — neither is an error.
- Members of a dynamic group follow its membership rule, so the server refuses to add or remove them by hand.
- Add or remove owners, change a group's name, nickname, description or visibility, and delete a group.

Called by an AI client on behalf of an Entra ID admin; 4 of the 11 tools are read-only, 3 change the tenant without destroying anything, 4 are destructive.

## 🏢 Business domain

Group lifecycle and membership in the directory. Reading a user's own memberships belongs to [Users](../users/index.md); app role assignments to [Applications](../applications/index.md) and [Service principals](../service-principals/index.md).

| Term | Meaning | In the code |
|---|---|---|
| **Group** | A Microsoft 365 or security group in the tenant | `/groups`, `groups.get_group_by_id` |
| **Member** | A directory object (user, group, device, service principal) directly in a group | `/groups/{id}/members` |
| **Owner** | A user who manages a group | `/groups/{id}/owners` |
| **Dynamic group** | A group whose members come from a membership rule | `DynamicMembership` in `groupTypes`, `membershipRule` |
| **Mail nickname** | The mail alias every group needs | `mailNickname` |
| **Reference (`$ref`)** | How Graph links a member or owner to a group: a directory object URL | `ReferenceCreate` in `groups.add_group_member`, `groups.add_group_owner` |

| Rule | Enforced by | A caller who breaks it gets |
|---|---|---|
| A group needs `displayName` and `mailNickname` | the checks in the `create_group` tool and `groups.create_group` | Error result `displayName is required for creating a group` (or `mailNickname …`) |
| A dynamic group needs a membership rule | the `DynamicMembership` check in `create_group` | Error result `membershipRule is required for dynamic membership groups` |
| A create with an existing display name returns that group (a check, not a lock: simultaneous creates can still both succeed) | the `displayName eq` lookup in `groups.create_group` | The existing group, `status: already_exists` |
| Dynamic group members are not edited by hand | the `groupTypes` checks in `add_group_member` and `remove_group_member` | `status: error`, `isDynamicGroup: true`; nothing written |
| A member is added once and removed only when present | `_is_member` in `groups` | `status: already_exists` or `status: not_found` |
| Each listed member is added once at creation | `dict.fromkeys` in `groups.create_group` | — |

## 🚀 Quick Start

No token on the call: the server signs in itself, as its [configuration](../../configuration.md) says, on the first Microsoft Graph call. Set the server up in your AI client first ([Quick Start](../../index.md#-quick-start)); then ask the client in plain words, or send the call yourself:

```json
{"jsonrpc": "2.0", "id": 1, "method": "tools/call", "params": {"name": "add_group_member", "arguments": {"group_id": "all-staff", "member_id": "aoife@contoso.com"}}}
```

The result's `structuredContent`:

```json
{
  "status": "success",
  "message": "Member aoife@contoso.com was added to group all-staff"
}
```

Captured by running the server's real MCP handler against the test suite's in-memory Microsoft Graph (`tests/conftest.py`), not a live tenant: the values are test data, the shape is the server's.

Called again for the same member, the result is `{"status": "already_exists", "message": "Member aoife@contoso.com is already in group all-staff"}`.

## 🔄 End-to-end flow

![An AI client calls add_group_member; the tool reads the group and returns status error for a dynamic group; otherwise the groups module re-checks the group, asks Graph whether the id is already a member, and posts a directoryObjects reference to members/$ref only when it is not, returning success or already_exists.](diagrams/groups-add-member.svg)

`add_group_member` stands for the feature's writes. The group is read twice (once by the tool, once by the resource module) before the membership check; then one `POST …/members/$ref`. There is no transaction across these calls, and the membership check reads Graph with eventual consistency, so a member added moments earlier may be posted again and refused by Graph. `create_group` chains the same member steps after creating the group and adding owners; a failure part-way leaves the group created with what was added so far.

## 📚 Pages

- [Endpoints](endpoints.md) — the 11 tools, each with its mark, arguments, result, errors and an example call
- [Integration](integration.md) — the Microsoft services the feature calls and what happens when they fail

## 🔗 Related features

- [Users](../users/index.md) — reach for it to find the user or read the groups a user is already in.
- [MFA](../mfa/index.md) — reach for it to report MFA status for every member of a group.
- [Conditional access](../conditional-access/index.md) — reach for it to see which policies include or exclude a group.

## ❓ Open questions

| Question | Why it matters | Checked | Who can answer |
|---|---|---|---|
| Is the prefix match of `search_groups_by_name` intended, or should it match anywhere in the name as its tool description ("partial match") says? | Admins searching for a word in the middle of a name get nothing back. | `groups.search_groups_by_name` filters `startswith(displayName, …)`; the tool's description says case-insensitive partial match. | Maintainer (baoduy) |
