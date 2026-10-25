# Groups — Endpoints

Part of [Groups](index.md).

Each endpoint is an MCP tool. An AI client calls it with a JSON-RPC `tools/call` message over the server's standard input; the server answers on standard output.

## 🔌 Endpoints

| Tool | Mark | Purpose | Auth |
|---|---|---|---|
| [`get_all_groups`](#get_all_groups) | read-only | List groups, up to `limit`. | signed-in identity |
| [`get_group_by_id`](#get_group_by_id) | read-only | Get one group by id. | signed-in identity |
| [`search_groups_by_name`](#search_groups_by_name) | read-only | Find groups whose display name starts with the text given. | signed-in identity |
| [`get_group_members`](#get_group_members) | read-only | List a group's direct members, up to `limit`. | signed-in identity |
| [`create_group`](#create_group) | changes the tenant, not destructive | Create a group, then add the listed owners and members. | signed-in identity |
| [`update_group`](#update_group) | destructive | Change a group's display name, mail nickname, description or visibility. | signed-in identity |
| [`delete_group`](#delete_group) | destructive | Delete a group. | signed-in identity |
| [`add_group_member`](#add_group_member) | changes the tenant, not destructive | Add a user, group, device or other directory object to a group. | signed-in identity |
| [`remove_group_member`](#remove_group_member) | destructive | Remove a direct member from a group. | signed-in identity |
| [`add_group_owner`](#add_group_owner) | changes the tenant, not destructive | Add an owner to a group. | signed-in identity |
| [`remove_group_owner`](#remove_group_owner) | destructive | Remove an owner from a group. | signed-in identity |

### `get_all_groups`

List groups, up to `limit`.

- **Mark:** read-only
- **Auth:** the identity the server signed in as ([configuration](../../configuration.md)); Microsoft Graph checks its permissions, the server checks none
- **Microsoft Graph:** `GET /groups?$top=<min(limit, 999)>`, following `@odata.nextLink` until `limit` groups are held.
- **Request:**

| Field | Type | Required | Default | Rules |
|---|---|---|---|---|
| `limit` | integer | no | `100` | 0 or less returns `[]` without a Graph call; page size is `limit`, capped at 999 |

- **Response:** Each group has `id`, `displayName`, `mail`, `mailNickname`, `description`, `groupTypes`, `securityEnabled`, `mailEnabled`, `createdDateTime`.

  Example `structuredContent`, captured from the test suite's in-memory Graph:

```json
[
  {
    "id": "group-0000",
    "displayName": "Group 0000",
    "mail": "group0000@contoso.com",
    "mailNickname": "group0000",
    "description": "Group number 0",
    "groupTypes": [
      "Unified"
    ],
    "securityEnabled": false,
    "mailEnabled": true,
    "createdDateTime": "2026-01-15T09:30:00+00:00"
  }
]
```

- **Errors:**

| Result | Text | When |
|---|---|---|
| error result | `Error calling tool 'get_all_groups': <reason>`; notification `Error fetching all groups: <reason>` | Microsoft Graph or sign-in fails |

- **Example:**

```json
{"jsonrpc": "2.0", "id": 1, "method": "tools/call", "params": {"name": "get_all_groups", "arguments": {}}}
```

### `get_group_by_id`

Get one group by id.

- **Mark:** read-only
- **Auth:** the identity the server signed in as ([configuration](../../configuration.md)); Microsoft Graph checks its permissions, the server checks none
- **Microsoft Graph:** `GET /groups/{group_id}`.
- **Request:**

| Field | Type | Required | Default | Rules |
|---|---|---|---|---|
| `group_id` | string | yes | — | — |

- **Response:** The `get_all_groups` fields plus `visibility`, or `null` with the warning `Group with ID <group_id> not found.` when Graph returns no object.

  Example `structuredContent`, captured from the test suite's in-memory Graph:

```json
{
  "id": "all-staff",
  "displayName": "All Staff",
  "mail": "allstaff@contoso.com",
  "mailNickname": "allstaff",
  "description": "Everyone at Contoso",
  "groupTypes": [
    "Unified"
  ],
  "securityEnabled": false,
  "mailEnabled": true,
  "visibility": "Private",
  "createdDateTime": "2026-10-20T07:45:00+00:00"
}
```

- **Errors:**

| Result | Text | When |
|---|---|---|
| error result | `Error calling tool 'get_group_by_id': <reason>`; notification `Error fetching group <group_id>: <reason>` | Microsoft Graph or sign-in fails |

- **Example:**

```json
{"jsonrpc": "2.0", "id": 1, "method": "tools/call", "params": {"name": "get_group_by_id", "arguments": {"group_id": "all-staff"}}}
```

### `search_groups_by_name`

Find groups whose display name starts with the text given.

- **Mark:** read-only
- **Auth:** the identity the server signed in as ([configuration](../../configuration.md)); Microsoft Graph checks its permissions, the server checks none
- **Microsoft Graph:** `GET /groups?$filter=startswith(displayName,'<name>')`, following `@odata.nextLink` until `limit` groups are held.
- **Request:**

| Field | Type | Required | Default | Rules |
|---|---|---|---|---|
| `name` | string | yes | — | Prefix match: `startswith(displayName, '<name>')`; a `'` is doubled |
| `limit` | integer | no | `50` | 0 or less returns `[]` without a Graph call; page size capped at 999 |

- **Response:** The `get_all_groups` fields.

  Example `structuredContent`, captured from the test suite's in-memory Graph:

```json
[
  {
    "id": "group-0000",
    "displayName": "All 0000",
    "mail": "group0000@contoso.com",
    "mailNickname": "group0000",
    "description": "All number 0",
    "groupTypes": [
      "Unified"
    ],
    "securityEnabled": false,
    "mailEnabled": true,
    "createdDateTime": "2026-01-15T09:30:00+00:00"
  }
]
```

- **Errors:**

| Result | Text | When |
|---|---|---|
| error result | `Error calling tool 'search_groups_by_name': <reason>`; notification `Error searching groups by name: <reason>` | Microsoft Graph or sign-in fails |

- **Example:**

```json
{"jsonrpc": "2.0", "id": 1, "method": "tools/call", "params": {"name": "search_groups_by_name", "arguments": {"name": "All Staff"}}}
```

### `get_group_members`

List a group's direct members, up to `limit`.

- **Mark:** read-only
- **Auth:** the identity the server signed in as ([configuration](../../configuration.md)); Microsoft Graph checks its permissions, the server checks none
- **Microsoft Graph:** `GET /groups/{group_id}/members?$top=<min(limit, 999)>`, following `@odata.nextLink` until `limit` members are held.
- **Request:**

| Field | Type | Required | Default | Rules |
|---|---|---|---|---|
| `group_id` | string | yes | — | — |
| `limit` | integer | no | `100` | 0 or less returns `[]` without a Graph call; page size capped at 999 |

- **Response:** Each member has the 10 user profile fields plus `type` (the member's `@odata.type`); fields a non-user member lacks are `null`.

  Example `structuredContent`, captured from the test suite's in-memory Graph:

```json
[
  {
    "id": "aoife@contoso.com",
    "displayName": "Aoife Byrne",
    "mail": "aoife@contoso.com",
    "userPrincipalName": "aoife@contoso.com",
    "givenName": "Aoife",
    "surname": "Byrne",
    "jobTitle": "Engineer",
    "officeLocation": "Dublin",
    "businessPhones": [
      "+353 1 555 0100"
    ],
    "mobilePhone": "+353 87 555 0100",
    "type": "#microsoft.graph.user"
  }
]
```

- **Errors:**

| Result | Text | When |
|---|---|---|
| error result | `Error calling tool 'get_group_members': <reason>`; notification `Error fetching group members for group <group_id>: <reason>` | Microsoft Graph or sign-in fails |

- **Example:**

```json
{"jsonrpc": "2.0", "id": 1, "method": "tools/call", "params": {"name": "get_group_members", "arguments": {"group_id": "all-staff"}}}
```

### `create_group`

Create a group, then add the listed owners and members.

- **Mark:** changes the tenant, not destructive
- **Auth:** the identity the server signed in as ([configuration](../../configuration.md)); Microsoft Graph checks its permissions, the server checks none
- **Idempotency:** By display name: when a group with the same `displayName` exists, that group is returned with `status: already_exists` and nothing is created — the listed owners and members are not added to it either.
- **Microsoft Graph:** `GET /groups?$filter=displayName eq '<displayName>'`; when no group matches, `POST /groups`, then `POST /groups/{id}/owners/$ref` per owner, then the `add_group_member` steps per distinct member (skipped for a dynamic group).
- **Request:**

| Field | Type | Required | Default | Rules |
|---|---|---|---|---|
| `group_data` | object | yes | — | See the `group_data` keys below |

`group_data` keys:

| Key | Type | Required | Rules |
|---|---|---|---|
| `displayName` | string | yes | Checked for an existing group with the same name first |
| `mailNickname` | string | yes | — |
| `description` | string | no | — |
| `groupTypes` | list of strings | no | `["Unified"]` for Microsoft 365; `DynamicMembership` makes a dynamic group |
| `mailEnabled` | boolean | no | — |
| `securityEnabled` | boolean | no | — |
| `visibility` | string | no | `Private` or `Public` |
| `owners` | list of ids | no | Must be a list; each id is posted to `owners/$ref` after the group is created |
| `members` | list of ids | no | Must be a list (not checked for a dynamic group); each distinct id is added after the owners; ignored for a dynamic group |
| `membershipRule` | string | for a dynamic group | — |
| `membershipRuleProcessingState` | string | no | `On` (default) or `Paused`; dynamic groups only |


- **Response:** The new group's `get_group_by_id` fields (plus `membershipRule` and `membershipRuleProcessingState` for a dynamic group), or the existing group's fields with `"status": "already_exists"`.

  Example `structuredContent`, captured from the test suite's in-memory Graph:

```json
{
  "id": "new-0001",
  "displayName": "All Staff",
  "mail": null,
  "mailNickname": "allstaff",
  "description": null,
  "groupTypes": null,
  "securityEnabled": null,
  "mailEnabled": null,
  "visibility": null,
  "createdDateTime": null
}
```

- **Errors:**

| Result | Text | When |
|---|---|---|
| error result | `displayName is required for creating a group` | `displayName` missing. Before any Graph call |
| error result | `mailNickname is required for creating a group` | `mailNickname` missing. Before any Graph call |
| error result | `membershipRule is required for dynamic membership groups` | `groupTypes` holds `DynamicMembership` and `membershipRule` is missing. Before any Graph call |
| error result | `owners must be a list of user IDs` | `owners` is not a list. After the duplicate-name check, before the group is created |
| error result | `members must be a list of user IDs` | `members` is not a list (non-dynamic group). After the duplicate-name check, before the group is created |
| error result | the Graph error | An owner or member add fails after the group was created. The group stays created with the owners and members added so far; nothing is rolled back |
| error result | `Error calling tool 'create_group': <reason>`; notification `Error creating group: <reason>` | Microsoft Graph or sign-in fails |

- **Example:**

```json
{"jsonrpc": "2.0", "id": 1, "method": "tools/call", "params": {"name": "create_group", "arguments": {"group_data": {"displayName": "All Staff", "mailNickname": "allstaff"}}}}
```

### `update_group`

Change a group's display name, mail nickname, description or visibility.

- **Mark:** destructive
- **Auth:** the identity the server signed in as ([configuration](../../configuration.md)); Microsoft Graph checks its permissions, the server checks none
- **Idempotency:** Not idempotent: each call is sent to Graph again.
- **Concurrency:** none: no ETag or version is sent, so the last write wins.
- **Microsoft Graph:** `PATCH /groups/{group_id}` with the given keys, then `GET /groups/{group_id}`.
- **Request:**

| Field | Type | Required | Default | Rules |
|---|---|---|---|---|
| `group_id` | string | yes | — | — |
| `group_data` | object | yes | — | Only `displayName`, `mailNickname`, `description`, `visibility` are sent; other keys are ignored |

- **Response:** The group's `get_group_by_id` fields after the change.

  Example `structuredContent`, captured from the test suite's in-memory Graph:

```json
{
  "id": "all-staff",
  "displayName": "All Staff",
  "mail": "allstaff@contoso.com",
  "mailNickname": "allstaff",
  "description": "Everyone at Contoso",
  "groupTypes": [
    "Unified"
  ],
  "securityEnabled": false,
  "mailEnabled": true,
  "visibility": "Private",
  "createdDateTime": "2026-10-20T07:45:00+00:00"
}
```

- **Errors:**

| Result | Text | When |
|---|---|---|
| error result | `Failed to retrieve updated group with ID <group_id>` | The group cannot be read back. After the PATCH succeeded |
| error result | `Error calling tool 'update_group': <reason>`; notification `Error updating group <group_id>: <reason>` | Microsoft Graph or sign-in fails |

- **Example:**

```json
{"jsonrpc": "2.0", "id": 1, "method": "tools/call", "params": {"name": "update_group", "arguments": {"group_id": "all-staff", "group_data": {"displayName": "All Staff"}}}}
```

### `delete_group`

Delete a group.

- **Mark:** destructive
- **Auth:** the identity the server signed in as ([configuration](../../configuration.md)); Microsoft Graph checks its permissions, the server checks none
- **Idempotency:** Not idempotent: a second call reaches Graph again and fails with Graph's error once the group is gone.
- **Microsoft Graph:** `DELETE /groups/{group_id}`.
- **Request:**

| Field | Type | Required | Default | Rules |
|---|---|---|---|---|
| `group_id` | string | yes | — | — |

- **Response:** `{"status": "success", "message": "Group <group_id> was deleted successfully"}`.

  Example `structuredContent`, captured from the test suite's in-memory Graph:

```json
{
  "status": "success",
  "message": "Group all-staff was deleted successfully"
}
```

- **Errors:**

| Result | Text | When |
|---|---|---|
| error result | `Error calling tool 'delete_group': <reason>`; notification `Error deleting group <group_id>: <reason>` | Microsoft Graph or sign-in fails |

- **Example:**

```json
{"jsonrpc": "2.0", "id": 1, "method": "tools/call", "params": {"name": "delete_group", "arguments": {"group_id": "all-staff"}}}
```

### `add_group_member`

Add a user, group, device or other directory object to a group.

- **Mark:** changes the tenant, not destructive
- **Auth:** the identity the server signed in as ([configuration](../../configuration.md)); Microsoft Graph checks its permissions, the server checks none
- **Idempotency:** Yes: an existing direct member gets `{"status": "already_exists", "message": "Member <member_id> is already in group <group_id>"}` and nothing is posted. The check reads Graph with eventual consistency, so a member added moments earlier may not be seen yet.
- **Microsoft Graph:** `GET /groups/{group_id}` (twice: once in the tool, once in the resource), then `GET /groups/{group_id}/members?$filter=id eq '<member_id>'&$count=true` with `ConsistencyLevel: eventual`, then `POST /groups/{group_id}/members/$ref` with `@odata.id` = `https://graph.microsoft.com/v1.0/directoryObjects/<member_id>`.
- **Request:**

| Field | Type | Required | Default | Rules |
|---|---|---|---|---|
| `group_id` | string | yes | — | — |
| `member_id` | string | yes | — | — |

- **Response:** `status` and `message` — see the normal results below.

  Example `structuredContent`, captured from the test suite's in-memory Graph:

```json
{
  "status": "success",
  "message": "Member aoife@contoso.com was added to group all-staff"
}
```

  Normal results (not errors):

  | `status` | `message` |
  |---|---|
  | `success` | `Member <member_id> was added to group <group_id>` |
  | `already_exists` | `Member <member_id> is already in group <group_id>` |
  | `error` | `Cannot add members to a dynamic membership group. Members are determined by the membership rule.` with `groupId`, `memberId`, `isDynamicGroup: true` |

- **Errors:**

| Result | Text | When |
|---|---|---|
| error result | `Group with ID <group_id> not found` (notification is this text alone) | Graph returns no group. Before any write |
| error result | notification is the error text alone | The server's own refusal (`ValueError`) |
| error result | `Error calling tool 'add_group_member': <reason>`; notification `Error adding member <member_id> to group <group_id>: <reason>` | Microsoft Graph or sign-in fails |

- **Example:**

```json
{"jsonrpc": "2.0", "id": 1, "method": "tools/call", "params": {"name": "add_group_member", "arguments": {"group_id": "all-staff", "member_id": "aoife@contoso.com"}}}
```

### `remove_group_member`

Remove a direct member from a group.

- **Mark:** destructive
- **Auth:** the identity the server signed in as ([configuration](../../configuration.md)); Microsoft Graph checks its permissions, the server checks none
- **Idempotency:** Yes: someone who is not a direct member gets `{"status": "not_found", "message": "Member <member_id> is not in group <group_id>"}` and nothing is deleted.
- **Microsoft Graph:** `GET /groups/{group_id}` (twice), then the same membership check as `add_group_member`, then `DELETE /groups/{group_id}/members/{member_id}/$ref`.
- **Request:**

| Field | Type | Required | Default | Rules |
|---|---|---|---|---|
| `group_id` | string | yes | — | — |
| `member_id` | string | yes | — | — |

- **Response:** `status` and `message` — see the normal results below.

  Example `structuredContent`, captured from the test suite's in-memory Graph:

```json
{
  "status": "success",
  "message": "Member aoife@contoso.com was removed from group all-staff"
}
```

  Normal results (not errors):

  | `status` | `message` |
  |---|---|
  | `success` | `Member <member_id> was removed from group <group_id>` |
  | `not_found` | `Member <member_id> is not in group <group_id>` |
  | `error` | `Cannot remove members from a dynamic membership group. Members are determined by the membership rule.` with `groupId`, `memberId`, `isDynamicGroup: true` |

- **Errors:**

| Result | Text | When |
|---|---|---|
| error result | `Group with ID <group_id> not found` (notification is this text alone) | Graph returns no group. Before any write |
| error result | notification is the error text alone | The server's own refusal (`ValueError`) |
| error result | `Error calling tool 'remove_group_member': <reason>`; notification `Error removing member <member_id> from group <group_id>: <reason>` | Microsoft Graph or sign-in fails |

- **Example:**

```json
{"jsonrpc": "2.0", "id": 1, "method": "tools/call", "params": {"name": "remove_group_member", "arguments": {"group_id": "all-staff", "member_id": "aoife@contoso.com"}}}
```

### `add_group_owner`

Add an owner to a group.

- **Mark:** changes the tenant, not destructive
- **Auth:** the identity the server signed in as ([configuration](../../configuration.md)); Microsoft Graph checks its permissions, the server checks none
- **Idempotency:** No check: adding an existing owner is posted again and Graph's answer decides the result.
- **Microsoft Graph:** `POST /groups/{group_id}/owners/$ref` with `@odata.id` = `https://graph.microsoft.com/v1.0/directoryObjects/<owner_id>`.
- **Request:**

| Field | Type | Required | Default | Rules |
|---|---|---|---|---|
| `group_id` | string | yes | — | — |
| `owner_id` | string | yes | — | — |

- **Response:** `{"status": "success", "message": "Owner <owner_id> was added to group <group_id>"}`.

  Example `structuredContent`, captured from the test suite's in-memory Graph:

```json
{
  "status": "success",
  "message": "Owner aoife@contoso.com was added to group all-staff"
}
```

- **Errors:**

| Result | Text | When |
|---|---|---|
| error result | `Error calling tool 'add_group_owner': <reason>`; notification `Error adding owner <owner_id> to group <group_id>: <reason>` | Microsoft Graph or sign-in fails |

- **Example:**

```json
{"jsonrpc": "2.0", "id": 1, "method": "tools/call", "params": {"name": "add_group_owner", "arguments": {"group_id": "all-staff", "owner_id": "aoife@contoso.com"}}}
```

### `remove_group_owner`

Remove an owner from a group.

- **Mark:** destructive
- **Auth:** the identity the server signed in as ([configuration](../../configuration.md)); Microsoft Graph checks its permissions, the server checks none
- **Idempotency:** No check: removing someone who is not an owner is sent to Graph and Graph's answer decides the result.
- **Microsoft Graph:** `DELETE /groups/{group_id}/owners/{owner_id}/$ref`.
- **Request:**

| Field | Type | Required | Default | Rules |
|---|---|---|---|---|
| `group_id` | string | yes | — | — |
| `owner_id` | string | yes | — | — |

- **Response:** `{"status": "success", "message": "Owner <owner_id> was removed from group <group_id>"}`.

  Example `structuredContent`, captured from the test suite's in-memory Graph:

```json
{
  "status": "success",
  "message": "Owner aoife@contoso.com was removed from group all-staff"
}
```

- **Errors:**

| Result | Text | When |
|---|---|---|
| error result | `Error calling tool 'remove_group_owner': <reason>`; notification `Error removing owner <owner_id> from group <group_id>: <reason>` | Microsoft Graph or sign-in fails |

- **Example:**

```json
{"jsonrpc": "2.0", "id": 1, "method": "tools/call", "params": {"name": "remove_group_owner", "arguments": {"group_id": "all-staff", "owner_id": "aoife@contoso.com"}}}
```

## ⚠️ Errors & limits

A failed call reaches the AI client twice: one error notification (`<message>: <reason>`, sent through the MCP log channel) and the error result `Error calling tool '<tool>': <reason>` with `isError: true`. The reason is the text of the exception Microsoft Graph, the SDK or the server raised. The server sets no timeout and no retry of its own and keeps no state between calls. Results that report a known case — `already_exists`, `not_found`, a dynamic group's `status: error`, the permission helper's `status: error` — are normal results, not errors.

- `search_groups_by_name` is a prefix match (`startswith`), not a contains match, though its tool description says partial match.
- Page size for `get_all_groups`, `search_groups_by_name` and `get_group_members` is `limit`, capped at 999 (Graph's largest `$top` there); a `limit` of 0 or less returns `[]` without a Graph call.
- Membership checks in `add_group_member` and `remove_group_member` read Graph with `ConsistencyLevel: eventual`: a change made moments earlier may not be seen yet.
- Members of a dynamic group cannot be added or removed: both tools return `status: error` with `isDynamicGroup: true` and write nothing.
- `create_group` is several calls with no transaction: when an owner or member add fails, the group stays created with what was added so far.
- `create_group` with an existing display name returns that group (`status: already_exists`) and adds none of the listed owners or members.
- `add_group_owner` and `remove_group_owner` do no membership check of their own; Graph's answer decides a repeat.
