# Permissions helper

Work out which Microsoft Graph permissions a task needs, and search the full list Graph publishes.

## 📖 Overview

- Ask for the delegated and application permissions for a common task (read users, manage groups, read audit logs, …) from a built-in table.
- List the categories and tasks that table knows: 8 categories, 19 tasks.
- List every permission Microsoft Graph publishes, or search it by keyword and type.
- All four tools are read-only; the first two never call Graph.

Called by an AI client on behalf of an admin setting up an app registration with least privilege.

## 🏢 Business domain

Permission guidance. Granting a permission to an app belongs to [Applications](../applications/index.md) and [Service principals](../service-principals/index.md); this feature only suggests and lists.

| Term | Meaning | In the code |
|---|---|---|
| **Delegated permission** | A permission an app uses on behalf of a signed-in user (`oauth2PermissionScopes`) | `delegated_permissions` |
| **Application permission** | A permission an app uses as itself (`appRoles`) | `application_permissions` |
| **Task** | A named job in the built-in table, such as `groups/manage_groups` | `COMMON_PERMISSION_MAPPINGS` |

| Rule | Enforced by | A caller who breaks it gets |
|---|---|---|
| A suggestion exists only for a known category and task | the lookups in `suggest_permissions_for_task` | `status: error` with `available_categories` or `available_tasks` |
| The search is case-insensitive over value, display name and description | `search_permissions` | — |

## 🚀 Quick Start

No token on the call: the server signs in itself, as its [configuration](../../configuration.md) says, on the first Microsoft Graph call. Set the server up in your AI client first ([Quick Start](../../index.md#-quick-start)); then ask the client in plain words, or send the call yourself:

```json
{"jsonrpc": "2.0", "id": 1, "method": "tools/call", "params": {"name": "suggest_permissions_for_task", "arguments": {"task_category": "groups", "task_name": "manage_groups"}}}
```

The result's `structuredContent`:

```json
{
  "status": "success",
  "task_category": "groups",
  "task_name": "manage_groups",
  "description": "Create, update, and delete groups, and add/remove members",
  "delegated_permissions": [
    "Group.ReadWrite.All"
  ],
  "application_permissions": [
    "Group.ReadWrite.All"
  ],
  "notes": "These are suggested permissions based on common usage patterns. Always follow the principle of least privilege."
}
```

Captured by running the server's real MCP handler against the test suite's in-memory Microsoft Graph (`tests/conftest.py`), not a live tenant: the values are test data, the shape is the server's.

## 🔄 End-to-end flow

![An AI client calls search_permissions; the module reads the Microsoft Graph service principal by its application id, matches the term against its scopes and app roles in the server, and returns the matches or, when Graph refuses, an error.](diagrams/permissions-search.svg)

`search_permissions` stands for the feature's Graph route. Nothing is stored: the full permission list is downloaded on every call and filtered in the server. The suggestion tools use a fixed table in the code and need no Graph access at all, so they work before any permission is granted.

## 📚 Pages

- [Endpoints](endpoints.md) — the 4 tools, each with its mark, arguments, result, errors and an example call
- [Integration](integration.md) — the Microsoft services the feature calls and what happens when they fail

## 🔗 Related features

- [Applications](../applications/index.md) — reach for it to set `requiredResourceAccess` on an app registration once you know the permission.
- [Service principals](../service-principals/index.md) — reach for it to see the permissions an app already holds.

## ❓ Open questions

| Question | Why it matters | Checked | Who can answer |
|---|---|---|---|
| Does `GET /servicePrincipals/00000003-0000-0000-c000-000000000000` resolve in a real tenant? | `get_all_graph_permissions` and `search_permissions` pass Microsoft Graph's application id where Graph expects the service principal's object id; if Graph answers 404, both tools always fail. | `permissions_helper.get_all_graph_permissions` calls `by_service_principal_id(MS_GRAPH_APP_ID)`, while `service_principals.get_service_principal_by_app_id` finds a service principal with `$filter=appId eq …`. The test suite's Graph accepts any id. | Maintainer (baoduy), by a run against a tenant |
