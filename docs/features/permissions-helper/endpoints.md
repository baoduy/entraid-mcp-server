# Permissions helper — Endpoints

Part of [Permissions helper](index.md).

Each endpoint is an MCP tool. An AI client calls it with a JSON-RPC `tools/call` message over the server's standard input; the server answers on standard output.

## 🔌 Endpoints

| Tool | Mark | Purpose | Auth |
|---|---|---|---|
| [`suggest_permissions_for_task`](#suggest_permissions_for_task) | read-only | Suggest delegated and application Graph permissions for a common task, from a built-in table. | none (no Graph call) |
| [`list_permission_categories_and_tasks`](#list_permission_categories_and_tasks) | read-only | List the categories and tasks `suggest_permissions_for_task` knows. | none (no Graph call) |
| [`get_all_graph_permissions`](#get_all_graph_permissions) | read-only | List every delegated and application permission Microsoft Graph publishes. | signed-in identity |
| [`search_permissions`](#search_permissions) | read-only | Search Graph permissions by keyword. | signed-in identity |

### `suggest_permissions_for_task`

Suggest delegated and application Graph permissions for a common task, from a built-in table.

- **Mark:** read-only
- **Auth:** the identity the server signed in as ([configuration](../../configuration.md)); Microsoft Graph checks its permissions, the server checks none; no Graph call at all
- **Microsoft Graph:** None: reads `COMMON_PERMISSION_MAPPINGS` in `permissions_helper`.
- **Request:**

| Field | Type | Required | Default | Rules |
|---|---|---|---|---|
| `task_category` | string | yes | — | One of the categories `list_permission_categories_and_tasks` returns |
| `task_name` | string | yes | — | One of that category's task names |

- **Response:** `status: success` with `description`, `delegated_permissions`, `application_permissions`, `notes`.

  Example `structuredContent`, captured from the test suite's in-memory Graph:

```json
{
  "status": "error",
  "message": "Unknown task name: read_user_profiles",
  "available_tasks": [
    "read_user_profile"
  ]
}
```

  Normal results (not errors):

  | `status` | `message` |
  |---|---|
  | `success` | the suggestion |
  | `error` | `Unknown task category: <category>` with `available_categories`, or `Unknown task name: <name>` with `available_tasks` |

- **Errors:**

| Result | Text | When |
|---|---|---|
| error result | `Error calling tool 'suggest_permissions_for_task': <reason>`; notification `Error suggesting permissions for task <task_category>/<task_name>: <reason>` | An unexpected exception in the tool |

- **Example:**

```json
{"jsonrpc": "2.0", "id": 1, "method": "tools/call", "params": {"name": "suggest_permissions_for_task", "arguments": {"task_category": "groups", "task_name": "manage_groups"}}}
```

### `list_permission_categories_and_tasks`

List the categories and tasks `suggest_permissions_for_task` knows.

- **Mark:** read-only
- **Auth:** the identity the server signed in as ([configuration](../../configuration.md)); Microsoft Graph checks its permissions, the server checks none; no Graph call at all
- **Microsoft Graph:** None: reads `COMMON_PERMISSION_MAPPINGS`.
- **Request:**

  No arguments.

- **Response:** `status: success` and `categories`: `users`, `groups`, `mail`, `calendar`, `files`, `devices`, `audit_logs`, `directory`, each with its `tasks` (`name`, `description`).

  Example `structuredContent`, captured from the test suite's in-memory Graph:

```json
{
  "status": "success",
  "categories": {
    "users": {
      "tasks": [
        {
          "name": "read_user_profile",
          "description": "Read user profile information"
        }
      ]
    },
    "groups": {
      "tasks": [
        {
          "name": "read_user_groups",
          "description": "Read groups a user is a member of"
        }
      ]
    },
    "mail": {
      "tasks": [
        {
          "name": "read_user_mail",
          "description": "Read user's mail"
        }
      ]
    },
    "calendar": {
      "tasks": [
        {
          "name": "read_user_calendar",
          "description": "Read user's calendar"
        }
      ]
    },
    "files": {
      "tasks": [
        {
          "name": "read_user_files",
          "description": "Read user's files"
        }
      ]
    },
    "devices": {
      "tasks": [
        {
          "name": "read_devices",
          "description": "Read device information"
        }
      ]
    },
    "audit_logs": {
      "tasks": [
        {
          "name": "read_audit_logs",
          "description": "Read audit logs"
        }
      ]
    },
    "directory": {
      "tasks": [
        {
          "name": "read_directory",
          "description": "Read directory data (users, groups, apps, etc.)"
        }
      ]
    }
  }
}
```

- **Errors:**

| Result | Text | When |
|---|---|---|
| error result | `Error calling tool 'list_permission_categories_and_tasks': <reason>`; notification `Error listing permission categories and tasks: <reason>` | An unexpected exception in the tool |

- **Example:**

```json
{"jsonrpc": "2.0", "id": 1, "method": "tools/call", "params": {"name": "list_permission_categories_and_tasks", "arguments": {}}}
```

### `get_all_graph_permissions`

List every delegated and application permission Microsoft Graph publishes.

- **Mark:** read-only
- **Auth:** the identity the server signed in as ([configuration](../../configuration.md)); Microsoft Graph checks its permissions, the server checks none
- **Microsoft Graph:** `GET /servicePrincipals/00000003-0000-0000-c000-000000000000` (the Microsoft Graph application id passed as the object id).
- **Request:**

  No arguments.

- **Response:** `status: success`, `delegated_permissions` (from `oauth2PermissionScopes`) and `application_permissions` (from `appRoles`). When Graph returns no object: `{"status": "error", "message": "Microsoft Graph service principal not found"}`.

  Example `structuredContent`, captured from the test suite's in-memory Graph:

```json
{
  "status": "success",
  "delegated_permissions": [
    {
      "id": null,
      "value": "User.Read",
      "type": "delegated",
      "adminConsentDisplayName": "Sign in and read user profile",
      "adminConsentDescription": "Allows users to sign in and read their profile.",
      "userConsentDisplayName": "Sign you in and read your profile",
      "userConsentDescription": "Allows you to sign in and read your profile.",
      "isEnabled": true
    }
  ],
  "application_permissions": [
    {
      "id": null,
      "value": "User.Read.All",
      "type": "application",
      "displayName": "Read all users' full profiles",
      "description": "Allows the app to read user profiles.",
      "isEnabled": true
    }
  ]
}
```

  Normal results (not errors):

  | `status` | `message` |
  |---|---|
  | `success` | the two permission lists |
  | `error` | `Microsoft Graph service principal not found` |

- **Errors:**

| Result | Text | When |
|---|---|---|
| error result | `Error calling tool 'get_all_graph_permissions': <reason>`; notification `Error retrieving Microsoft Graph permissions: <reason>` | Microsoft Graph or sign-in fails |

- **Example:**

```json
{"jsonrpc": "2.0", "id": 1, "method": "tools/call", "params": {"name": "get_all_graph_permissions", "arguments": {}}}
```

### `search_permissions`

Search Graph permissions by keyword.

- **Mark:** read-only
- **Auth:** the identity the server signed in as ([configuration](../../configuration.md)); Microsoft Graph checks its permissions, the server checks none
- **Microsoft Graph:** The `get_all_graph_permissions` call, then a case-insensitive match in the server.
- **Request:**

| Field | Type | Required | Default | Rules |
|---|---|---|---|---|
| `search_term` | string | yes | — | Case-insensitive substring of the permission value, display name or description |
| `permission_type` | string or null | no | `null` | `delegated` or `application` (any case); any other value matches nothing; empty searches both |

- **Response:** `status`, `search_term` (lower-cased), `matching_delegated_permissions`, `matching_application_permissions`, `total_matches`.

  Example `structuredContent`, captured from the test suite's in-memory Graph:

```json
{
  "status": "success",
  "search_term": "user.read",
  "matching_delegated_permissions": [
    {
      "id": null,
      "value": "User.Read",
      "type": "delegated",
      "adminConsentDisplayName": "Sign in and read user profile",
      "adminConsentDescription": "Allows users to sign in and read their profile.",
      "userConsentDisplayName": "Sign you in and read your profile",
      "userConsentDescription": "Allows you to sign in and read your profile.",
      "isEnabled": true
    }
  ],
  "matching_application_permissions": [
    {
      "id": null,
      "value": "User.Read.All",
      "type": "application",
      "displayName": "Read all users' full profiles",
      "description": "Allows the app to read user profiles.",
      "isEnabled": true
    }
  ],
  "total_matches": 2
}
```

  Normal results (not errors):

  | `status` | `message` |
  |---|---|
  | `success` | the matches |
  | `error` | `Microsoft Graph service principal not found` |

- **Errors:**

| Result | Text | When |
|---|---|---|
| error result | `Error calling tool 'search_permissions': <reason>`; notification `Error searching for permissions with term '<search_term>': <reason>` | Microsoft Graph or sign-in fails |

- **Example:**

```json
{"jsonrpc": "2.0", "id": 1, "method": "tools/call", "params": {"name": "search_permissions", "arguments": {"search_term": "User.Read"}}}
```

## ⚠️ Errors & limits

A failed call reaches the AI client twice: one error notification (`<message>: <reason>`, sent through the MCP log channel) and the error result `Error calling tool '<tool>': <reason>` with `isError: true`. The reason is the text of the exception Microsoft Graph, the SDK or the server raised. The server sets no timeout and no retry of its own and keeps no state between calls. Results that report a known case — `already_exists`, `not_found`, a dynamic group's `status: error`, the permission helper's `status: error` — are normal results, not errors.

- `suggest_permissions_for_task` and `list_permission_categories_and_tasks` read a built-in table: 8 categories, 19 tasks; they never call Graph.
- `get_all_graph_permissions` and `search_permissions` read the Microsoft Graph service principal by passing its application id as the object id — see ❓ Open questions in [index.md](index.md).
- `search_permissions` downloads the whole permission list on every call and filters it in the server.
