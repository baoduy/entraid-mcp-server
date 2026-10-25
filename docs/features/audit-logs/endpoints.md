# Audit logs — Endpoints

Part of [Audit logs](index.md).

Each endpoint is an MCP tool. An AI client calls it with a JSON-RPC `tools/call` message over the server's standard input; the server answers on standard output.

## 🔌 Endpoints

| Tool | Mark | Purpose | Auth |
|---|---|---|---|
| [`get_user_audit_logs`](#get_user_audit_logs) | read-only | List the directory audit entries a user initiated over the last `days` days, newest first. | signed-in identity |

### `get_user_audit_logs`

List the directory audit entries a user initiated over the last `days` days, newest first.

- **Mark:** read-only
- **Auth:** the identity the server signed in as ([configuration](../../configuration.md)); Microsoft Graph checks its permissions, the server checks none
- **Microsoft Graph:** `GET /auditLogs/directoryAudits` with `$filter=initiatedBy/user/id eq '<user_id>' and activityDateTime ge <start> and activityDateTime le <end>`, `$orderby=activityDateTime desc`, `$top=1000`, `ConsistencyLevel: eventual`, every page.
- **Request:**

| Field | Type | Required | Default | Rules |
|---|---|---|---|---|
| `user_id` | string | yes | — | — |
| `days` | integer | no | `30` | Window is now (UTC) minus `days` to now |

- **Response:** Each entry has `id`, `activityDateTime`, `activityDisplayName`, `category`, `operationType`, `result`, `resultReason`, `initiatedBy` (user and app), `targetResources` (with `modifiedProperties`), `loggedByService`, `correlationId`, `additionalDetails`.

  Example `structuredContent`, captured from the test suite's in-memory Graph:

```json
[
  {
    "id": "audit-1",
    "activityDateTime": "2026-10-20T07:45:00+00:00",
    "activityDisplayName": "Update user",
    "category": null,
    "operationType": null,
    "result": null,
    "resultReason": null,
    "initiatedBy": {
      "user": {
        "id": "7f3c9a",
        "displayName": "Mai",
        "userPrincipalName": "mai@contoso.com"
      },
      "app": {
        "appId": "app-1",
        "displayName": "Portal"
      }
    },
    "targetResources": [
      {
        "id": "aoife",
        "displayName": "Aoife",
        "type": null,
        "userPrincipalName": null,
        "modifiedProperties": [
          {
            "displayName": "JobTitle",
            "oldValue": "a",
            "newValue": "b"
          }
        ]
      }
    ],
    "loggedByService": null,
    "correlationId": null,
    "additionalDetails": [
      {
        "key": "UserType",
        "value": "Member"
      }
    ]
  }
]
```

- **Errors:**

| Result | Text | When |
|---|---|---|
| error result | `Error calling tool 'get_user_audit_logs': <reason>`; notification `Error fetching directory audit logs for user <user_id>: <reason>` | Microsoft Graph or sign-in fails |

- **Example:**

```json
{"jsonrpc": "2.0", "id": 1, "method": "tools/call", "params": {"name": "get_user_audit_logs", "arguments": {"user_id": "aoife@contoso.com"}}}
```

## ⚠️ Errors & limits

A failed call reaches the AI client twice: one error notification (`<message>: <reason>`, sent through the MCP log channel) and the error result `Error calling tool '<tool>': <reason>` with `isError: true`. The reason is the text of the exception Microsoft Graph, the SDK or the server raised. The server sets no timeout and no retry of its own and keeps no state between calls. Results that report a known case — `already_exists`, `not_found`, a dynamic group's `status: error`, the permission helper's `status: error` — are normal results, not errors.

- Only entries the user initiated (`initiatedBy/user/id`) are returned, not entries where the user is the target.
- The window is computed in UTC at call time; every page is read.
