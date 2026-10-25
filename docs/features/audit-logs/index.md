# Audit logs

See what a user changed in the directory over the last days, with the before and after values.

## 📖 Overview

- Read every directory audit entry a user initiated in a window of days (30 by default), newest first.
- Each entry names the activity, its result, who started it (user and app) and the targets, with each changed property's old and new value.

Called by an AI client on behalf of an Entra ID admin investigating a change. The tool is read-only.

## 🏢 Business domain

Directory change history. Sign-in events belong to [Sign-in logs](../sign-in-logs/index.md).

| Term | Meaning | In the code |
|---|---|---|
| **Directory audit entry** | One recorded change in the directory | `/auditLogs/directoryAudits` |
| **Initiated by** | The user or app that made the change | `initiatedBy` |
| **Target resource** | What was changed, with its modified properties | `targetResources`, `modifiedProperties` |

| Rule | Enforced by | A caller who breaks it gets |
|---|---|---|
| Only entries the user initiated are returned | the `initiatedBy/user/id` filter in `audit_logs.get_user_audit_logs` | — |
| The user id cannot add a filter condition: a `'` is doubled | `odata_string_literal` in `utils/odata.py` | Their text matched literally |

## 🚀 Quick Start

No token on the call: the server signs in itself, as its [configuration](../../configuration.md) says, on the first Microsoft Graph call. Set the server up in your AI client first ([Quick Start](../../index.md#-quick-start)); then ask the client in plain words, or send the call yourself:

```json
{"jsonrpc": "2.0", "id": 1, "method": "tools/call", "params": {"name": "get_user_audit_logs", "arguments": {"user_id": "7f3c9a", "days": 30}}}
```

The result's `structuredContent` — list and optional results arrive under `result`:

```json
{
  "result": [
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
}
```

Captured by running the server's real MCP handler against the test suite's in-memory Microsoft Graph (`tests/conftest.py`), not a live tenant: the values are test data, the shape is the server's.

## 🔄 End-to-end flow

![An AI client calls get_user_audit_logs; the module sends GET directoryAudits filtered on initiatedBy user id and the window, newest first with top 1000, follows every nextLink, and returns the entries or, when Graph refuses, an error.](diagrams/audit-logs-user.svg)

Nothing is stored and no event is raised. Every page is read. Changes made *to* the user by someone else are not returned, because the filter is on who initiated the change.

## 📚 Pages

- [Endpoints](endpoints.md) — the 1 tool, each with its mark, arguments, result, errors and an example call
- [Integration](integration.md) — the Microsoft services the feature calls and what happens when they fail

## 🔗 Related features

- [Sign-in logs](../sign-in-logs/index.md) — reach for it to see the user's sign-ins instead of their changes.
