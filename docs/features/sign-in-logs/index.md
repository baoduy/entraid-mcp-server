# Sign-in logs

Review how and where a user signed in over the last days: app, address, device, location, outcome and risk.

## 📖 Overview

- Read every sign-in of one user in a window of days you choose (7 by default), newest first.
- Each sign-in carries its outcome (`status.errorCode`, `failureReason`), its risk levels, and the device and location when Microsoft records them.
- Every page is read, so the list is complete for the window.

Called by an AI client on behalf of an Entra ID admin investigating access.

## 🏢 Business domain

Sign-in activity. What a user changed belongs to [Audit logs](../audit-logs/index.md); registered methods to [MFA](../mfa/index.md).

| Term | Meaning | In the code |
|---|---|---|
| **Sign-in** | One authentication event of a user | `/auditLogs/signIns`, `get_user_sign_in_logs` |
| **Window** | The last `days` days up to now, in UTC | `days`, the `createdDateTime` filter |
| **Risk information** | Entra ID Protection's risk detail, levels, state and event types for the sign-in | `riskInformation` |

| Rule | Enforced by | A caller who breaks it gets |
|---|---|---|
| Only the asked user and window are read | the `$filter` in `get_user_sign_in_logs` | — |
| The user id cannot add a filter condition: a `'` is doubled | `odata_string_literal` in `utils/odata.py` | Their text matched literally |

## 🚀 Quick Start

No token on the call: the server signs in itself, as its [configuration](../../configuration.md) says, on the first Microsoft Graph call. Set the server up in your AI client first ([Quick Start](../../index.md#-quick-start)); then ask the client in plain words, or send the call yourself:

```json
{"jsonrpc": "2.0", "id": 1, "method": "tools/call", "params": {"name": "get_user_sign_ins", "arguments": {"user_id": "7f3c9a", "days": 7}}}
```

The result's `structuredContent` — list and optional results arrive under `result`:

```json
{
  "result": [
    {
      "id": "sign-in-1",
      "createdDateTime": "2026-10-20T07:45:00+00:00",
      "userId": "7f3c9a",
      "userDisplayName": null,
      "userPrincipalName": null,
      "appDisplayName": null,
      "appId": null,
      "ipAddress": null,
      "clientAppUsed": null,
      "correlationId": null,
      "isInteractive": null,
      "resourceDisplayName": null,
      "status": {
        "errorCode": 0,
        "failureReason": null,
        "additionalDetails": null
      },
      "riskInformation": {
        "riskDetail": null,
        "riskLevelAggregated": null,
        "riskLevelDuringSignIn": null,
        "riskState": null,
        "riskEventTypes": null
      },
      "deviceDetail": {
        "deviceId": "device-1",
        "displayName": "LAPTOP-0042",
        "operatingSystem": null,
        "browser": null,
        "isCompliant": null,
        "isManaged": null,
        "trustType": null
      },
      "location": {
        "city": "Dublin",
        "state": null,
        "countryOrRegion": "IE",
        "coordinates": {
          "latitude": 53.35,
          "longitude": -6.26
        }
      }
    }
  ]
}
```

Captured by running the server's real MCP handler against the test suite's in-memory Microsoft Graph (`tests/conftest.py`), not a live tenant: the values are test data, the shape is the server's.

## 🔄 End-to-end flow

![An AI client calls get_user_sign_ins; the module sends GET auditLogs/signIns filtered on the window and user id, newest first with top 1000, follows every nextLink, and returns the sign-ins or, when Graph refuses, an error with an AuditLog.Read.All hint on denial.](diagrams/sign-in-logs-user.svg)

Nothing is stored and no event is raised. Pages are read one after another until Graph sends no `@odata.nextLink`; a long window over a busy account means many calls and a large result. A denied read (`Authorization_RequestDenied`) adds the hint `(Ensure the application has AuditLog.Read.All permission)` to the error notification.

## 📚 Pages

- [Endpoints](endpoints.md) — the 1 tool, each with its mark, arguments, result, errors and an example call
- [Integration](integration.md) — the Microsoft services the feature calls and what happens when they fail

## 🔗 Related features

- [Audit logs](../audit-logs/index.md) — reach for it to see the changes a user made, not their sign-ins.
- [Users](../users/index.md) — reach for it to find the user id first.
