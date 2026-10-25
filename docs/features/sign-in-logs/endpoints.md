# Sign-in logs — Endpoints

Part of [Sign-in logs](index.md).

Each endpoint is an MCP tool. An AI client calls it with a JSON-RPC `tools/call` message over the server's standard input; the server answers on standard output.

## 🔌 Endpoints

| Tool | Mark | Purpose | Auth |
|---|---|---|---|
| [`get_user_sign_ins`](#get_user_sign_ins) | read-only | List a user's sign-ins over the last `days` days, newest first. | signed-in identity |

### `get_user_sign_ins`

List a user's sign-ins over the last `days` days, newest first.

- **Mark:** read-only
- **Auth:** the identity the server signed in as ([configuration](../../configuration.md)); Microsoft Graph checks its permissions, the server checks none — the tool's description names `AuditLog.Read.All`
- **Microsoft Graph:** `GET /auditLogs/signIns` with `$filter=createdDateTime ge <start> and createdDateTime le <end> and userId eq '<user_id>'`, `$orderby=createdDateTime desc`, `$top=1000`, `ConsistencyLevel: eventual`, every page.
- **Request:**

| Field | Type | Required | Default | Rules |
|---|---|---|---|---|
| `user_id` | string | yes | — | — |
| `days` | integer | no | `7` | Window is now (UTC) minus `days` to now |

- **Response:** Each sign-in has `id`, `createdDateTime`, user and app fields, `ipAddress`, `clientAppUsed`, `correlationId`, `isInteractive`, `resourceDisplayName`, `status`, `riskInformation`, and `deviceDetail` and `location` when Graph sends them. An empty list comes with the info message `No sign-in logs found for user <user_id> in the last <days> days.`

  Example `structuredContent`, captured from the test suite's in-memory Graph:

```json
[
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
    "…": "6 more fields"
  }
]
```

- **Errors:**

| Result | Text | When |
|---|---|---|
| error result | notification ends with ` (Ensure the application has AuditLog.Read.All permission)` | Graph answers `Authorization_RequestDenied`. The error result is unchanged |
| error result | notification `Authentication error: <reason>` | Sign-in cannot be set up (`AuthenticationError`) |
| error result | `Error calling tool 'get_user_sign_ins': <reason>`; notification `Error fetching sign-in logs for <user_id>: <reason>` | Microsoft Graph or sign-in fails |

- **Example:**

```json
{"jsonrpc": "2.0", "id": 1, "method": "tools/call", "params": {"name": "get_user_sign_ins", "arguments": {"user_id": "aoife@contoso.com"}}}
```

## ⚠️ Errors & limits

A failed call reaches the AI client twice: one error notification (`<message>: <reason>`, sent through the MCP log channel) and the error result `Error calling tool '<tool>': <reason>` with `isError: true`. The reason is the text of the exception Microsoft Graph, the SDK or the server raised. The server sets no timeout and no retry of its own and keeps no state between calls. Results that report a known case — `already_exists`, `not_found`, a dynamic group's `status: error`, the permission helper's `status: error` — are normal results, not errors.

- The window is computed in UTC at call time; there is no upper limit on `days` in the server.
- Every page is read: a busy user over many days means many Graph calls and a large result.
- Graph's sign-in logs need the `AuditLog.Read.All` permission; a denied read adds a hint to the notification.
