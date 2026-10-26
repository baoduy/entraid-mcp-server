# MFA — Endpoints

Part of [MFA](index.md).

Each endpoint is an MCP tool. An AI client calls it with a JSON-RPC `tools/call` message over the server's standard input; the server answers on standard output.

A tool that returns a list, or an object that can be `null`, sends it under `result` in `structuredContent` (`{"result": …}`), as its output schema declares; a tool that always returns an object sends that object itself. `content[0].text` carries the value itself, without the wrapper, as JSON text.

## 🔌 Endpoints

| Tool | Mark | Purpose | Auth |
|---|---|---|---|
| [`get_user_mfa_status`](#get_user_mfa_status) | read-only | Report which authentication methods a user has registered, and whether that counts as MFA. | signed-in identity |
| [`get_group_mfa_status`](#get_group_mfa_status) | read-only | Report MFA status for every user who is a direct member of a group. | signed-in identity |

### `get_user_mfa_status`

Report which authentication methods a user has registered, and whether that counts as MFA.

- **Mark:** read-only
- **Auth:** the identity the server signed in as ([configuration](../../configuration.md)); Microsoft Graph checks its permissions, the server checks none
- **Microsoft Graph:** `GET /users/{user_id}?$select=companyName,mail`, then `GET /users/{user_id}/authentication/methods`.
- **Request:**

| Field | Type | Required | Default | Rules |
|---|---|---|---|---|
| `user_id` | string | yes | — | — |

- **Response:** `userPrincipalName` (the `user_id` you passed, echoed), `mail`, `companyName`, `mfaStatus` (`Enabled` or `Disabled`) and `methods`, eight booleans: `email`, `fido2`, `authenticatorApp`, `password`, `phone`, `softwareOath`, `temporaryAccessPass`, `windowsHelloForBusiness`. `null` with the warning `No MFA data found for user <user_id>` when Graph returns no method list.

  Example `structuredContent`, captured from the test suite's in-memory Graph:

```json
{
  "result": {
    "userPrincipalName": "aoife@contoso.com",
    "mail": "aoife@contoso.com",
    "companyName": "Contoso",
    "mfaStatus": "Enabled",
    "methods": {
      "email": false,
      "fido2": false,
      "authenticatorApp": false,
      "password": false,
      "phone": true,
      "softwareOath": false,
      "temporaryAccessPass": false,
      "windowsHelloForBusiness": false
    }
  }
}
```

- **Errors:**

| Result | Text | When |
|---|---|---|
| error result | `Error calling tool 'get_user_mfa_status': <reason>`; notification `Authentication error: <reason>` | Sign-in cannot be set up (`AuthenticationError`) |
| error result | `Error calling tool 'get_user_mfa_status': <reason>`; notification `Error fetching MFA status for <user_id>: <reason>` | Microsoft Graph or sign-in fails |

- **Example:**

```json
{"jsonrpc": "2.0", "id": 1, "method": "tools/call", "params": {"name": "get_user_mfa_status", "arguments": {"user_id": "aoife@contoso.com"}}}
```

### `get_group_mfa_status`

Report MFA status for every user who is a direct member of a group.

- **Mark:** read-only
- **Auth:** the identity the server signed in as ([configuration](../../configuration.md)); Microsoft Graph checks its permissions, the server checks none
- **Microsoft Graph:** `GET /groups/{group_id}/members`, every page; members that are not `#microsoft.graph.user` are skipped. Then the two `get_user_mfa_status` calls per user.
- **Request:**

| Field | Type | Required | Default | Rules |
|---|---|---|---|---|
| `group_id` | string | yes | — | — |

- **Response:** A list of `get_user_mfa_status` results, one per user whose status could be read. When some could not be read, one warning: `MFA status could not be read for <n> of <users> members`. An empty list comes with the warning `No MFA data found for group <group_id>`.

  Example `structuredContent`, captured from the test suite's in-memory Graph:

```json
{
  "result": [
    {
      "userPrincipalName": "aoife@contoso.com",
      "mail": "aoife@contoso.com",
      "companyName": "Contoso",
      "mfaStatus": "Enabled",
      "methods": {
        "email": false,
        "fido2": false,
        "authenticatorApp": false,
        "password": false,
        "phone": true,
        "softwareOath": false,
        "temporaryAccessPass": false,
        "windowsHelloForBusiness": false
      }
    }
  ]
}
```

- **Errors:**

| Result | Text | When |
|---|---|---|
| error result | `Error calling tool 'get_group_mfa_status': <reason>`; notification `Authentication error: <reason>` | Sign-in cannot be set up (`AuthenticationError`) |
| error result | `Error calling tool 'get_group_mfa_status': <reason>`; notification `Error fetching group MFA status for <group_id>: <reason>` | Microsoft Graph or sign-in fails |

- **Example:**

```json
{"jsonrpc": "2.0", "id": 1, "method": "tools/call", "params": {"name": "get_group_mfa_status", "arguments": {"group_id": "all-staff"}}}
```

## ⚠️ Errors & limits

A failed call reaches the AI client twice: one error notification (`<message>: <reason>`, sent through the MCP log channel) and the error result `Error calling tool '<tool>': <reason>` with `isError: true`. The reason is the text of the exception Microsoft Graph, the SDK or the server raised. The server sets no timeout and no retry of its own and keeps no state between calls. Results that report a known case — `already_exists`, `not_found`, a dynamic group's `status: error`, the permission helper's `status: error` — are normal results, not errors.

- `mfaStatus` is `Enabled` when any method other than password is registered — email included.
- `userPrincipalName` in the result is the `user_id` you passed, not a value read from Graph.
- `get_group_mfa_status` covers direct user members only: nested groups are not expanded; devices, service principals and groups are skipped.
- `get_group_mfa_status` makes two Graph calls per user member, one after another; a member whose status cannot be read is counted in one warning, and the report goes on.
