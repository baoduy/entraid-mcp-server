# Password management — Endpoints

Part of [Password management](index.md).

Each endpoint is an MCP tool. An AI client calls it with a JSON-RPC `tools/call` message over the server's standard input; the server answers on standard output.

## 🔌 Endpoints

| Tool | Mark | Purpose | Auth |
|---|---|---|---|
| [`list_user_password_methods`](#list_user_password_methods) | read-only | List a user's password authentication methods. | signed-in identity |
| [`get_user_password_method`](#get_user_password_method) | read-only | Get one password authentication method of a user. | signed-in identity |
| [`reset_user_password_direct`](#reset_user_password_direct) | destructive | Set a user's password to a value you give or one the server generates. | signed-in identity |

### `list_user_password_methods`

List a user's password authentication methods.

- **Mark:** read-only
- **Auth:** the identity the server signed in as ([configuration](../../configuration.md)); Microsoft Graph checks its permissions, the server checks none
- **Microsoft Graph:** `GET /users/{user_id}/authentication/passwordMethods`.
- **Request:**

| Field | Type | Required | Default | Rules |
|---|---|---|---|---|
| `user_id` | string | yes | — | — |

- **Response:** Each method has `id` and `createdDateTime`.

  Example `structuredContent`, captured from the test suite's in-memory Graph:

```json
[
  {
    "id": "password-1",
    "createdDateTime": "2026-10-20T07:45:00+00:00"
  }
]
```

- **Errors:**

| Result | Text | When |
|---|---|---|
| error result | `Error calling tool 'list_user_password_methods': <reason>`; notification `Error listing password methods for user <user_id>: <reason>` | Microsoft Graph or sign-in fails |

- **Example:**

```json
{"jsonrpc": "2.0", "id": 1, "method": "tools/call", "params": {"name": "list_user_password_methods", "arguments": {"user_id": "aoife@contoso.com"}}}
```

### `get_user_password_method`

Get one password authentication method of a user.

- **Mark:** read-only
- **Auth:** the identity the server signed in as ([configuration](../../configuration.md)); Microsoft Graph checks its permissions, the server checks none
- **Microsoft Graph:** `GET /users/{user_id}/authentication/passwordMethods/{method_id}`.
- **Request:**

| Field | Type | Required | Default | Rules |
|---|---|---|---|---|
| `user_id` | string | yes | — | — |
| `method_id` | string | yes | — | — |

- **Response:** `id` and `createdDateTime`, or `null` with the warning `Password method <method_id> not found for user <user_id>` when Graph returns no object.

  Example `structuredContent`, captured from the test suite's in-memory Graph:

```json
{
  "id": "password-1",
  "createdDateTime": "2026-10-20T07:45:00+00:00"
}
```

- **Errors:**

| Result | Text | When |
|---|---|---|
| error result | `Error calling tool 'get_user_password_method': <reason>`; notification `Error getting password method <method_id> for user <user_id>: <reason>` | Microsoft Graph or sign-in fails |

- **Example:**

```json
{"jsonrpc": "2.0", "id": 1, "method": "tools/call", "params": {"name": "get_user_password_method", "arguments": {"user_id": "aoife@contoso.com", "method_id": "password-1"}}}
```

### `reset_user_password_direct`

Set a user's password to a value you give or one the server generates.

- **Mark:** destructive
- **Auth:** the identity the server signed in as ([configuration](../../configuration.md)); Microsoft Graph checks its permissions, the server checks none
- **Idempotency:** Each call sets the password again; with `generate_password` each call sets a new random one.
- **Concurrency:** none: no ETag or version is sent, so the last write wins.
- **Microsoft Graph:** `PATCH /users/{user_id}` with `passwordProfile` (`password`, `forceChangePasswordNextSignIn`).
- **Request:**

| Field | Type | Required | Default | Rules |
|---|---|---|---|---|
| `user_id` | string | yes | — | — |
| `password` | string or null | no | `null` | Used when `generate_password` is false; masked as `***` in any error text |
| `require_change_on_next_sign_in` | boolean | no | `true` | Sent as `passwordProfile.forceChangePasswordNextSignIn` |
| `generate_password` | boolean | no | `false` | When true, a generated password replaces any `password` given |
| `password_length` | integer | no | `12` | 8 to 256; used only when `generate_password` is true |

- **Response:** `status`, `userId`, `passwordResetRequired`, `message`, and `generated_password` when the server generated it.

  Example `structuredContent`, captured from the test suite's in-memory Graph:

```json
{
  "status": "success",
  "userId": "aoife@contoso.com",
  "passwordResetRequired": true,
  "message": "Password has been reset using the direct method."
}
```

- **Errors:**

| Result | Text | When |
|---|---|---|
| error result | `Password must be provided or generate_password must be set to True` | `generate_password` is false and no `password` is given. Before any Graph call |
| error result | `The password length must be 8 to 256 characters, got <n>` | `password_length` outside 8 to 256 with `generate_password`. Before any Graph call |
| error result | the Graph error, with the password replaced by `***` | Graph refuses the change. The original error is not chained, so no log or result carries the password |
| error result | `Error calling tool 'reset_user_password_direct': <reason>`; notification `Error directly resetting password for user <user_id>: <reason>` | Microsoft Graph or sign-in fails |

- **Example:**

```json
{"jsonrpc": "2.0", "id": 1, "method": "tools/call", "params": {"name": "reset_user_password_direct", "arguments": {"user_id": "aoife@contoso.com", "generate_password": true, "password_length": 16}}}
```

## ⚠️ Errors & limits

A failed call reaches the AI client twice: one error notification (`<message>: <reason>`, sent through the MCP log channel) and the error result `Error calling tool '<tool>': <reason>` with `isError: true`. The reason is the text of the exception Microsoft Graph, the SDK or the server raised. The server sets no timeout and no retry of its own and keeps no state between calls. Results that report a known case — `already_exists`, `not_found`, a dynamic group's `status: error`, the permission helper's `status: error` — are normal results, not errors.

- The generated password is returned in the result, so it reaches the AI client and its conversation.
- Error texts replace the password with `***` before any log line, notification or result carries them.
- Generated passwords hold at least one digit, upper-case letter, lower-case letter and symbol, drawn from the operating system's secure random source; the letters exclude `l` and `L`.
