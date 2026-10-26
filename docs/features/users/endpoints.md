# Users — Endpoints

Part of [Users](index.md).

Each endpoint is an MCP tool. An AI client calls it with a JSON-RPC `tools/call` message over the server's standard input; the server answers on standard output.

A tool that returns a list, or an object that can be `null`, sends it under `result` in `structuredContent` (`{"result": …}`), as its output schema declares; a tool that always returns an object sends that object itself. `content[0].text` carries the value itself, without the wrapper, as JSON text.

## 🔌 Endpoints

| Tool | Mark | Purpose | Auth |
|---|---|---|---|
| [`search_users`](#search_users) | read-only | Search users by name or email. | signed-in identity |
| [`get_user_by_id`](#get_user_by_id) | read-only | Get one user's profile by object id or user principal name. | signed-in identity |
| [`get_privileged_users`](#get_privileged_users) | read-only | List every user who holds an activated directory role, with the role names. | signed-in identity |
| [`get_user_groups`](#get_user_groups) | read-only | List every group a user belongs to, directly or through nested groups. | signed-in identity |
| [`get_user_roles`](#get_user_roles) | read-only | List the directory roles a user holds directly. | signed-in identity |

### `search_users`

Search users by name or email.

- **Mark:** read-only
- **Auth:** the identity the server signed in as ([configuration](../../configuration.md)); Microsoft Graph checks its permissions, the server checks none
- **Microsoft Graph:** `GET /users` with `$search` over six name and mail fields, header `ConsistencyLevel: eventual`, following `@odata.nextLink` until `limit` users are held.
- **Request:**

| Field | Type | Required | Default | Rules |
|---|---|---|---|---|
| `query` | string | yes | — | Matched against displayName, mail, userPrincipalName, givenName, surname and otherMails with `$search`; `\` and `"` are escaped |
| `limit` | integer | no | `10` | Sent as `$top`; next pages are read until `limit` users are held |

- **Response:** Each user has `id`, `displayName`, `mail`, `userPrincipalName`, `givenName`, `surname`, `jobTitle`, `officeLocation`, `businessPhones`, `mobilePhone`.

  Example `structuredContent`, captured through an MCP client call against the test suite's in-memory Graph:

```json
{
  "result": [
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
      "mobilePhone": "+353 87 555 0100"
    }
  ]
}
```

- **Errors:**

| Result | Text | When |
|---|---|---|
| error result | `Error calling tool 'search_users': <reason>`; notification `Authentication error: <reason>` | Sign-in cannot be set up (`AuthenticationError`) |
| error result | `Error calling tool 'search_users': <reason>`; notification `Error searching users: <reason>` | Microsoft Graph or sign-in fails |

- **Example:**

```json
{"jsonrpc": "2.0", "id": 1, "method": "tools/call", "params": {"name": "search_users", "arguments": {"query": "Aoife"}}}
```

### `get_user_by_id`

Get one user's profile by object id or user principal name.

- **Mark:** read-only
- **Auth:** the identity the server signed in as ([configuration](../../configuration.md)); Microsoft Graph checks its permissions, the server checks none
- **Microsoft Graph:** `GET /users/{user_id}`.
- **Request:**

| Field | Type | Required | Default | Rules |
|---|---|---|---|---|
| `user_id` | string | yes | — | — |

- **Response:** The same 10 fields as `search_users`, or `null` with the warning `User with ID <user_id> not found.` when Graph returns no object.

  Example `structuredContent`, captured from the test suite's in-memory Graph:

```json
{
  "result": {
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
    "mobilePhone": "+353 87 555 0100"
  }
}
```

- **Errors:**

| Result | Text | When |
|---|---|---|
| error result | `Error calling tool 'get_user_by_id': <reason>`; notification `Authentication error: <reason>` | Sign-in cannot be set up (`AuthenticationError`) |
| error result | `Error calling tool 'get_user_by_id': <reason>`; notification `Error fetching user <user_id>: <reason>` | Microsoft Graph or sign-in fails |

- **Example:**

```json
{"jsonrpc": "2.0", "id": 1, "method": "tools/call", "params": {"name": "get_user_by_id", "arguments": {"user_id": "aoife@contoso.com"}}}
```

### `get_privileged_users`

List every user who holds an activated directory role, with the role names.

- **Mark:** read-only
- **Auth:** the identity the server signed in as ([configuration](../../configuration.md)); Microsoft Graph checks its permissions, the server checks none
- **Microsoft Graph:** `GET /directoryRoles` (one page), then `GET /directoryRoles/{id}/members` for each role, every page. Only `#microsoft.graph.user` members are kept; a user in several roles appears once.
- **Request:**

  No arguments.

- **Response:** Each user has the 10 profile fields plus `roles`, a list of role display names.

  Example `structuredContent`, captured from the test suite's in-memory Graph:

```json
{
  "result": [
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
      "roles": [
        "Global Administrator"
      ]
    }
  ]
}
```

- **Errors:**

| Result | Text | When |
|---|---|---|
| error result | `Error calling tool 'get_privileged_users': <reason>`; notification `Error fetching privileged users: <reason>` | Microsoft Graph or sign-in fails |

- **Example:**

```json
{"jsonrpc": "2.0", "id": 1, "method": "tools/call", "params": {"name": "get_privileged_users", "arguments": {}}}
```

### `get_user_groups`

List every group a user belongs to, directly or through nested groups.

- **Mark:** read-only
- **Auth:** the identity the server signed in as ([configuration](../../configuration.md)); Microsoft Graph checks its permissions, the server checks none
- **Microsoft Graph:** `GET /users/{user_id}/transitiveMemberOf`, every page, then `GET /groups/{id}` for each `#microsoft.graph.group` entry.
- **Request:**

| Field | Type | Required | Default | Rules |
|---|---|---|---|---|
| `user_id` | string | yes | — | — |

- **Response:** Each group has `id`, `displayName`, `mail`, `groupTypes`, `description`.

  Example `structuredContent`, captured from the test suite's in-memory Graph:

```json
{
  "result": [
    {
      "id": "all-staff",
      "displayName": "All Staff",
      "mail": "allstaff@contoso.com",
      "groupTypes": [
        "Unified"
      ],
      "description": "Everyone at Contoso"
    }
  ]
}
```

- **Errors:**

| Result | Text | When |
|---|---|---|
| error result | `Error calling tool 'get_user_groups': <reason>`; notification `Error fetching groups for user <user_id>: <reason>` | Microsoft Graph or sign-in fails |

- **Example:**

```json
{"jsonrpc": "2.0", "id": 1, "method": "tools/call", "params": {"name": "get_user_groups", "arguments": {"user_id": "aoife@contoso.com"}}}
```

### `get_user_roles`

List the directory roles a user holds directly.

- **Mark:** read-only
- **Auth:** the identity the server signed in as ([configuration](../../configuration.md)); Microsoft Graph checks its permissions, the server checks none
- **Microsoft Graph:** `GET /users/{user_id}/memberOf`, every page, then `GET /directoryRoles/{id}` for each `#microsoft.graph.directoryRole` entry.
- **Request:**

| Field | Type | Required | Default | Rules |
|---|---|---|---|---|
| `user_id` | string | yes | — | — |

- **Response:** Each role has `id`, `displayName`, `description`, `roleTemplateId`.

  Example `structuredContent`, captured from the test suite's in-memory Graph:

```json
{
  "result": [
    {
      "id": "role-1",
      "displayName": "Global Administrator",
      "description": "Can manage all aspects of Entra ID",
      "roleTemplateId": "62e90394-69f5-4237-9190-012177145e10"
    }
  ]
}
```

- **Errors:**

| Result | Text | When |
|---|---|---|
| error result | `Error calling tool 'get_user_roles': <reason>`; notification `Error fetching roles for user <user_id>: <reason>` | Microsoft Graph or sign-in fails |

- **Example:**

```json
{"jsonrpc": "2.0", "id": 1, "method": "tools/call", "params": {"name": "get_user_roles", "arguments": {"user_id": "aoife@contoso.com"}}}
```

## ⚠️ Errors & limits

A failed call reaches the AI client twice: one error notification (`<message>: <reason>`, sent through the MCP log channel) and the error result `Error calling tool '<tool>': <reason>` with `isError: true`. The reason is the text of the exception Microsoft Graph, the SDK or the server raised. The server sets no timeout and no retry of its own and keeps no state between calls. Results that report a known case — `already_exists`, `not_found`, a dynamic group's `status: error`, the permission helper's `status: error` — are normal results, not errors.

- `search_users` sends `ConsistencyLevel: eventual`: a user created moments ago may not be found yet.
- `get_privileged_users` reads activated directory roles only (`/directoryRoles`, first page); role-eligible assignments are not read, and only user members are listed.
- `get_user_groups` makes one extra Graph call per group the user belongs to; `get_user_roles` one per directory role.
- Paging: `search_users` stops once `limit` users are held; the other user tools read every page.
