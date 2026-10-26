# Service principals — Endpoints

Part of [Service principals](index.md).

Each endpoint is an MCP tool. An AI client calls it with a JSON-RPC `tools/call` message over the server's standard input; the server answers on standard output.

A tool that returns a list, or an object that can be `null`, sends it under `result` in `structuredContent` (`{"result": …}`), as its output schema declares; a tool that always returns an object sends that object itself. `content[0].text` carries the value itself, without the wrapper, as JSON text.

## 🔌 Endpoints

| Tool | Mark | Purpose | Auth |
|---|---|---|---|
| [`list_service_principals`](#list_service_principals) | read-only | List service principals, up to `limit`. | signed-in identity |
| [`get_service_principal_by_id`](#get_service_principal_by_id) | read-only | Get a service principal with its app role assignments and delegated grants. | signed-in identity |
| [`create_service_principal`](#create_service_principal) | changes the tenant, not destructive | Create a service principal for an application id. | signed-in identity |
| [`update_service_principal`](#update_service_principal) | destructive | Change a service principal. | signed-in identity |
| [`delete_service_principal`](#delete_service_principal) | destructive | Delete a service principal. | signed-in identity |

### `list_service_principals`

List service principals, up to `limit`.

- **Mark:** read-only
- **Auth:** the identity the server signed in as ([configuration](../../configuration.md)); Microsoft Graph checks its permissions, the server checks none
- **Microsoft Graph:** `GET /servicePrincipals` (Graph's default page size), following `@odata.nextLink` until `limit` items are held.
- **Request:**

| Field | Type | Required | Default | Rules |
|---|---|---|---|---|
| `limit` | integer | no | `100` | Pages are read until `limit` items are held |

- **Response:** Each service principal has `id`, `appId`, `displayName`, `createdDateTime`, `accountEnabled`, `appOwnerOrganizationId`, `tags`.

  Example `structuredContent`, captured from the test suite's in-memory Graph:

```json
{
  "result": [
    {
      "id": "payroll-sync-sp",
      "appId": "payroll-sync-app",
      "displayName": "Payroll Sync",
      "createdDateTime": null,
      "accountEnabled": true,
      "appOwnerOrganizationId": "contoso-tenant",
      "tags": [
        "payroll"
      ]
    }
  ]
}
```

- **Errors:**

| Result | Text | When |
|---|---|---|
| error result | `Error calling tool 'list_service_principals': <reason>`; notification `Error listing service principals: <reason>` | Microsoft Graph or sign-in fails |

- **Example:**

```json
{"jsonrpc": "2.0", "id": 1, "method": "tools/call", "params": {"name": "list_service_principals", "arguments": {}}}
```

### `get_service_principal_by_id`

Get a service principal with its app role assignments and delegated grants.

- **Mark:** read-only
- **Auth:** the identity the server signed in as ([configuration](../../configuration.md)); Microsoft Graph checks its permissions, the server checks none
- **Microsoft Graph:** `GET /servicePrincipals/{sp_id}`, then `appRoleAssignments` and `oauth2PermissionGrants`, every page.
- **Request:**

| Field | Type | Required | Default | Rules |
|---|---|---|---|---|
| `sp_id` | string | yes | — | — |

- **Response:** The `list_service_principals` fields plus `appRoleAssignments` and `oauth2PermissionGrants` (empty when they cannot be read; logged as a warning). `null` with the warning `Service principal with ID <sp_id> not found.` when Graph returns no object.

  Example `structuredContent`, captured from the test suite's in-memory Graph:

```json
{
  "result": {
    "id": "payroll-sync-sp",
    "appId": "payroll-sync-app",
    "displayName": "Payroll Sync",
    "createdDateTime": null,
    "accountEnabled": true,
    "appOwnerOrganizationId": "contoso-tenant",
    "tags": [
      "payroll"
    ],
    "appRoleAssignments": [
      {
        "id": "assignment-1",
        "createdDateTime": "2026-10-20T07:45:00Z",
        "appRoleId": null,
        "principalDisplayName": null,
        "principalId": null,
        "principalType": null,
        "resourceDisplayName": null,
        "resourceId": null
      }
    ],
    "oauth2PermissionGrants": [
      {
        "id": "grant-1",
        "clientId": null,
        "consentType": null,
        "principalId": null,
        "resourceId": null,
        "scope": "User.Read"
      }
    ]
  }
}
```

- **Errors:**

| Result | Text | When |
|---|---|---|
| error result | `Error calling tool 'get_service_principal_by_id': <reason>`; notification `Error fetching service principal <sp_id>: <reason>` | Microsoft Graph or sign-in fails |

- **Example:**

```json
{"jsonrpc": "2.0", "id": 1, "method": "tools/call", "params": {"name": "get_service_principal_by_id", "arguments": {"sp_id": "payroll-sync-sp"}}}
```

### `create_service_principal`

Create a service principal for an application id.

- **Mark:** changes the tenant, not destructive
- **Auth:** the identity the server signed in as ([configuration](../../configuration.md)); Microsoft Graph checks its permissions, the server checks none
- **Idempotency:** Not idempotent: no duplicate check in the server; Graph's answer decides a second create for the same `appId`.
- **Microsoft Graph:** `POST /servicePrincipals` with the supported `sp_data` keys.
- **Request:**

| Field | Type | Required | Default | Rules |
|---|---|---|---|---|
| `sp_data` | object | yes | — | Keys sent: `appId`, `accountEnabled`, `tags`, `appRoleAssignmentRequired`, `displayName`; others ignored |

- **Response:** The `list_service_principals` fields of the new service principal.

  Example `structuredContent`, captured from the test suite's in-memory Graph:

```json
{
  "id": "new-0001",
  "appId": "payroll-sync-app",
  "displayName": null,
  "createdDateTime": null,
  "accountEnabled": null,
  "appOwnerOrganizationId": null,
  "tags": null
}
```

- **Errors:**

| Result | Text | When |
|---|---|---|
| error result | `Error calling tool 'create_service_principal': Failed to create service principal`; notification `Error creating service principal: Failed to create service principal` | Graph returns no object. After the POST |
| error result | `Error calling tool 'create_service_principal': <reason>`; notification `Error creating service principal: <reason>` | Microsoft Graph or sign-in fails |

- **Example:**

```json
{"jsonrpc": "2.0", "id": 1, "method": "tools/call", "params": {"name": "create_service_principal", "arguments": {"sp_data": {"appId": "payroll-sync-app"}}}}
```

### `update_service_principal`

Change a service principal.

- **Mark:** destructive
- **Auth:** the identity the server signed in as ([configuration](../../configuration.md)); Microsoft Graph checks its permissions, the server checks none
- **Idempotency:** Safe to retry: the same `PATCH` is sent again and leaves the same state; nothing new is created.
- **Concurrency:** none: no ETag or version is sent, so the last write wins.
- **Microsoft Graph:** `PATCH /servicePrincipals/{sp_id}` with the supported `sp_data` keys, then the `get_service_principal_by_id` reads.
- **Request:**

| Field | Type | Required | Default | Rules |
|---|---|---|---|---|
| `sp_id` | string | yes | — | — |
| `sp_data` | object | yes | — | Keys sent: `accountEnabled`, `tags`, `appRoleAssignmentRequired`, `displayName`; others ignored |

- **Response:** The `get_service_principal_by_id` result after the change.

  Example `structuredContent`, captured from the test suite's in-memory Graph:

```json
{
  "id": "payroll-sync-sp",
  "appId": "payroll-sync-app",
  "displayName": "Payroll Sync",
  "createdDateTime": null,
  "accountEnabled": true,
  "appOwnerOrganizationId": "contoso-tenant",
  "tags": [
    "payroll"
  ],
  "appRoleAssignments": [
    {
      "id": "assignment-1",
      "createdDateTime": "2026-10-20T07:45:00Z",
      "appRoleId": null,
      "principalDisplayName": null,
      "principalId": null,
      "principalType": null,
      "resourceDisplayName": null,
      "resourceId": null
    }
  ],
  "oauth2PermissionGrants": [
    {
      "id": "grant-1",
      "clientId": null,
      "consentType": null,
      "principalId": null,
      "resourceId": null,
      "scope": "User.Read"
    }
  ]
}
```

- **Errors:**

| Result | Text | When |
|---|---|---|
| error result | `Error calling tool 'update_service_principal': <reason>`; notification `Error updating service principal <sp_id>: <reason>` | Microsoft Graph or sign-in fails |

- **Example:**

```json
{"jsonrpc": "2.0", "id": 1, "method": "tools/call", "params": {"name": "update_service_principal", "arguments": {"sp_id": "payroll-sync-sp", "sp_data": {"accountEnabled": true}}}}
```

### `delete_service_principal`

Delete a service principal.

- **Mark:** destructive
- **Auth:** the identity the server signed in as ([configuration](../../configuration.md)); Microsoft Graph checks its permissions, the server checks none
- **Idempotency:** Not idempotent: a second call reaches Graph again and fails with Graph's error once the service principal is gone.
- **Concurrency:** none: no ETag or version is sent; the delete applies to the service principal as it is when Graph receives it.
- **Microsoft Graph:** `DELETE /servicePrincipals/{sp_id}`.
- **Request:**

| Field | Type | Required | Default | Rules |
|---|---|---|---|---|
| `sp_id` | string | yes | — | — |

- **Response:** `{"status": "success", "message": "Service principal <sp_id> was deleted successfully"}`.

  Example `structuredContent`, captured from the test suite's in-memory Graph:

```json
{
  "status": "success",
  "message": "Service principal payroll-sync-sp was deleted successfully"
}
```

- **Errors:**

| Result | Text | When |
|---|---|---|
| error result | `Error calling tool 'delete_service_principal': <reason>`; notification `Error deleting service principal <sp_id>: <reason>` | Microsoft Graph or sign-in fails |

- **Example:**

```json
{"jsonrpc": "2.0", "id": 1, "method": "tools/call", "params": {"name": "delete_service_principal", "arguments": {"sp_id": "payroll-sync-sp"}}}
```

## ⚠️ Errors & limits

A failed call reaches the AI client twice: one error notification (`<message>: <reason>`, sent through the MCP log channel) and the error result `Error calling tool '<tool>': <reason>` with `isError: true`. The reason is the text of the exception Microsoft Graph, the SDK or the server raised. The server sets no timeout and no retry of its own and keeps no state between calls. Results that report a known case — `already_exists`, `not_found`, a dynamic group's `status: error`, the permission helper's `status: error` — are normal results, not errors.

- `get_service_principal_by_id` makes three Graph calls plus paging; unreadable grants give empty lists and a warning in the server log.
- `list_service_principals` uses Graph's default page size and stops once `limit` items are held.
