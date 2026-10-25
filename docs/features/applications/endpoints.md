# Applications — Endpoints

Part of [Applications](index.md).

Each endpoint is an MCP tool. An AI client calls it with a JSON-RPC `tools/call` message over the server's standard input; the server answers on standard output.

## 🔌 Endpoints

| Tool | Mark | Purpose | Auth |
|---|---|---|---|
| [`list_applications`](#list_applications) | read-only | List app registrations, up to `limit`. | signed-in identity |
| [`get_application_by_id`](#get_application_by_id) | read-only | Get an app registration with the permissions its service principal holds. | signed-in identity |
| [`create_application`](#create_application) | changes the tenant, not destructive | Create an app registration. | signed-in identity |
| [`update_application`](#update_application) | destructive | Change an app registration. | signed-in identity |
| [`delete_application`](#delete_application) | destructive | Delete an app registration. | signed-in identity |

### `list_applications`

List app registrations, up to `limit`.

- **Mark:** read-only
- **Auth:** the identity the server signed in as ([configuration](../../configuration.md)); Microsoft Graph checks its permissions, the server checks none
- **Microsoft Graph:** `GET /applications` (Graph's default page size), following `@odata.nextLink` until `limit` items are held.
- **Request:**

| Field | Type | Required | Default | Rules |
|---|---|---|---|---|
| `limit` | integer | no | `100` | Pages are read until `limit` items are held |

- **Response:** Each application has `id`, `appId`, `displayName`, `createdDateTime`, `signInAudience`, `publisherDomain`, `tags`.

  Example `structuredContent`, captured from the test suite's in-memory Graph:

```json
[
  {
    "id": "payroll-sync-object",
    "appId": "payroll-sync-app",
    "displayName": "Payroll Sync",
    "createdDateTime": "2026-10-20T07:45:00+00:00",
    "signInAudience": "AzureADMyOrg",
    "publisherDomain": "contoso.com",
    "tags": [
      "payroll"
    ]
  }
]
```

- **Errors:**

| Result | Text | When |
|---|---|---|
| error result | `Error calling tool 'list_applications': <reason>`; notification `Error listing applications: <reason>` | Microsoft Graph or sign-in fails |

- **Example:**

```json
{"jsonrpc": "2.0", "id": 1, "method": "tools/call", "params": {"name": "list_applications", "arguments": {}}}
```

### `get_application_by_id`

Get an app registration with the permissions its service principal holds.

- **Mark:** read-only
- **Auth:** the identity the server signed in as ([configuration](../../configuration.md)); Microsoft Graph checks its permissions, the server checks none
- **Microsoft Graph:** `GET /applications/{app_id}`, then `GET /servicePrincipals?$filter=appId eq '<appId>'`, then that service principal's `appRoleAssignments` and `oauth2PermissionGrants`, every page.
- **Request:**

| Field | Type | Required | Default | Rules |
|---|---|---|---|---|
| `app_id` | string | yes | — | — |

- **Response:** The `list_applications` fields plus `appRoleAssignments` and `oauth2PermissionGrants`; both are empty when the app has no service principal or the grants cannot be read (logged as a warning). `null` with the warning `Application with ID <app_id> not found.` when Graph returns no object.

  Example `structuredContent`, captured from the test suite's in-memory Graph:

```json
{
  "id": "payroll-sync-object",
  "appId": "payroll-sync-app",
  "displayName": "Payroll Sync",
  "createdDateTime": "2026-10-20T07:45:00+00:00",
  "signInAudience": "AzureADMyOrg",
  "publisherDomain": "contoso.com",
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
| error result | `Error calling tool 'get_application_by_id': <reason>`; notification `Error fetching application <app_id>: <reason>` | Microsoft Graph or sign-in fails |

- **Example:**

```json
{"jsonrpc": "2.0", "id": 1, "method": "tools/call", "params": {"name": "get_application_by_id", "arguments": {"app_id": "payroll-sync-object"}}}
```

### `create_application`

Create an app registration.

- **Mark:** changes the tenant, not destructive
- **Auth:** the identity the server signed in as ([configuration](../../configuration.md)); Microsoft Graph checks its permissions, the server checks none
- **Idempotency:** Not idempotent: no duplicate check; each call creates another application.
- **Microsoft Graph:** `POST /applications` with the supported `app_data` keys.
- **Request:**

| Field | Type | Required | Default | Rules |
|---|---|---|---|---|
| `app_data` | object | yes | — | Keys sent: `displayName`, `signInAudience`, `tags`, `identifierUris`, `web`, `api`, `requiredResourceAccess`; others ignored |

- **Response:** The `list_applications` fields of the new application.

  Example `structuredContent`, captured from the test suite's in-memory Graph:

```json
{
  "id": "new-0001",
  "appId": null,
  "displayName": "Payroll Sync",
  "createdDateTime": null,
  "signInAudience": null,
  "publisherDomain": null,
  "tags": null
}
```

- **Errors:**

| Result | Text | When |
|---|---|---|
| error result | `Failed to create application` | Graph returns no object. After the POST |
| error result | `Error calling tool 'create_application': <reason>`; notification `Error creating application: <reason>` | Microsoft Graph or sign-in fails |

- **Example:**

```json
{"jsonrpc": "2.0", "id": 1, "method": "tools/call", "params": {"name": "create_application", "arguments": {"app_data": {"displayName": "Payroll Sync"}}}}
```

### `update_application`

Change an app registration.

- **Mark:** destructive
- **Auth:** the identity the server signed in as ([configuration](../../configuration.md)); Microsoft Graph checks its permissions, the server checks none
- **Idempotency:** Not idempotent: each call is sent to Graph again.
- **Concurrency:** none: no ETag or version is sent, so the last write wins.
- **Microsoft Graph:** `PATCH /applications/{app_id}` with the supported `app_data` keys, then the `get_application_by_id` reads.
- **Request:**

| Field | Type | Required | Default | Rules |
|---|---|---|---|---|
| `app_id` | string | yes | — | — |
| `app_data` | object | yes | — | Same 7 keys as `create_application`; others ignored |

- **Response:** The `get_application_by_id` result after the change.

  Example `structuredContent`, captured from the test suite's in-memory Graph:

```json
{
  "id": "payroll-sync-object",
  "appId": "payroll-sync-app",
  "displayName": "Payroll Sync",
  "createdDateTime": "2026-10-20T07:45:00+00:00",
  "signInAudience": "AzureADMyOrg",
  "publisherDomain": "contoso.com",
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
| error result | `Error calling tool 'update_application': <reason>`; notification `Error updating application <app_id>: <reason>` | Microsoft Graph or sign-in fails |

- **Example:**

```json
{"jsonrpc": "2.0", "id": 1, "method": "tools/call", "params": {"name": "update_application", "arguments": {"app_id": "payroll-sync-object", "app_data": {"displayName": "Payroll Sync"}}}}
```

### `delete_application`

Delete an app registration.

- **Mark:** destructive
- **Auth:** the identity the server signed in as ([configuration](../../configuration.md)); Microsoft Graph checks its permissions, the server checks none
- **Idempotency:** Not idempotent: a second call reaches Graph again and fails with Graph's error once the application is gone.
- **Microsoft Graph:** `DELETE /applications/{app_id}`.
- **Request:**

| Field | Type | Required | Default | Rules |
|---|---|---|---|---|
| `app_id` | string | yes | — | — |

- **Response:** `{"status": "success", "message": "Application <app_id> was deleted successfully"}`.

  Example `structuredContent`, captured from the test suite's in-memory Graph:

```json
{
  "status": "success",
  "message": "Application payroll-sync-object was deleted successfully"
}
```

- **Errors:**

| Result | Text | When |
|---|---|---|
| error result | `Error calling tool 'delete_application': <reason>`; notification `Error deleting application <app_id>: <reason>` | Microsoft Graph or sign-in fails |

- **Example:**

```json
{"jsonrpc": "2.0", "id": 1, "method": "tools/call", "params": {"name": "delete_application", "arguments": {"app_id": "payroll-sync-object"}}}
```

## ⚠️ Errors & limits

A failed call reaches the AI client twice: one error notification (`<message>: <reason>`, sent through the MCP log channel) and the error result `Error calling tool '<tool>': <reason>` with `isError: true`. The reason is the text of the exception Microsoft Graph, the SDK or the server raised. The server sets no timeout and no retry of its own and keeps no state between calls. Results that report a known case — `already_exists`, `not_found`, a dynamic group's `status: error`, the permission helper's `status: error` — are normal results, not errors.

- `create_application` has no duplicate check: each call creates another application.
- `get_application_by_id` makes up to four Graph calls plus paging; unreadable grants give empty lists and a warning in the server log, not an error.
- `list_applications` uses Graph's default page size and stops once `limit` items are held.
