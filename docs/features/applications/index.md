# Applications

Manage app registrations: list and read them with the permissions they hold, create, change and delete them.

## 📖 Overview

- List app registrations, or read one with the app role assignments and delegated grants of its service principal.
- Create an app registration with its name, sign-in audience, tags, identifier URIs, web and API settings and required permissions.
- Change those same settings, and read the result back in the same call.
- Delete an app registration.

Called by an AI client on behalf of an Entra ID admin. Two tools are read-only, the create changes the tenant without destroying anything, the update and delete are destructive.

## 🏢 Business domain

App registrations (the application object). The tenant-local instance and its grants belong to [Service principals](../service-principals/index.md); which permissions to ask for, to [Permissions helper](../permissions-helper/index.md).

| Term | Meaning | In the code |
|---|---|---|
| **Application (app registration)** | The global definition of an app | `/applications` |
| **Object id** | The application object's own id — what `app_id` means in these tools | `id` |
| **Application (client) id** | The app's id shared with its service principals | `appId` |
| **App role assignment** | An application permission granted to the app's service principal | `appRoleAssignments` |
| **OAuth2 permission grant** | A delegated permission consented for the app | `oauth2PermissionGrants` |

| Rule | Enforced by | A caller who breaks it gets |
|---|---|---|
| Only 7 settings are sent on create and update; other keys are ignored | `applications.create_application`, `applications.update_application` | — |
| Grants that cannot be read do not fail the read | the `try` blocks in `applications.get_application_by_id` | Empty lists and a server log warning |

## 🚀 Quick Start

No token on the call: the server signs in itself, as its [configuration](../../configuration.md) says, on the first Microsoft Graph call. Set the server up in your AI client first ([Quick Start](../../index.md#-quick-start)); then ask the client in plain words, or send the call yourself:

```json
{"jsonrpc": "2.0", "id": 1, "method": "tools/call", "params": {"name": "get_application_by_id", "arguments": {"app_id": "payroll-sync-object"}}}
```

The result's `structuredContent` — list and optional results arrive under `result`:

```json
{
  "result": {
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
}
```

Captured by running the server's real MCP handler against the test suite's in-memory Microsoft Graph (`tests/conftest.py`), not a live tenant: the values are test data, the shape is the server's.

## 🔄 End-to-end flow

![An AI client calls update_application; the module PATCHes the application in Graph, then re-reads it, finds its service principal by appId, reads that principal's app role assignments and delegated grants, and returns the updated application; when Graph refuses, the client gets an error notification and the error result.](diagrams/applications-update.svg)

`update_application` stands for the feature's writes: one `PATCH`, then the full read of `get_application_by_id`. Nothing is stored and no ETag is sent, so the last write wins. When the read-back of grants fails, the update has still happened and the grants come back empty.

## 📚 Pages

- [Endpoints](endpoints.md) — the 5 tools, each with its mark, arguments, result, errors and an example call
- [Integration](integration.md) — the Microsoft services the feature calls and what happens when they fail

## 🔗 Related features

- [Service principals](../service-principals/index.md) — reach for it to manage the app's instance in this tenant.
- [Permissions helper](../permissions-helper/index.md) — reach for it to choose the permissions to put in `requiredResourceAccess`.

## ❓ Open questions

| Question | Why it matters | Checked | Who can answer |
|---|---|---|---|
| Do `web`, `api` and `requiredResourceAccess`, given as plain JSON objects, reach Graph intact? | If the SDK does not serialise a plain object in those fields, creating or updating an app with them fails or drops them. | `applications.create_application` assigns the raw values to the SDK model's fields; `test_an_application_is_created_with_every_field` checks only that the fields are set on the request body, not what is sent. | Maintainer (baoduy), by a run against a tenant |
