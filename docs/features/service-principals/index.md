# Service principals

Manage service principals — an app's identity in this tenant: list and read them with their grants, create, change and delete them.

## 📖 Overview

- List service principals, or read one with its app role assignments and delegated permission grants.
- Create a service principal for an application id.
- Enable or disable it, require role assignment, change its name or tags, and read the result back in the same call.
- Delete a service principal.

Called by an AI client on behalf of an Entra ID admin. Two tools are read-only, the create changes the tenant without destroying anything, the update and delete are destructive.

## 🏢 Business domain

Enterprise apps (service principals) in this tenant. The global app definition belongs to [Applications](../applications/index.md).

| Term | Meaning | In the code |
|---|---|---|
| **Service principal** | An app's identity and permissions in this tenant | `/servicePrincipals` |
| **App role assignment** | An application permission granted to the service principal | `appRoleAssignments` |
| **OAuth2 permission grant** | A delegated permission consented for it | `oauth2PermissionGrants` |
| **Account enabled** | Whether the app can sign in | `accountEnabled` |

| Rule | Enforced by | A caller who breaks it gets |
|---|---|---|
| Only 5 settings are sent on create (4 on update); other keys are ignored | `service_principals.create_service_principal`, `update_service_principal` | — |
| Grants that cannot be read do not fail the read | the `try` blocks in `service_principals.get_service_principal_by_id` | Empty lists and a server log warning |

## 🚀 Quick Start

No token on the call: the server signs in itself, as its [configuration](../../configuration.md) says, on the first Microsoft Graph call. Set the server up in your AI client first ([Quick Start](../../index.md#-quick-start)); then ask the client in plain words, or send the call yourself:

```json
{"jsonrpc": "2.0", "id": 1, "method": "tools/call", "params": {"name": "get_service_principal_by_id", "arguments": {"sp_id": "payroll-sync-sp"}}}
```

The result's `structuredContent`:

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

Captured by running the server's real MCP handler against the test suite's in-memory Microsoft Graph (`tests/conftest.py`), not a live tenant: the values are test data, the shape is the server's.

## 🔄 End-to-end flow

![An AI client calls update_service_principal; the module PATCHes the service principal in Graph, then re-reads it with its app role assignments and delegated grants, and returns the updated service principal.](diagrams/service-principals-update.svg)

`update_service_principal` stands for the feature's writes: one `PATCH`, then the full read of `get_service_principal_by_id`. Nothing is stored and no ETag is sent, so the last write wins.

## 📚 Pages

- [Endpoints](endpoints.md) — the 5 tools, each with its mark, arguments, result, errors and an example call
- [Integration](integration.md) — the Microsoft services the feature calls and what happens when they fail

## 🔗 Related features

- [Applications](../applications/index.md) — reach for it to manage the app registration behind it.
- [Permissions helper](../permissions-helper/index.md) — reach for it to look up what a granted permission allows.
