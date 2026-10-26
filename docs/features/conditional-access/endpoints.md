# Conditional access — Endpoints

Part of [Conditional access](index.md).

Each endpoint is an MCP tool. An AI client calls it with a JSON-RPC `tools/call` message over the server's standard input; the server answers on standard output.

A tool that returns a list, or an object that can be `null`, sends it under `result` in `structuredContent` (`{"result": …}`), as its output schema declares; a tool that always returns an object sends that object itself. `content[0].text` carries the value itself, without the wrapper, as JSON text.

## 🔌 Endpoints

| Tool | Mark | Purpose | Auth |
|---|---|---|---|
| [`get_conditional_access_policies`](#get_conditional_access_policies) | read-only | List the tenant's conditional access policies as flat rows. | signed-in identity |
| [`get_conditional_access_policy_by_id`](#get_conditional_access_policy_by_id) | read-only | Get one conditional access policy as a flat row. | signed-in identity |

### `get_conditional_access_policies`

List the tenant's conditional access policies as flat rows.

- **Mark:** read-only
- **Auth:** the identity the server signed in as ([configuration](../../configuration.md)); Microsoft Graph checks its permissions, the server checks none
- **Microsoft Graph:** `GET /identity/conditionalAccess/policies`, following `@odata.nextLink` until every page is read, then `GET /groups/{id}` for each included or excluded group id.
- **Request:**

  No arguments.

- **Response:** Each policy is one flat object: `id`, `displayName`, `state`, `createdDateTime`, `modifiedDateTime`, the condition columns (`Users_Include` … `Client_Applications`), the grant columns prefixed `Grant_`, and the session-control columns. List values are joined into one string with `; `.

  Example `structuredContent` (abridged: objects over 14 fields show their first 10), captured from the test suite's in-memory Graph:

```json
{
  "result": [
    {
      "id": "policy-1",
      "displayName": "Require MFA for All Staff",
      "state": "enabled",
      "createdDateTime": "2026-10-20T07:45:00+00:00",
      "modifiedDateTime": "2026-10-20T07:45:00+00:00",
      "Users_Include": "",
      "Users_Exclude": "break-glass",
      "Groups_Include": "all-staff",
      "Groups_Include_Names": "All Staff (all-staff)",
      "Groups_Exclude": "",
      "…": "38 more fields"
    }
  ]
}
```

- **Errors:**

| Result | Text | When |
|---|---|---|
| error result | `Error calling tool 'get_conditional_access_policies': <reason>`; notification `Error fetching conditional access policies: <reason>` | Microsoft Graph or sign-in fails |

- **Example:**

```json
{"jsonrpc": "2.0", "id": 1, "method": "tools/call", "params": {"name": "get_conditional_access_policies", "arguments": {}}}
```

### `get_conditional_access_policy_by_id`

Get one conditional access policy as a flat row.

- **Mark:** read-only
- **Auth:** the identity the server signed in as ([configuration](../../configuration.md)); Microsoft Graph checks its permissions, the server checks none
- **Microsoft Graph:** `GET /identity/conditionalAccess/policies/{policy_id}`, then `GET /groups/{id}` for each included or excluded group id.
- **Request:**

| Field | Type | Required | Default | Rules |
|---|---|---|---|---|
| `policy_id` | string | yes | — | — |

- **Response:** One flat policy object (same columns as `get_conditional_access_policies`), or `{}` with the warning `Policy with ID <policy_id> not found.` when Graph returns no object.

  Example `structuredContent` (abridged: objects over 14 fields show their first 10), captured from the test suite's in-memory Graph:

```json
{
  "id": "policy-1",
  "displayName": "Require MFA for All Staff",
  "state": "enabled",
  "createdDateTime": "2026-10-20T07:45:00+00:00",
  "modifiedDateTime": "2026-10-20T07:45:00+00:00",
  "Users_Include": "",
  "Users_Exclude": "break-glass",
  "Groups_Include": "all-staff",
  "Groups_Include_Names": "All Staff (all-staff)",
  "Groups_Exclude": "",
  "…": "38 more fields"
}
```

- **Errors:**

| Result | Text | When |
|---|---|---|
| error result | `Error calling tool 'get_conditional_access_policy_by_id': <reason>`; notification `Error fetching conditional access policy <policy_id>: <reason>` | Microsoft Graph or sign-in fails |

- **Example:**

```json
{"jsonrpc": "2.0", "id": 1, "method": "tools/call", "params": {"name": "get_conditional_access_policy_by_id", "arguments": {"policy_id": "policy-1"}}}
```

## ⚠️ Errors & limits

A failed call reaches the AI client twice: one error notification (`<message>: <reason>`, sent through the MCP log channel) and the error result `Error calling tool '<tool>': <reason>` with `isError: true`. The reason is the text of the exception Microsoft Graph, the SDK or the server raised. The server sets no timeout and no retry of its own and keeps no state between calls. Results that report a known case — `already_exists`, `not_found`, a dynamic group's `status: error`, the permission helper's `status: error` — are normal results, not errors.

- Each included or excluded group costs one Graph call; a group that cannot be read shows as `Unknown Group (<id>)`.
- A policy part that cannot be parsed is returned with empty strings and a warning in the server log, not an error.
- Values are flattened for reading, not round-tripping: lists become `; `-joined strings and nested objects their Python text form.
