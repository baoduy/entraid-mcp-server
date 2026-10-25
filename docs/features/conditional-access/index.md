# Conditional access

Review the tenant's conditional access policies as flat, readable rows — who they cover, what they require, and session controls.

## 📖 Overview

- List the policies, or read one, with its state, conditions, grant controls and session controls.
- Group ids in a policy are shown with the group's name next to them; a group that cannot be read shows as `Unknown Group (<id>)`.
- List values are joined into one string, so each policy reads as one row.

Called by an AI client on behalf of an Entra ID admin auditing access policy. Both tools are read-only.

## 🏢 Business domain

Access policy. The groups a policy names belong to [Groups](../groups/index.md); the methods it may require to [MFA](../mfa/index.md).

| Term | Meaning | In the code |
|---|---|---|
| **Conditional access policy** | An Entra ID rule: when these conditions hold, require these controls | `/identity/conditionalAccess/policies` |
| **Conditions** | Users, groups, roles, apps, platforms, locations, risk levels and client app types the policy covers | `parse_conditions` |
| **Grant controls** | What a user must satisfy: MFA, compliant device, authentication strength, terms of use | `parse_grant_controls`, columns prefixed `Grant_` |
| **Session controls** | Limits on the session: sign-in frequency, persistent browser and others | `parse_session_controls` |
| **State** | The policy state as Graph reports it, such as `enabled` | `state` |

| Rule | Enforced by | A caller who breaks it gets |
|---|---|---|
| A part that cannot be parsed is left blank, not fatal | the `try` blocks in `parse_conditions`, `parse_grant_controls`, `parse_session_controls` | Empty strings and a server log warning |
| A group that cannot be read does not fail the policy | `get_group_details` | `Unknown Group (<id>)` |

## 🚀 Quick Start

No token on the call: the server signs in itself, as its [configuration](../../configuration.md) says, on the first Microsoft Graph call. Set the server up in your AI client first ([Quick Start](../../index.md#-quick-start)); then ask the client in plain words, or send the call yourself:

```json
{"jsonrpc": "2.0", "id": 1, "method": "tools/call", "params": {"name": "get_conditional_access_policy_by_id", "arguments": {"policy_id": "policy-1"}}}
```

The result's `structuredContent` (abridged: the first 10 of 48 fields):

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

Captured by running the server's real MCP handler against the test suite's in-memory Microsoft Graph (`tests/conftest.py`), not a live tenant: the values are test data, the shape is the server's.

## 🔄 End-to-end flow

![An AI client calls get_conditional_access_policies; the module reads the first page of policies, reads each included or excluded group's name, and returns flat rows of conditions, grant and session controls.](diagrams/conditional-access-list.svg)

`get_conditional_access_policies` stands for the feature. Nothing is stored. Only the first page of policies is read. Each group named in a policy costs one more Graph call; a group that cannot be read is named `Unknown Group (<id>)` and the policy is still returned.

## 📚 Pages

- [Endpoints](endpoints.md) — the 2 tools, each with its mark, arguments, result, errors and an example call
- [Integration](integration.md) — the Microsoft services the feature calls and what happens when they fail

## 🔗 Related features

- [Groups](../groups/index.md) — reach for it to read a group a policy includes or excludes.
- [MFA](../mfa/index.md) — reach for it to check whether users can satisfy a policy that requires MFA.

## ❓ Open questions

| Question | Why it matters | Checked | Who can answer |
|---|---|---|---|
| Should `get_conditional_access_policies` read every page of policies? | A tenant with more policies than Graph puts on one page gets a partial list without being told. | `get_conditional_access_policies` makes one `GET` and does not follow `@odata.nextLink`; Graph's page size for this list is not in the repo. | Maintainer (baoduy) |
