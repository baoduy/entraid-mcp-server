# MFA

Check which authentication methods a user, or every user in a group, has registered — and whether that counts as MFA.

## 📖 Overview

- Read one user's registered methods as eight yes/no flags, and an overall `Enabled` or `Disabled`.
- Report on every direct user member of a group in one call; members that are not users are skipped.
- One member whose methods cannot be read does not stop the report: the count of unreadable members comes back in one warning.

Called by an AI client on behalf of an Entra ID admin reviewing MFA coverage.

## 🏢 Business domain

Authentication method registration. Sign-in events belong to [Sign-in logs](../sign-in-logs/index.md); conditional access policies that require MFA to [Conditional access](../conditional-access/index.md).

| Term | Meaning | In the code |
|---|---|---|
| **Authentication method** | A way a user can prove who they are: password, phone, authenticator app, FIDO2 key, and so on | `/users/{id}/authentication/methods` |
| **MFA status** | `Enabled` when any method other than password is registered, else `Disabled` | `mfaStatus` in `mfa.get_mfa_status` |
| **Unreadable member** | A user member whose methods Graph would not return | `unreadable` in `mfa.get_group_mfa_status` |

| Rule | Enforced by | A caller who breaks it gets |
|---|---|---|
| Password alone is not MFA; any other method, email included, is | the method checks in `mfa.get_mfa_status` | — |
| Only direct user members are reported; nested groups are not expanded | the `#microsoft.graph.user` filter in `mfa.get_group_mfa_status` | — |
| An unreadable member is counted, not fatal | the `try` around each member in `mfa.get_group_mfa_status` | Warning `MFA status could not be read for <n> of <users> members` |

## 🚀 Quick Start

No token on the call: the server signs in itself, as its [configuration](../../configuration.md) says, on the first Microsoft Graph call. Set the server up in your AI client first ([Quick Start](../../index.md#-quick-start)); then ask the client in plain words, or send the call yourself:

```json
{"jsonrpc": "2.0", "id": 1, "method": "tools/call", "params": {"name": "get_user_mfa_status", "arguments": {"user_id": "aoife@contoso.com"}}}
```

The result's `structuredContent`:

```json
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
```

Captured by running the server's real MCP handler against the test suite's in-memory Microsoft Graph (`tests/conftest.py`), not a live tenant: the values are test data, the shape is the server's.

## 🔄 End-to-end flow

![An AI client calls get_group_mfa_status; the module reads every page of group members, keeps users, then for each user reads companyName and the authentication methods, counts unreadable members, and returns the statuses with one count warning when some could not be read.](diagrams/mfa-group.svg)

`get_group_mfa_status` stands for the feature. Nothing is stored. Members are read page by page, then two Graph calls per user run one after another, so a large group takes a while. A failed member lookup is logged and counted; the report still returns the members that could be read. `userPrincipalName` in each status is the id that was looked up, not a value from Graph.

## 📚 Pages

- [Endpoints](endpoints.md) — the 2 tools, each with its mark, arguments, result, errors and an example call
- [Integration](integration.md) — the Microsoft services the feature calls and what happens when they fail

## 🔗 Related features

- [Groups](../groups/index.md) — reach for it to find the group or its members.
- [Password management](../password-management/index.md) — reach for it to list a user's password methods or reset the password.

## ❓ Open questions

| Question | Why it matters | Checked | Who can answer |
|---|---|---|---|
| Does Microsoft Graph return `mail` when the user read selects only `companyName`? | If not, `mail` is always `null` in MFA results. | `mfa.get_mfa_status` sends `$select=companyName` and then reads `mail`; the test suite's Graph ignores `$select`, so the captured example shows `mail` filled. | Maintainer (baoduy), by a run against a tenant |
