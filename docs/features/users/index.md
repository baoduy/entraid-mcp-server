# Users

Look up people in the tenant: find a user, read a profile, and see which groups and directory roles they hold.

## 📖 Overview

- Find a user by part of their name or email address; the search covers display name, mail, user principal name, given name, surname and other mail addresses.
- Read one user's profile by object id or user principal name.
- List everyone who holds an activated directory role — each person once, with all their role names — to review privileged access.
- List every group a user belongs to, nested memberships included, and the directory roles assigned to them directly.
- Nothing here changes the tenant: all five tools are read-only.

Called by an AI client (Claude Desktop, Cursor) on behalf of an Entra ID admin.

## 🏢 Business domain

Identity lookup in the directory. It reads users and their memberships; changing a membership belongs to [Groups](../groups/index.md), and authentication methods to [MFA](../mfa/index.md).

| Term | Meaning | In the code |
|---|---|---|
| **User** | A person's account in the tenant | `users.get_user_by_id`, `/users` |
| **Privileged user** | A user who is a member of an activated directory role | `users.get_privileged_users` |
| **Directory role** | An Entra ID admin role such as Global Administrator, activated in the tenant | `/directoryRoles`, `roleTemplateId` |
| **Transitive membership** | Membership through nested groups as well as direct membership | `transitiveMemberOf` in `users.get_user_groups` |
| **User principal name** | The sign-in name, such as `aoife@contoso.com`; accepted wherever a user id is | `userPrincipalName` |

| Rule | Enforced by | A caller who breaks it gets |
|---|---|---|
| Search text cannot add query syntax: `\` and `"` are escaped | `graph_search_escape` in `utils/odata.py` | Their text matched literally |
| A user appears once in the privileged list, whatever the number of roles | `users.get_privileged_users` keys on the user id | — |
| Only users are listed as privileged; service principals in a role are left out | the `#microsoft.graph.user` check in `users.get_privileged_users` | — |

## 🚀 Quick Start

No token on the call: the server signs in itself, as its [configuration](../../configuration.md) says, on the first Microsoft Graph call. Set the server up in your AI client first ([Quick Start](../../index.md#-quick-start)); then ask the client in plain words, or send the call yourself:

```json
{"jsonrpc": "2.0", "id": 1, "method": "tools/call", "params": {"name": "get_user_by_id", "arguments": {"user_id": "aoife@contoso.com"}}}
```

The result's `structuredContent` — list and optional results arrive under `result`:

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

Captured by running the server's real MCP handler against the test suite's in-memory Microsoft Graph (`tests/conftest.py`), not a live tenant: the values are test data, the shape is the server's.

## 🔄 End-to-end flow

![An AI client calls search_users; the tool asks the users module, which sends GET /users with $search and ConsistencyLevel eventual to Microsoft Graph, follows nextLink until the limit, and returns up to limit users or, when Graph refuses, an error notification and error result.](diagrams/users-search.svg)

`search_users` stands for the feature: the most common way in. Nothing is stored and no event is raised — the flow ends at Microsoft Graph. The search reads with eventual consistency, so a user created moments ago may be missing.

## 📚 Pages

- [Endpoints](endpoints.md) — the 5 tools, each with its mark, arguments, result, errors and an example call
- [Integration](integration.md) — the Microsoft services the feature calls and what happens when they fail

## 🔗 Related features

- [Groups](../groups/index.md) — reach for it to change a membership you found here.
- [MFA](../mfa/index.md) — reach for it to check the authentication methods a user registered.
- [Sign-in logs](../sign-in-logs/index.md) — reach for it to see where and how a user signed in.
- [Audit logs](../audit-logs/index.md) — reach for it to see what a user changed.
