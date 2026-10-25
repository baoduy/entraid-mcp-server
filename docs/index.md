# EntraID MCP Server

An MCP server that lets an AI client such as Claude Desktop or Cursor read and manage your Microsoft Entra ID tenant — users, groups, MFA, devices, policies, logs and app registrations — through 48 Microsoft Graph tools.

## 📖 Overview

- [Users](features/users/index.md) — search users, read a profile, list privileged users and a user's groups and roles (`resources/users.py`).
- [Groups](features/groups/index.md) — find, create, change and delete groups; add and remove members and owners (`resources/groups.py`).
- [Sign-in logs](features/sign-in-logs/index.md) — a user's sign-ins over the last days (`resources/signin_logs.py`).
- [MFA](features/mfa/index.md) — registered authentication methods for a user or every user in a group (`resources/mfa.py`).
- [Devices (Intune)](features/devices/index.md) — managed devices, detected apps, compliance and configuration, on the Graph beta endpoint (`resources/managed_devices.py`).
- [Conditional access](features/conditional-access/index.md) — the tenant's policies as flat rows (`resources/conditional_access.py`).
- [Audit logs](features/audit-logs/index.md) — the directory changes a user made (`resources/audit_logs.py`).
- [Password management](features/password-management/index.md) — reset a password to a given or generated value; read password methods (`resources/password_auth.py`).
- [Permissions helper](features/permissions-helper/index.md) — suggest and search Microsoft Graph permissions (`resources/permissions_helper.py`).
- [Applications](features/applications/index.md) — list, read, create, change and delete app registrations (`resources/applications.py`).
- [Service principals](features/service-principals/index.md) — list, read, create, change and delete service principals (`resources/service_principals.py`).

Every tool carries an MCP mark, so a client can let lookups run and ask before a change: 34 tools are read-only; 5 change the tenant without destroying anything (the 3 creates, add member, add owner); 9 are destructive (the 3 deletes, remove member, remove owner, the password reset and the 3 updates).

The server is started and called by an AI client on behalf of an Entra ID admin. It runs on the admin's machine, talks to the client over standard input and output, and signs in to Microsoft Graph as the admin (or as an app you configure).

## 🏗️ Runtime architecture

![An AI client sends tools/call over stdio to the FastMCP server, which dispatches to a tool function in server.py, then to a resource module, which uses GraphClient to call Microsoft Graph v1.0 (or beta for devices); on the first call GraphAuthManager gets a token from Microsoft Entra ID; main() starts the server after loading the first .env file found.](diagrams/runtime.svg)

## 🌐 Downstream systems

| System | What for | How | When it is down |
|---|---|---|---|
| Microsoft Graph v1.0 | Every tool except the 9 device tools and the 2 permission-suggestion tools | HTTPS through msgraph-sdk | The tool call ends in an error result; the server sets no timeout or retry of its own |
| Microsoft Graph beta | The 9 device (Intune) tools | HTTPS through msgraph-beta-sdk | As above |
| Microsoft Entra ID | A token for Graph, on the first Graph call | The azure-identity credential the [configuration](configuration.md) selects | The tool call ends in an error result |

The identity the server signs in as needs Graph permissions for the tools you use. Microsoft Graph checks them; the server checks none. This list is carried over from the previous README (application permissions):

| Permission | For |
|---|---|
| `User.Read.All` | Users |
| `Directory.Read.All` | Directory data: users, groups, roles |
| `RoleManagement.Read.Directory` | Directory roles |
| `Group.Read.All`, `GroupMember.Read.All` | Reading groups and members |
| `Group.ReadWrite.All` | Creating, changing and deleting groups; adding and removing members and owners |
| `AuditLog.Read.All` | Sign-in logs and audit logs |
| `UserAuthenticationMethod.Read.All` | MFA and password methods |
| `User-PasswordProfile.ReadWrite.All` | Password reset (see ❓ Open questions) |
| `Policy.Read.All`, `AuthenticationContext.Read.All` | Conditional access policies |
| `DeviceManagementManagedDevices.Read.All`, `DeviceManagementConfiguration.Read.All`, `DeviceManagementApps.Read.All` | Device tools |
| `Application.ReadWrite.All` | Creating, changing and deleting applications and service principals |

## 🚀 Quick Start

You need [`uv`](https://docs.astral.sh/uv/) and Python 3.10 or later.

1. Sign in to Azure with an account that has the permissions above. The server uses this login by default:

   ```bash
   az login
   ```

2. Add the server to the `mcpServers` object of your AI client's MCP settings — Claude Desktop: `claude_desktop_config.json`; Cursor: `.cursor/mcp.json`:

   ```json
   {
     "mcpServers": {
       "entraid": {
         "command": "uvx",
         "args": ["entraid-mcp-server"]
       }
     }
   }
   ```

   `uvx` downloads the latest release from PyPI and starts it; add `@<version>` to the package name to pin one.

3. To sign in as an app instead of your Azure login (no interactive login available, for example), add an `env` block to the same entry. Never commit real values:

   ```json
   "env": {
     "TENANT_ID": "<your-tenant-id>",
     "CLIENT_ID": "<your-client-id>",
     "CLIENT_SECRET": "<your-client-secret>"
   }
   ```

   Certificate sign-in, `.env` files and the order the server tries credentials in: [configuration](configuration.md).

4. Restart the client so it starts the server.

To run from a clone instead (from the repo root, with `uv` installed):

```bash
uv sync
uv run entraid-mcp-server          # or: uv run python -m msgraph_mcp_server
```

## ✅ Verify it works

In your AI client, list the server's tools: you should see **48**. Then call `list_permission_categories_and_tasks` — it needs no Graph access, so it works before any permission is granted. The result starts:

```json
{"status": "success", "categories": {"users": {"tasks": [{"name": "read_user_profile", "description": "Read user profile information"}, ...
```

with the categories `users`, `groups`, `mail`, `calendar`, `files`, `devices`, `audit_logs` and `directory`. The server signs in only on the first tool that calls Microsoft Graph, so a sign-in problem shows up there, as the tool's error result.

## 🛠️ Common commands

All from the repo root:

| Command | What it does |
|---|---|
| `uv sync --locked` | Install the locked dependencies, dev tools included |
| `uv build` | Build the wheel and source archive into `dist/` |
| `uv run pytest` | Run the unit tests (an in-memory Microsoft Graph; no tenant needed) |
| `uv run ruff check` | Lint |
| `uv run ruff format --check` | Check formatting |
| `uv run entraid-mcp-server` | Start the server on stdio |

Releases are automatic: there is no publish command to run. A pull request to `dev`, and a push to `dev`, run the check workflow (build, tests, lint, format). When `dev` is merged into `main` and the push touches `src/msgraph_mcp_server/`, `pyproject.toml`, `README.md`, `LICENSE` or the publish workflow, the publish workflow (which can also be started by hand on `main`) computes the next version from the git tags, builds the package, publishes it to PyPI and creates the GitHub release.

## ⚠️ Known limitations

- `search_users` declares every user field a string, but `businessPhones` is a list and empty fields are `null`. A client that validates a tool's structured output rejects any result holding a user ([Users](features/users/index.md)).
- The device tools use the Microsoft Graph **beta** endpoint, which Microsoft can change without notice.
- The server sets no timeout or retry of its own; a slow or failing Graph call ends the tool call with an error result.
- `get_conditional_access_policies` reads the first page of policies only.
- `search_groups_by_name` matches the start of the display name, not any part of it.
- `create_group` is several Graph calls with no rollback: if an owner or member add fails, the group stays created.
- `reset_user_password_direct` returns a generated password in its result, so it reaches the AI client's conversation.
- Membership checks and searches read Graph with eventual consistency: a change made moments earlier may not be seen yet.

## 📚 Documentation

| Page | Answers |
|---|---|
| [Users](features/users/index.md) | How to find users and read their groups and roles |
| [Groups](features/groups/index.md) | How groups, members and owners are managed, and what each result means |
| [Sign-in logs](features/sign-in-logs/index.md) | How to read a user's sign-ins |
| [MFA](features/mfa/index.md) | How MFA status is worked out for a user or a group |
| [Devices (Intune)](features/devices/index.md) | Which device, compliance and configuration data you can read |
| [Conditional access](features/conditional-access/index.md) | How policies are flattened into rows |
| [Audit logs](features/audit-logs/index.md) | How to read the changes a user made |
| [Password management](features/password-management/index.md) | How a reset works and how the password is protected |
| [Permissions helper](features/permissions-helper/index.md) | How to find the Graph permissions a task needs |
| [Applications](features/applications/index.md) | How app registrations are managed |
| [Service principals](features/service-principals/index.md) | How service principals are managed |
| [Configuration](configuration.md) | Which settings choose how the server signs in, and where it reads them |
| [LICENSE](../LICENSE) | MIT licence |

## ❓ Open questions

| Question | Why it matters | Checked | Who can answer |
|---|---|---|---|
| Is the permissions table above complete and least-privilege for the current 48 tools? | Admins grant exactly what this table says. | The code checks no permission; the table is carried over from the previous README. Only `get_user_sign_ins` names a permission in code (`AuditLog.Read.All`). | Maintainer (baoduy) |
| Does the password reset need `User-PasswordProfile.ReadWrite.All` (this table) or `User.ReadWrite.All` (the permissions helper's suggestion)? | A wrong grant makes the reset fail, or grants more than needed. | `reset_user_password_direct` sends `PATCH /users/{id}` with `passwordProfile`; `COMMON_PERMISSION_MAPPINGS` lists `User.ReadWrite.All`. | Maintainer (baoduy) |
| Where should users report problems or ask for support? | The README names no support route. | `pyproject.toml` lists the GitHub issues page as `Issues`; no support policy is in the repo. | Maintainer (baoduy) |
