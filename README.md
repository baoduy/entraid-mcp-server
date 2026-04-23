# EntraID MCP Server (Microsoft Graph FastMCP)

This project provides a modular, resource-oriented FastMCP server for interacting with Microsoft Graph API. It is designed for extensibility, maintainability, and security, supporting advanced queries for users, sign-in logs, MFA status, and privileged users.

## Features

- **Modular Resource Structure:**
  - Each resource (users, sign-in logs, MFA, etc.) is implemented in its own module under `src/msgraph_mcp_server/resources/`.
  - Easy to extend with new resources (e.g., groups, devices).
- **Centralized Graph Client:**
  - Handles authentication and client initialization.
  - Shared by all resource modules.
- **Comprehensive User Operations:**
  - Search users by name/email.
  - Get user by ID.
  - List all privileged users (directory role members).
- **Full Group Lifecycle & Membership Management:**
  - Create, read, update, and delete groups.
  - Add/remove group members and owners.
  - Search and list groups and group members.
- **Application & Service Principal Management:**
  - List, create, update, and delete applications (app registrations).
  - List, create, update, and delete service principals.
  - View app role assignments and delegated permissions for both applications and service principals.
- **Sign-in Log Operations:**
  - Query sign-in logs for a user for the last X days.
- **MFA Operations:**
  - Get MFA status for a user.
  - Get MFA status for all members of a group.
- **Password Management:**
  - Reset user passwords directly with custom or auto-generated secure passwords.
  - Option to require password change on next sign-in.
- **Permissions Helper:**
  - Suggest appropriate Microsoft Graph permissions for common tasks.
  - Search and explore available Graph permissions.
  - Helps implement the principle of least privilege by recommending only necessary permissions.
- **Error Handling & Logging:**
  - Consistent error handling and progress reporting via FastMCP context.
  - Detailed logging for troubleshooting.
- **Security:**
  - `.env` and secret files are excluded from version control.
  - Uses Microsoft best practices for authentication.

## Project Structure

```
src/msgraph_mcp_server/
├── auth/           # Authentication logic (GraphAuthManager)
├── resources/      # Resource modules (users, signin_logs, mfa, ...)
│   ├── users.py            # User operations (search, get by ID, etc.)
│   ├── signin_logs.py      # Sign-in log operations
│   ├── mfa.py              # MFA status operations
│   ├── permissions_helper.py # Graph permissions utilities and suggestions
│   ├── applications.py       # Application (app registration) operations
│   ├── service_principals.py # Service principal operations
│   └── ...                 # Other resource modules
├── utils/          # Core GraphClient and other ultilities tool, such as password generator..
├── server.py       # FastMCP server entry point (registers tools/resources)
├── __init__.py     # Package marker
```

## Usage

### 1. Setup
- Clone the repo.
- **Authentication (default – use your current Azure login):**
  By default the server uses
  [`DefaultAzureCredential`](https://learn.microsoft.com/python/api/azure-identity/azure.identity.defaultazurecredential)
  from `azure-identity`. It will automatically pick up your current Azure
  login context in this order: environment variables, workload identity,
  managed identity, Azure CLI (`az login`), Azure Developer CLI
  (`azd auth login`), Azure PowerShell, and Visual Studio / VS Code. For
  local development simply run:
  ```bash
  az login
  ```
  and the server will use that token – **no tenant/client id or secret
  configuration is required**.
- **Authentication (optional – explicit credentials):**
  If you need to force a specific credential (for example, for CI where
  no interactive login is available), create a `config/.env` file based on
  [`config/.env.example`](config/.env.example). Two flows are supported:
  - Client secret: set `TENANT_ID`, `CLIENT_ID`, `CLIENT_SECRET`.
  - Certificate: set `TENANT_ID`, `CLIENT_ID`, `CERTIFICATE_PATH`, and
    optionally `CERTIFICATE_PWD`.

  The standard `AZURE_*` environment variables consumed by
  `DefaultAzureCredential` (e.g. `AZURE_TENANT_ID`, `AZURE_CLIENT_ID`,
  `AZURE_CLIENT_SECRET`) are also honoured.

### 2. Testing & Development

You can test and develop your MCP server directly using the FastMCP CLI:

```bash
fastmcp dev '/path/to/src/msgraph_mcp_server/server.py'
```

This launches an interactive development environment with the MCP Inspector. For more information and advanced usage, see the [FastMCP documentation](https://github.com/jlowin/fastmcp).

### 3. Available Tools

#### User Tools
- `search_users(query, ctx, limit=10)` — Search users by name/email
- `get_user_by_id(user_id, ctx)` — Get user details by ID
- `get_privileged_users(ctx)` — List all users in privileged directory roles
- `get_user_roles(user_id, ctx)` — Get all directory roles assigned to a user
- `get_user_groups(user_id, ctx)` — Get all groups (including transitive memberships) for a user

#### Group Tools
- `get_all_groups(ctx, limit=100)` — Get all groups (with paging)
- `get_group_by_id(group_id, ctx)` — Get a specific group by its ID
- `search_groups_by_name(name, ctx, limit=50)` — Search for groups by display name
- `get_group_members(group_id, ctx, limit=100)` — Get members of a group by group ID
- `create_group(ctx, group_data)` — Create a new group (see below for group_data fields)
- `update_group(group_id, ctx, group_data)` — Update an existing group (fields: displayName, mailNickname, description, visibility)
- `delete_group(group_id, ctx)` — Delete a group by its ID
- `add_group_member(group_id, member_id, ctx)` — Add a member (user, group, device, etc.) to a group
- `remove_group_member(group_id, member_id, ctx)` — Remove a member from a group
- `add_group_owner(group_id, owner_id, ctx)` — Add an owner to a group
- `remove_group_owner(group_id, owner_id, ctx)` — Remove an owner from a group

**Group Creation/Update Example:**
- `group_data` for `create_group` and `update_group` should be a dictionary with keys such as:
  - `displayName` (required for create)
  - `mailNickname` (required for create)
  - `description` (optional)
  - `groupTypes` (optional, e.g., `["Unified"]`)
  - `mailEnabled` (optional)
  - `securityEnabled` (optional)
  - `visibility` (optional, "Private" or "Public")
  - `owners` (optional, list of user IDs)
  - `members` (optional, list of IDs)
  - `membershipRule` (required for dynamic groups)
  - `membershipRuleProcessingState` (optional, "On" or "Paused")

See the `groups.py` docstrings for more details on supported fields and behaviors.

#### Sign-in Log Tools
- `get_user_sign_ins(user_id, ctx, days=7)` — Get sign-in logs for a user

#### MFA Tools
- `get_user_mfa_status(user_id, ctx)` — Get MFA status for a user
- `get_group_mfa_status(group_id, ctx)` — Get MFA status for all group members

#### Device Tools
- `get_all_managed_devices(filter_os=None)` — Get all managed devices (optionally filter by OS)
- `get_managed_devices_by_user(user_id)` — Get all managed devices for a specific user

#### Conditional Access Policy Tools
- `get_conditional_access_policies(ctx)` — Get all conditional access policies
- `get_conditional_access_policy_by_id(policy_id, ctx)` — Get a single conditional access policy by its ID

#### Audit Log Tools
- `get_user_audit_logs(user_id, days=30)` — Get all relevant directory audit logs for a user by user_id within the last N days

#### Password Management Tools
- `reset_user_password_direct(user_id, password=None, require_change_on_next_sign_in=True, generate_password=False, password_length=12)` — Reset a user's password with a specific password value or generate a secure random password

#### Permissions Helper Tools
- `suggest_permissions_for_task(task_category, task_name)` — Suggest Microsoft Graph permissions for a specific task based on common mappings
- `list_permission_categories_and_tasks()` — List all available categories and tasks for permission suggestions
- `get_all_graph_permissions()` — Get all Microsoft Graph permissions directly from the Microsoft Graph API
- `search_permissions(search_term, permission_type=None)` — Search for Microsoft Graph permissions by keyword

#### Application Tools
- `list_applications(ctx, limit=100)` — List all applications (app registrations) in the tenant, with paging
- `get_application_by_id(app_id, ctx)` — Get a specific application by its object ID (includes app role assignments and delegated permissions)
- `create_application(ctx, app_data)` — Create a new application (see below for app_data fields)
- `update_application(app_id, ctx, app_data)` — Update an existing application (fields: displayName, signInAudience, tags, identifierUris, web, api, requiredResourceAccess)
- `delete_application(app_id, ctx)` — Delete an application by its object ID

**Application Creation/Update Example:**
- `app_data` for `create_application` and `update_application` should be a dictionary with keys such as:
  - `displayName` (required for create)
  - `signInAudience` (optional)
  - `tags` (optional)
  - `identifierUris` (optional)
  - `web` (optional)
  - `api` (optional)
  - `requiredResourceAccess` (optional)

#### Service Principal Tools
- `list_service_principals(ctx, limit=100)` — List all service principals in the tenant, with paging
- `get_service_principal_by_id(sp_id, ctx)` — Get a specific service principal by its object ID (includes app role assignments and delegated permissions)
- `create_service_principal(ctx, sp_data)` — Create a new service principal (see below for sp_data fields)
- `update_service_principal(sp_id, ctx, sp_data)` — Update an existing service principal (fields: displayName, accountEnabled, tags, appRoleAssignmentRequired)
- `delete_service_principal(sp_id, ctx)` — Delete a service principal by its object ID

**Service Principal Creation/Update Example:**
- `sp_data` for `create_service_principal` and `update_service_principal` should be a dictionary with keys such as:
  - `appId` (required for create)
  - `accountEnabled` (optional)
  - `tags` (optional)
  - `appRoleAssignmentRequired` (optional)
  - `displayName` (optional)

#### Example Resource
- `greeting://{name}` — Returns a personalized greeting

## Extending the Server
- Add new resource modules under `resources/` (e.g., `groups.py`, `devices.py`).
- Register new tools in `server.py` using the FastMCP `@mcp.tool()` decorator.
- Use the shared `GraphClient` for all API calls.

## Security & Best Practices
- **Never commit secrets:** `.env` and other sensitive files are gitignored.
- **Use least privilege:** Grant only the necessary Microsoft Graph permissions to your Azure AD app.
- **Audit & monitor:** Use the logging output for troubleshooting and monitoring.

## Required Graph API Permissions
| API / Permission            | Type        | Description                               |
|-----------------------------|-------------|-------------------------------------------|
| AuditLog.Read.All           | Application | Read all audit log data                   |
| AuthenticationContext.Read.All | Application | Read all authentication context information |
| DeviceManagementManagedDevices.Read.All | Application | Read Microsoft Intune devices |
| Directory.Read.All          | Application | Read directory data                       |
| Group.Read.All              | Application | Read all groups                           |
| GroupMember.Read.All        | Application | Read all group memberships                |
| Group.ReadWrite.All         | Application | Create, update, delete groups; manage group members and owners |
| Policy.Read.All             | Application | Read your organization's policies         |
| RoleManagement.Read.Directory | Application | Read all directory RBAC settings        |
| User.Read.All               | Application | Read all users' full profiles             |
| User-PasswordProfile.ReadWrite.All | Application | Least privileged permission to update the passwordProfile property |
| UserAuthenticationMethod.Read.All | Application | Read all users' authentication methods |
| Application.ReadWrite.All   | Application | Create, update, and delete applications (app registrations) and service principals |

**Note:** `Group.ReadWrite.All` is required for group creation, update, deletion, and for adding/removing group members or owners. `Group.Read.All` and `GroupMember.Read.All` are sufficient for read-only group and membership queries.

## Advanced: Using with Claude or Cursor

### Install from PyPI and run via `uvx`

Once the package has been published to PyPI you can run the server
directly with [`uvx`](https://docs.astral.sh/uv/) — no repo clone,
`--with` list, or file path required:

```bash
uvx entraid-mcp-server
```

Pin a specific version with either form:

```bash
uvx entraid-mcp-server@0.2.0
uvx --from "entraid-mcp-server==0.2.0" entraid-mcp-server
```

Equivalent `mcp.json` entry (Cursor / Claude Desktop):

```json
{
  "EntraID MCP Server": {
    "command": "uvx",
    "args": ["entraid-mcp-server"]
  }
}
```

Authentication still defaults to `DefaultAzureCredential`, so `az login`
is enough for local use; add an `env` block with `TENANT_ID` /
`CLIENT_ID` / `CLIENT_SECRET` only if you need non-interactive auth.

### Building and publishing the package

From the repo root:

```bash
uv build                     # produces dist/*.whl and dist/*.tar.gz
uvx --from ./dist/entraid_mcp_server-<version>-py3-none-any.whl \
    entraid-mcp-server       # local smoke test
```

Upload to PyPI (create an API token at https://pypi.org first):

```bash
# optional dry run against TestPyPI
uv publish --publish-url https://test.pypi.org/legacy/ --token <test-token>

# real release
uv publish --token <pypi-token>
```

For a hands-off release flow, configure
[PyPI Trusted Publishing (OIDC)](https://docs.pypi.org/trusted-publishers/)
and add a GitHub Actions workflow that runs `uv build` + `uv publish`
on every tag push — no long-lived tokens required.

### Using with Claude (Anthropic)
To install and run this server as a Claude MCP tool, use:

```bash
fastmcp install '/path/to/src/msgraph_mcp_server/server.py' \
  --with msgraph-sdk --with azure-identity --with azure-core --with msgraph-core \
  -f /path/to/.env
```
- Replace `/path/to/` with your actual project path.
- The `-f` flag points to your `.env` file (never commit secrets!).

### Using with Cursor
Add the following to your `.cursor/mcp.json`. By default the server uses
your current Azure login via `DefaultAzureCredential`, so no `env` block
is required:

```json
{
  "EntraID MCP Server": {
    "command": "uv",
    "args": [
      "run",
      "--with", "azure-core",
      "--with", "azure-identity",
      "--with", "fastmcp",
      "--with", "msgraph-core",
      "--with", "msgraph-sdk",
      "fastmcp",
      "run",
      "/path/to/src/msgraph_mcp_server/server.py"
    ]
  }
}
```

If you need to force client-secret authentication instead (do **not**
commit real secrets), add an `env` block:

```json
"env": {
  "TENANT_ID": "<your-tenant-id>",
  "CLIENT_ID": "<your-client-id>",
  "CLIENT_SECRET": "<your-client-secret>"
}
```

## License

MIT
