# entraid-mcp-server configuration

The server reads only environment variables, and only to choose how it signs in to Microsoft Graph. With none of them set it uses your current Azure login (`az login`), so most local setups need no configuration at all.

## How settings reach the server

Each key is an environment variable with the same name; there is no nesting and no prefix to translate. A key can come from two places:

1. **The process environment** — for example the `env` block of your AI client's server entry. A variable set here always wins.
2. **A `.env` file.** When the server starts (`main()`, used by `uvx entraid-mcp-server`, the `entraid-mcp-server` command and `python -m msgraph_mcp_server`), `load_env_files` looks in these 4 places, in this order, and loads the **first** file it finds:
   1. `config/.env` three folders above the installed `auth` package — the repo's own `config/.env` when you run from a clone
   2. `config/.env` under the current working folder
   3. `~/.entraid/.env`
   4. `/etc/entraid/.env`

   A variable the file sets is ignored when the process environment already has it. No file found is normal: the server carries on with the environment as it is. Importing the package loads nothing; only start-up does.

`config/.env.example` in the repo lists the keys with every line commented out.

## How the server picks a credential

`GraphAuthManager` builds the credential on the **first Microsoft Graph call**, not at start-up, and keeps it for the life of the process. It takes the first rule that holds:

1. **Client secret** — tenant id, client id and client secret are all set → `ClientSecretCredential`.
2. **Certificate** — tenant id, client id and certificate path are all set → `CertificateCredential`, with the certificate password when one is set.
3. **Otherwise** → `DefaultAzureCredential`, which tries, among others, environment variables, workload identity, managed identity, the Azure CLI (`az login`), the Azure Developer CLI, Azure PowerShell and Visual Studio / VS Code sign-ins.

Each setting is read from its plain name first, then from its `AZURE_*` name. The token is requested for the `https://graph.microsoft.com/.default` scope, so the identity's granted Graph permissions decide what each tool may do.

## ⚙️ Configuration reference

| Key | Type | Required | Default | Rules | Secret | Takes effect | Effect |
|---|---|---|---|---|---|---|---|
| `TENANT_ID` (or `AZURE_TENANT_ID`) | string | For client-secret or certificate sign-in | none | Entra ID tenant id | no | startup — read on the first Graph call, then cached; restart to change | Tenant for the secret or certificate credential |
| `CLIENT_ID` (or `AZURE_CLIENT_ID`) | string | For client-secret or certificate sign-in | none | Application (client) id of the app registration | no | startup — as above | App the server signs in as |
| `CLIENT_SECRET` (or `AZURE_CLIENT_SECRET`) | string | For client-secret sign-in | none | With `TENANT_ID` and `CLIENT_ID`, selects the client-secret credential | yes — keep it in the client's `env` block or a `.env` file outside version control; example value `<your-client-secret>` | startup — as above | Client secret of that app |
| `CERTIFICATE_PATH` (or `AZURE_CLIENT_CERTIFICATE_PATH`) | string (file path) | For certificate sign-in | none | Used only when no client secret is set | no (the file it points at is) | startup — as above | Certificate file for the certificate credential |
| `CERTIFICATE_PWD` (or `AZURE_CLIENT_CERTIFICATE_PASSWORD`) | string | no | none | Used only with `CERTIFICATE_PATH` | yes — as `CLIENT_SECRET`; example value `<your-certificate-password>` | startup — as above | Password of the certificate file |

`DefaultAzureCredential` also reads its own environment variables (the `AZURE_*` names above among them); see the azure-identity documentation for the full list.

Nothing else is configurable: the server always talks stdio, logs at `INFO` to standard error, and uses Microsoft's public Graph endpoints (`https://graph.microsoft.com/v1.0` and, for the device tools, `/beta`).
