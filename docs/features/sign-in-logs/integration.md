# Sign-in logs — Integration

Part of [Sign-in logs](index.md).

## 🌐 Downstream systems

| System | Direction | How | What for | When it is down |
|---|---|---|---|---|
| Microsoft Graph v1.0 (`https://graph.microsoft.com/v1.0`) | we call it | HTTPS through msgraph-sdk's `GraphServiceClient`, built once by `GraphClient.get_client` | Read sign-in logs (`/auditLogs/signIns`) | The server sets no timeout and no retry of its own. The SDK's exception ends the call in the error result `Error calling tool '<tool>': <reason>`, after one error notification. |
| Microsoft Entra ID | we call it | The azure-identity credential `GraphAuthManager` picks, asking for the `https://graph.microsoft.com/.default` scope | A token for the Graph calls; the credential is built on the first Graph call and kept for the life of the process | A credential that cannot be built raises `AuthenticationError`; a token that cannot be obtained fails the Graph call. Both end the call in the error result. |

No configuration key points at these systems: the SDKs use Microsoft's public cloud addresses. The keys that choose how the server signs in are listed once, in [the configuration reference](../../configuration.md).
