# Devices (Intune) — Integration

Part of [Devices (Intune)](index.md).

## 🌐 Downstream systems

| System | Direction | How | What for | When it is down |
|---|---|---|---|---|
| Microsoft Graph beta (`https://graph.microsoft.com/beta`) | we call it | HTTPS through msgraph-beta-sdk's `GraphServiceClient`, built once by `GraphClient.get_beta_client` | Read Intune managed devices and their relationships, compliance policies, configuration profiles and categories (`/deviceManagement`) | Same as v1.0: no timeout or retry of the server's own; the call ends in the error result. |
| Microsoft Entra ID | we call it | The azure-identity credential `GraphAuthManager` picks, asking for the `https://graph.microsoft.com/.default` scope | A token for the Graph calls; the credential is built on the first Graph call and kept for the life of the process | A credential that cannot be built raises `AuthenticationError`; a token that cannot be obtained fails the Graph call. Both end the call in the error result. |

No configuration key points at these systems: the SDKs use Microsoft's public cloud addresses. The keys that choose how the server signs in are listed once, in [the configuration reference](../../configuration.md).
