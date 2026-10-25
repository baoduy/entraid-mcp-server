# Devices (Intune)

Inspect Intune-managed devices: who owns them, how they are built and enrolled, and whether they meet compliance and configuration policies.

## 📖 Overview

- List every managed device, or only one operating system's, or one user's devices.
- Read one device in full: hardware, enrolment, compliance state, encryption, malware counts.
- See the apps Intune detected on a device and how the device stands against each compliance policy and configuration profile.
- List the tenant's compliance policies, configuration profiles and device categories.
- All nine tools are read-only and use the Microsoft Graph beta endpoint, where these Intune fields live.

Called by an AI client on behalf of an Entra ID or Intune admin.

## 🏢 Business domain

Device management (Intune). Entra ID device objects as group members belong to [Groups](../groups/index.md); a device's sign-ins appear in [Sign-in logs](../sign-in-logs/index.md).

| Term | Meaning | In the code |
|---|---|---|
| **Managed device** | A device enrolled in Intune | beta `/deviceManagement/managedDevices` |
| **Compliance policy** | Intune rules a device must meet | `deviceCompliancePolicies`, `deviceCompliancePolicyStates` |
| **Configuration profile** | Intune settings pushed to devices | `deviceConfigurations`, `deviceConfigurationStates` |
| **Detected app** | An application Intune found installed on a device | `detectedApps` |
| **Device category** | An admin-defined label for devices | `deviceCategories` |
| **Beta endpoint** | Microsoft Graph's preview API at `graph.microsoft.com/beta` | `GraphClient.get_beta_client` |

| Rule | Enforced by | A caller who breaks it gets |
|---|---|---|
| The OS filter and user id cannot add a filter condition: a `'` is doubled | `odata_string_literal` in `utils/odata.py` | Their text matched literally |
| Device records keep the same 50 fields everywhere | `_format_managed_device` in `managed_devices` | — |

## 🚀 Quick Start

No token on the call: the server signs in itself, as its [configuration](../../configuration.md) says, on the first Microsoft Graph call. Set the server up in your AI client first ([Quick Start](../../index.md#-quick-start)); then ask the client in plain words, or send the call yourself:

```json
{"jsonrpc": "2.0", "id": 1, "method": "tools/call", "params": {"name": "get_managed_device_by_id", "arguments": {"device_id": "device-1"}}}
```

The result's `structuredContent` (abridged: the first 10 of 50 fields):

```json
{
  "id": "device-1",
  "deviceName": "LAPTOP-0042",
  "managedDeviceName": null,
  "userId": "7f3c9a",
  "userPrincipalName": null,
  "userDisplayName": null,
  "emailAddress": null,
  "operatingSystem": "Windows",
  "osVersion": null,
  "manufacturer": null,
  "…": "40 more fields"
}
```

Captured by running the server's real MCP handler against the test suite's in-memory Microsoft Graph (`tests/conftest.py`), not a live tenant: the values are test data, the shape is the server's.

## 🔄 End-to-end flow

![An AI client calls get_managed_devices_by_user; the module sends a beta GET managedDevices filtered on userId with ConsistencyLevel eventual, follows every nextLink, and returns the flattened device records or, when Graph refuses, an error.](diagrams/devices-by-user.svg)

`get_managed_devices_by_user` stands for the feature. Nothing is stored. List tools follow every `@odata.nextLink`, so a large tenant means many calls. The beta endpoint can change without notice; a field Microsoft renames there comes back `null`.

## 📚 Pages

- [Endpoints](endpoints.md) — the 9 tools, each with its mark, arguments, result, errors and an example call
- [Integration](integration.md) — the Microsoft services the feature calls and what happens when they fail

## 🔗 Related features

- [Users](../users/index.md) — reach for it to find the user id whose devices you want.
- [Conditional access](../conditional-access/index.md) — reach for it to see the policies that require compliant devices.
