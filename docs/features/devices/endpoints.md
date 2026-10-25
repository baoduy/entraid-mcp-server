# Devices (Intune) — Endpoints

Part of [Devices (Intune)](index.md).

Each endpoint is an MCP tool. An AI client calls it with a JSON-RPC `tools/call` message over the server's standard input; the server answers on standard output.

## 🔌 Endpoints

| Tool | Mark | Purpose | Auth |
|---|---|---|---|
| [`get_all_managed_devices`](#get_all_managed_devices) | read-only | List every Intune managed device, optionally only one operating system. | signed-in identity |
| [`get_managed_devices_by_user`](#get_managed_devices_by_user) | read-only | List the Intune managed devices of one user. | signed-in identity |
| [`get_managed_device_by_id`](#get_managed_device_by_id) | read-only | Get one Intune managed device: hardware, compliance, enrolment. | signed-in identity |
| [`get_detected_apps_for_device`](#get_detected_apps_for_device) | read-only | List the applications Intune detected on a device. | signed-in identity |
| [`get_device_compliance_policy_states`](#get_device_compliance_policy_states) | read-only | List how a device stands against each compliance policy. | signed-in identity |
| [`get_device_configuration_states`](#get_device_configuration_states) | read-only | List how a device stands against each configuration profile. | signed-in identity |
| [`get_device_compliance_policies`](#get_device_compliance_policies) | read-only | List the tenant's Intune compliance policies. | signed-in identity |
| [`get_device_configurations`](#get_device_configurations) | read-only | List the tenant's Intune configuration profiles. | signed-in identity |
| [`get_device_categories`](#get_device_categories) | read-only | List the tenant's Intune device categories. | signed-in identity |

### `get_all_managed_devices`

List every Intune managed device, optionally only one operating system.

- **Mark:** read-only
- **Auth:** the identity the server signed in as ([configuration](../../configuration.md)); Microsoft Graph checks its permissions, the server checks none
- **Microsoft Graph:** Beta `GET /deviceManagement/managedDevices` (with `$filter` when `filter_os` is set), `ConsistencyLevel: eventual`, every page.
- **Request:**

| Field | Type | Required | Default | Rules |
|---|---|---|---|---|
| `filter_os` | string or null | no | `null` | When set: `$filter=operatingSystem eq '<filter_os>'`; a `'` is doubled |

- **Response:** Each device is the flattened managed-device record: 49 fields plus `hardwareInformation` (see `get_managed_device_by_id`).

  Example `structuredContent`, captured from the test suite's in-memory Graph:

```json
[
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
]
```

- **Errors:**

| Result | Text | When |
|---|---|---|
| error result | `Error calling tool 'get_all_managed_devices': <reason>`; notification `Error fetching all managed devices: <reason>` | Microsoft Graph or sign-in fails |

- **Example:**

```json
{"jsonrpc": "2.0", "id": 1, "method": "tools/call", "params": {"name": "get_all_managed_devices", "arguments": {}}}
```

### `get_managed_devices_by_user`

List the Intune managed devices of one user.

- **Mark:** read-only
- **Auth:** the identity the server signed in as ([configuration](../../configuration.md)); Microsoft Graph checks its permissions, the server checks none
- **Microsoft Graph:** Beta `GET /deviceManagement/managedDevices?$filter=userId eq '<user_id>'`, `ConsistencyLevel: eventual`, every page.
- **Request:**

| Field | Type | Required | Default | Rules |
|---|---|---|---|---|
| `user_id` | string | yes | — | — |

- **Response:** The same device records as `get_all_managed_devices`.

  Example `structuredContent`, captured from the test suite's in-memory Graph:

```json
[
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
]
```

- **Errors:**

| Result | Text | When |
|---|---|---|
| error result | `Error calling tool 'get_managed_devices_by_user': <reason>`; notification `Error fetching managed devices for user <user_id>: <reason>` | Microsoft Graph or sign-in fails |

- **Example:**

```json
{"jsonrpc": "2.0", "id": 1, "method": "tools/call", "params": {"name": "get_managed_devices_by_user", "arguments": {"user_id": "aoife@contoso.com"}}}
```

### `get_managed_device_by_id`

Get one Intune managed device: hardware, compliance, enrolment.

- **Mark:** read-only
- **Auth:** the identity the server signed in as ([configuration](../../configuration.md)); Microsoft Graph checks its permissions, the server checks none
- **Microsoft Graph:** Beta `GET /deviceManagement/managedDevices/{device_id}`.
- **Request:**

| Field | Type | Required | Default | Rules |
|---|---|---|---|---|
| `device_id` | string | yes | — | — |

- **Response:** One device record, or `null` with the warning `Managed device with ID <device_id> not found.` when Graph returns no object.

  Example `structuredContent`, captured from the test suite's in-memory Graph:

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

- **Errors:**

| Result | Text | When |
|---|---|---|
| error result | `Error calling tool 'get_managed_device_by_id': <reason>`; notification `Error fetching managed device <device_id>: <reason>` | Microsoft Graph or sign-in fails |

- **Example:**

```json
{"jsonrpc": "2.0", "id": 1, "method": "tools/call", "params": {"name": "get_managed_device_by_id", "arguments": {"device_id": "device-1"}}}
```

### `get_detected_apps_for_device`

List the applications Intune detected on a device.

- **Mark:** read-only
- **Auth:** the identity the server signed in as ([configuration](../../configuration.md)); Microsoft Graph checks its permissions, the server checks none
- **Microsoft Graph:** Beta `GET /deviceManagement/managedDevices/{device_id}/detectedApps`, every page.
- **Request:**

| Field | Type | Required | Default | Rules |
|---|---|---|---|---|
| `device_id` | string | yes | — | — |

- **Response:** Each app has `id`, `displayName`, `version`, `sizeInByte`, `deviceCount`, `platform`, `publisher`.

  Example `structuredContent`, captured from the test suite's in-memory Graph:

```json
[
  {
    "id": "app-1",
    "displayName": "Contoso VPN",
    "version": "2.1.0",
    "sizeInByte": null,
    "deviceCount": null,
    "platform": null,
    "publisher": null
  }
]
```

- **Errors:**

| Result | Text | When |
|---|---|---|
| error result | `Error calling tool 'get_detected_apps_for_device': <reason>`; notification `Error fetching detected apps for device <device_id>: <reason>` | Microsoft Graph or sign-in fails |

- **Example:**

```json
{"jsonrpc": "2.0", "id": 1, "method": "tools/call", "params": {"name": "get_detected_apps_for_device", "arguments": {"device_id": "device-1"}}}
```

### `get_device_compliance_policy_states`

List how a device stands against each compliance policy.

- **Mark:** read-only
- **Auth:** the identity the server signed in as ([configuration](../../configuration.md)); Microsoft Graph checks its permissions, the server checks none
- **Microsoft Graph:** Beta `GET /deviceManagement/managedDevices/{device_id}/deviceCompliancePolicyStates`, every page.
- **Request:**

| Field | Type | Required | Default | Rules |
|---|---|---|---|---|
| `device_id` | string | yes | — | — |

- **Response:** Each state has `id`, `displayName`, `state`, `version`, `platformType`, `settingCount`, `userPrincipalName`, `userId`.

  Example `structuredContent`, captured from the test suite's in-memory Graph:

```json
[
  {
    "id": "state-1",
    "displayName": "Windows baseline",
    "state": null,
    "version": null,
    "platformType": null,
    "settingCount": null,
    "userPrincipalName": null,
    "userId": null
  }
]
```

- **Errors:**

| Result | Text | When |
|---|---|---|
| error result | `Error calling tool 'get_device_compliance_policy_states': <reason>`; notification `Error fetching compliance policy states for device <device_id>: <reason>` | Microsoft Graph or sign-in fails |

- **Example:**

```json
{"jsonrpc": "2.0", "id": 1, "method": "tools/call", "params": {"name": "get_device_compliance_policy_states", "arguments": {"device_id": "device-1"}}}
```

### `get_device_configuration_states`

List how a device stands against each configuration profile.

- **Mark:** read-only
- **Auth:** the identity the server signed in as ([configuration](../../configuration.md)); Microsoft Graph checks its permissions, the server checks none
- **Microsoft Graph:** Beta `GET /deviceManagement/managedDevices/{device_id}/deviceConfigurationStates`, every page.
- **Request:**

| Field | Type | Required | Default | Rules |
|---|---|---|---|---|
| `device_id` | string | yes | — | — |

- **Response:** The same fields as `get_device_compliance_policy_states`.

  Example `structuredContent`, captured from the test suite's in-memory Graph:

```json
[
  {
    "id": "state-1",
    "displayName": "Wi-Fi profile",
    "state": null,
    "version": null,
    "platformType": null,
    "settingCount": null,
    "userPrincipalName": null,
    "userId": null
  }
]
```

- **Errors:**

| Result | Text | When |
|---|---|---|
| error result | `Error calling tool 'get_device_configuration_states': <reason>`; notification `Error fetching configuration states for device <device_id>: <reason>` | Microsoft Graph or sign-in fails |

- **Example:**

```json
{"jsonrpc": "2.0", "id": 1, "method": "tools/call", "params": {"name": "get_device_configuration_states", "arguments": {"device_id": "device-1"}}}
```

### `get_device_compliance_policies`

List the tenant's Intune compliance policies.

- **Mark:** read-only
- **Auth:** the identity the server signed in as ([configuration](../../configuration.md)); Microsoft Graph checks its permissions, the server checks none
- **Microsoft Graph:** Beta `GET /deviceManagement/deviceCompliancePolicies`, every page.
- **Request:**

  No arguments.

- **Response:** Each policy has `id`, `displayName`, `description`, `version`, `odataType`, `createdDateTime`, `lastModifiedDateTime`, `roleScopeTagIds`.

  Example `structuredContent`, captured from the test suite's in-memory Graph:

```json
[
  {
    "id": "compliance-1",
    "displayName": "Windows baseline",
    "description": null,
    "version": null,
    "odataType": "#microsoft.graph.windows10CompliancePolicy",
    "createdDateTime": null,
    "lastModifiedDateTime": null,
    "roleScopeTagIds": null
  }
]
```

- **Errors:**

| Result | Text | When |
|---|---|---|
| error result | `Error calling tool 'get_device_compliance_policies': <reason>`; notification `Error fetching device compliance policies: <reason>` | Microsoft Graph or sign-in fails |

- **Example:**

```json
{"jsonrpc": "2.0", "id": 1, "method": "tools/call", "params": {"name": "get_device_compliance_policies", "arguments": {}}}
```

### `get_device_configurations`

List the tenant's Intune configuration profiles.

- **Mark:** read-only
- **Auth:** the identity the server signed in as ([configuration](../../configuration.md)); Microsoft Graph checks its permissions, the server checks none
- **Microsoft Graph:** Beta `GET /deviceManagement/deviceConfigurations`, every page.
- **Request:**

  No arguments.

- **Response:** The same fields as `get_device_compliance_policies`.

  Example `structuredContent`, captured from the test suite's in-memory Graph:

```json
[
  {
    "id": "configuration-1",
    "displayName": "Wi-Fi profile",
    "description": null,
    "version": null,
    "odataType": "#microsoft.graph.windows10GeneralConfiguration",
    "createdDateTime": null,
    "lastModifiedDateTime": null,
    "roleScopeTagIds": null
  }
]
```

- **Errors:**

| Result | Text | When |
|---|---|---|
| error result | `Error calling tool 'get_device_configurations': <reason>`; notification `Error fetching device configurations: <reason>` | Microsoft Graph or sign-in fails |

- **Example:**

```json
{"jsonrpc": "2.0", "id": 1, "method": "tools/call", "params": {"name": "get_device_configurations", "arguments": {}}}
```

### `get_device_categories`

List the tenant's Intune device categories.

- **Mark:** read-only
- **Auth:** the identity the server signed in as ([configuration](../../configuration.md)); Microsoft Graph checks its permissions, the server checks none
- **Microsoft Graph:** Beta `GET /deviceManagement/deviceCategories`, every page.
- **Request:**

  No arguments.

- **Response:** Each category has `id`, `displayName`, `description`, `roleScopeTagIds`.

  Example `structuredContent`, captured from the test suite's in-memory Graph:

```json
[
  {
    "id": "category-1",
    "displayName": "Corporate",
    "description": null,
    "roleScopeTagIds": null
  }
]
```

- **Errors:**

| Result | Text | When |
|---|---|---|
| error result | `Error calling tool 'get_device_categories': <reason>`; notification `Error fetching device categories: <reason>` | Microsoft Graph or sign-in fails |

- **Example:**

```json
{"jsonrpc": "2.0", "id": 1, "method": "tools/call", "params": {"name": "get_device_categories", "arguments": {}}}
```

## ⚠️ Errors & limits

A failed call reaches the AI client twice: one error notification (`<message>: <reason>`, sent through the MCP log channel) and the error result `Error calling tool '<tool>': <reason>` with `isError: true`. The reason is the text of the exception Microsoft Graph, the SDK or the server raised. The server sets no timeout and no retry of its own and keeps no state between calls. Results that report a known case — `already_exists`, `not_found`, a dynamic group's `status: error`, the permission helper's `status: error` — are normal results, not errors.

- All nine tools call the Microsoft Graph **beta** endpoint (`msgraph-beta-sdk`), which Microsoft can change without notice.
- List tools read every page; a large tenant means many Graph calls and a large result.
- Enum values (device type, compliance state and so on) are returned as their string values.
