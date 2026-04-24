"""Intune managed device operations using the Microsoft Graph **beta** endpoint.

This module wraps the ``msgraph-beta-sdk`` to expose Intune device-management
functionality (managed devices, detected apps, compliance policy states,
configuration states, compliance policies, configuration profiles and device
categories). The beta endpoint is used because many Intune properties and
relationships are only available under ``https://graph.microsoft.com/beta``.
"""

import logging
from typing import Any, Dict, List, Optional

from kiota_abstractions.base_request_configuration import RequestConfiguration
from msgraph_beta.generated.device_management.managed_devices.managed_devices_request_builder import (
    ManagedDevicesRequestBuilder,
)

from ..utils.graph_client import GraphClient

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def _enum_value(value: Any) -> Any:
    """Return ``value.value`` for enums, otherwise the value itself."""
    return getattr(value, "value", value) if value is not None else None


def _iso(value: Any) -> Optional[str]:
    """Return ``value.isoformat()`` when supported, else ``None``."""
    return value.isoformat() if value is not None and hasattr(value, "isoformat") else value


def _format_managed_device(device: Any) -> Dict[str, Any]:
    """Flatten a beta ``ManagedDevice`` into a JSON-serializable dict.

    Only the most useful fields for Intune administration are included to
    keep responses compact; callers that need a single device can use
    :func:`get_managed_device_by_id` which returns the full payload.
    """
    hw = getattr(device, "hardware_information", None)
    hardware = None
    if hw is not None:
        hardware = {
            "serialNumber": getattr(hw, "serial_number", None),
            "manufacturer": getattr(hw, "manufacturer", None),
            "model": getattr(hw, "model", None),
            "totalStorageSpace": getattr(hw, "total_storage_space", None),
            "freeStorageSpace": getattr(hw, "free_storage_space", None),
            "imei": getattr(hw, "imei", None),
            "meid": getattr(hw, "meid", None),
            "operatingSystemLanguage": getattr(hw, "operating_system_language", None),
            "wifiMac": getattr(hw, "wifi_mac", None),
        }

    return {
        "id": getattr(device, "id", None),
        "deviceName": getattr(device, "device_name", None),
        "managedDeviceName": getattr(device, "managed_device_name", None),
        "userId": getattr(device, "user_id", None),
        "userPrincipalName": getattr(device, "user_principal_name", None),
        "userDisplayName": getattr(device, "user_display_name", None),
        "emailAddress": getattr(device, "email_address", None),
        "operatingSystem": getattr(device, "operating_system", None),
        "osVersion": getattr(device, "os_version", None),
        "manufacturer": getattr(device, "manufacturer", None),
        "model": getattr(device, "model", None),
        "serialNumber": getattr(device, "serial_number", None),
        "deviceType": _enum_value(getattr(device, "device_type", None)),
        "chassisType": _enum_value(getattr(device, "chassis_type", None)),
        "joinType": _enum_value(getattr(device, "join_type", None)),
        "azureADDeviceId": getattr(device, "azure_a_d_device_id", None),
        "azureADRegistered": getattr(device, "azure_a_d_registered", None),
        "autopilotEnrolled": getattr(device, "autopilot_enrolled", None),
        "enrolledDateTime": _iso(getattr(device, "enrolled_date_time", None)),
        "enrollmentProfileName": getattr(device, "enrollment_profile_name", None),
        "enrolledByUserPrincipalName": getattr(device, "enrolled_by_user_principal_name", None),
        "managementAgent": _enum_value(getattr(device, "management_agent", None)),
        "managementState": _enum_value(getattr(device, "management_state", None)),
        "ownerType": _enum_value(getattr(device, "owner_type", None)),
        "managedDeviceOwnerType": _enum_value(getattr(device, "managed_device_owner_type", None)),
        "complianceState": _enum_value(getattr(device, "compliance_state", None)),
        "complianceGracePeriodExpirationDateTime": _iso(
            getattr(device, "compliance_grace_period_expiration_date_time", None)
        ),
        "jailBroken": getattr(device, "jail_broken", None),
        "isEncrypted": getattr(device, "is_encrypted", None),
        "isSupervised": getattr(device, "is_supervised", None),
        "deviceEnrollmentType": _enum_value(getattr(device, "device_enrollment_type", None)),
        "deviceRegistrationState": _enum_value(getattr(device, "device_registration_state", None)),
        "deviceCategoryDisplayName": getattr(device, "device_category_display_name", None),
        "lastSyncDateTime": _iso(getattr(device, "last_sync_date_time", None)),
        "freeStorageSpaceInBytes": getattr(device, "free_storage_space_in_bytes", None),
        "totalStorageSpaceInBytes": getattr(device, "total_storage_space_in_bytes", None),
        "physicalMemoryInBytes": getattr(device, "physical_memory_in_bytes", None),
        "processorArchitecture": _enum_value(getattr(device, "processor_architecture", None)),
        "wiFiMacAddress": getattr(device, "wi_fi_mac_address", None),
        "ethernetMacAddress": getattr(device, "ethernet_mac_address", None),
        "imei": getattr(device, "imei", None),
        "meid": getattr(device, "meid", None),
        "iccid": getattr(device, "iccid", None),
        "udid": getattr(device, "udid", None),
        "notes": getattr(device, "notes", None),
        "partnerReportedThreatState": _enum_value(
            getattr(device, "partner_reported_threat_state", None)
        ),
        "windowsActiveMalwareCount": getattr(device, "windows_active_malware_count", None),
        "windowsRemediatedMalwareCount": getattr(device, "windows_remediated_malware_count", None),
        "roleScopeTagIds": getattr(device, "role_scope_tag_ids", None),
        "hardwareInformation": hardware,
    }


async def _collect_paged(initial_response, request_builder_factory, request_configuration):
    """Collect all items from a paged Graph response.

    ``request_builder_factory`` is a callable accepting a URL that returns a
    request builder whose ``get`` method can be awaited.
    """
    items: List[Any] = []
    response = initial_response
    if response and getattr(response, "value", None):
        items.extend(response.value)
    while response is not None and getattr(response, "odata_next_link", None):
        response = await request_builder_factory(response.odata_next_link).get(
            request_configuration=request_configuration
        )
        if response and getattr(response, "value", None):
            items.extend(response.value)
    return items


# ---------------------------------------------------------------------------
# Managed devices
# ---------------------------------------------------------------------------
async def get_all_managed_devices(
    graph_client: GraphClient, filter_os: Optional[str] = None
) -> List[Dict[str, Any]]:
    """Get all Intune managed devices (optionally filtered by OS)."""
    try:
        client = graph_client.get_beta_client()
        query_params = ManagedDevicesRequestBuilder.ManagedDevicesRequestBuilderGetQueryParameters()
        if filter_os:
            query_params.filter = f"operatingSystem eq '{filter_os}'"
        request_configuration = RequestConfiguration(query_parameters=query_params)
        request_configuration.headers.add("ConsistencyLevel", "eventual")

        response = await client.device_management.managed_devices.get(
            request_configuration=request_configuration
        )
        devices = await _collect_paged(
            response,
            lambda url: client.device_management.managed_devices.with_url(url),
            request_configuration,
        )
        return [_format_managed_device(d) for d in devices]
    except Exception as e:
        logger.error(f"Error fetching all managed devices: {str(e)}")
        raise


async def get_managed_devices_by_user(
    graph_client: GraphClient, user_id: str
) -> List[Dict[str, Any]]:
    """Get all Intune managed devices for a specific userId."""
    try:
        client = graph_client.get_beta_client()
        query_params = ManagedDevicesRequestBuilder.ManagedDevicesRequestBuilderGetQueryParameters(
            filter=f"userId eq '{user_id}'"
        )
        request_configuration = RequestConfiguration(query_parameters=query_params)
        request_configuration.headers.add("ConsistencyLevel", "eventual")

        response = await client.device_management.managed_devices.get(
            request_configuration=request_configuration
        )
        devices = await _collect_paged(
            response,
            lambda url: client.device_management.managed_devices.with_url(url),
            request_configuration,
        )
        return [_format_managed_device(d) for d in devices]
    except Exception as e:
        logger.error(f"Error fetching managed devices for user {user_id}: {str(e)}")
        raise


async def get_managed_device_by_id(
    graph_client: GraphClient, device_id: str
) -> Optional[Dict[str, Any]]:
    """Get the full details of a single Intune managed device."""
    try:
        client = graph_client.get_beta_client()
        device = await client.device_management.managed_devices.by_managed_device_id(
            device_id
        ).get()
        if device is None:
            return None
        return _format_managed_device(device)
    except Exception as e:
        logger.error(f"Error fetching managed device {device_id}: {str(e)}")
        raise


# ---------------------------------------------------------------------------
# Device relationships
# ---------------------------------------------------------------------------
async def get_detected_apps_for_device(
    graph_client: GraphClient, device_id: str
) -> List[Dict[str, Any]]:
    """Get applications detected on a specific managed device."""
    try:
        client = graph_client.get_beta_client()
        device_item = client.device_management.managed_devices.by_managed_device_id(device_id)
        response = await device_item.detected_apps.get()
        apps: List[Any] = []
        if response and getattr(response, "value", None):
            apps.extend(response.value)
        while response is not None and getattr(response, "odata_next_link", None):
            response = await device_item.detected_apps.with_url(
                response.odata_next_link
            ).get()
            if response and getattr(response, "value", None):
                apps.extend(response.value)
        return [
            {
                "id": getattr(a, "id", None),
                "displayName": getattr(a, "display_name", None),
                "version": getattr(a, "version", None),
                "sizeInByte": getattr(a, "size_in_byte", None),
                "deviceCount": getattr(a, "device_count", None),
                "platform": _enum_value(getattr(a, "platform", None)),
                "publisher": getattr(a, "publisher", None),
            }
            for a in apps
        ]
    except Exception as e:
        logger.error(f"Error fetching detected apps for device {device_id}: {str(e)}")
        raise


async def get_device_compliance_policy_states(
    graph_client: GraphClient, device_id: str
) -> List[Dict[str, Any]]:
    """Get compliance policy states for a specific managed device."""
    try:
        client = graph_client.get_beta_client()
        device_item = client.device_management.managed_devices.by_managed_device_id(device_id)
        response = await device_item.device_compliance_policy_states.get()
        items: List[Any] = []
        if response and getattr(response, "value", None):
            items.extend(response.value)
        while response is not None and getattr(response, "odata_next_link", None):
            response = await device_item.device_compliance_policy_states.with_url(
                response.odata_next_link
            ).get()
            if response and getattr(response, "value", None):
                items.extend(response.value)
        return [
            {
                "id": getattr(s, "id", None),
                "displayName": getattr(s, "display_name", None),
                "state": _enum_value(getattr(s, "state", None)),
                "version": getattr(s, "version", None),
                "platformType": _enum_value(getattr(s, "platform_type", None)),
                "settingCount": getattr(s, "setting_count", None),
                "userPrincipalName": getattr(s, "user_principal_name", None),
                "userId": getattr(s, "user_id", None),
            }
            for s in items
        ]
    except Exception as e:
        logger.error(
            f"Error fetching compliance policy states for device {device_id}: {str(e)}"
        )
        raise


async def get_device_configuration_states(
    graph_client: GraphClient, device_id: str
) -> List[Dict[str, Any]]:
    """Get configuration-profile states for a specific managed device."""
    try:
        client = graph_client.get_beta_client()
        device_item = client.device_management.managed_devices.by_managed_device_id(device_id)
        response = await device_item.device_configuration_states.get()
        items: List[Any] = []
        if response and getattr(response, "value", None):
            items.extend(response.value)
        while response is not None and getattr(response, "odata_next_link", None):
            response = await device_item.device_configuration_states.with_url(
                response.odata_next_link
            ).get()
            if response and getattr(response, "value", None):
                items.extend(response.value)
        return [
            {
                "id": getattr(s, "id", None),
                "displayName": getattr(s, "display_name", None),
                "state": _enum_value(getattr(s, "state", None)),
                "version": getattr(s, "version", None),
                "platformType": _enum_value(getattr(s, "platform_type", None)),
                "settingCount": getattr(s, "setting_count", None),
                "userPrincipalName": getattr(s, "user_principal_name", None),
                "userId": getattr(s, "user_id", None),
            }
            for s in items
        ]
    except Exception as e:
        logger.error(
            f"Error fetching configuration states for device {device_id}: {str(e)}"
        )
        raise


# ---------------------------------------------------------------------------
# Intune catalog resources (policies, profiles, categories)
# ---------------------------------------------------------------------------
async def get_device_compliance_policies(
    graph_client: GraphClient,
) -> List[Dict[str, Any]]:
    """List all Intune device compliance policies."""
    try:
        client = graph_client.get_beta_client()
        response = await client.device_management.device_compliance_policies.get()
        items: List[Any] = []
        if response and getattr(response, "value", None):
            items.extend(response.value)
        while response is not None and getattr(response, "odata_next_link", None):
            response = await client.device_management.device_compliance_policies.with_url(
                response.odata_next_link
            ).get()
            if response and getattr(response, "value", None):
                items.extend(response.value)
        return [
            {
                "id": getattr(p, "id", None),
                "displayName": getattr(p, "display_name", None),
                "description": getattr(p, "description", None),
                "version": getattr(p, "version", None),
                "odataType": getattr(p, "odata_type", None),
                "createdDateTime": _iso(getattr(p, "created_date_time", None)),
                "lastModifiedDateTime": _iso(getattr(p, "last_modified_date_time", None)),
                "roleScopeTagIds": getattr(p, "role_scope_tag_ids", None),
            }
            for p in items
        ]
    except Exception as e:
        logger.error(f"Error fetching device compliance policies: {str(e)}")
        raise


async def get_device_configurations(
    graph_client: GraphClient,
) -> List[Dict[str, Any]]:
    """List all Intune device configuration profiles."""
    try:
        client = graph_client.get_beta_client()
        response = await client.device_management.device_configurations.get()
        items: List[Any] = []
        if response and getattr(response, "value", None):
            items.extend(response.value)
        while response is not None and getattr(response, "odata_next_link", None):
            response = await client.device_management.device_configurations.with_url(
                response.odata_next_link
            ).get()
            if response and getattr(response, "value", None):
                items.extend(response.value)
        return [
            {
                "id": getattr(c, "id", None),
                "displayName": getattr(c, "display_name", None),
                "description": getattr(c, "description", None),
                "version": getattr(c, "version", None),
                "odataType": getattr(c, "odata_type", None),
                "createdDateTime": _iso(getattr(c, "created_date_time", None)),
                "lastModifiedDateTime": _iso(getattr(c, "last_modified_date_time", None)),
                "roleScopeTagIds": getattr(c, "role_scope_tag_ids", None),
            }
            for c in items
        ]
    except Exception as e:
        logger.error(f"Error fetching device configurations: {str(e)}")
        raise


async def get_device_categories(graph_client: GraphClient) -> List[Dict[str, Any]]:
    """List all Intune device categories."""
    try:
        client = graph_client.get_beta_client()
        response = await client.device_management.device_categories.get()
        items: List[Any] = []
        if response and getattr(response, "value", None):
            items.extend(response.value)
        while response is not None and getattr(response, "odata_next_link", None):
            response = await client.device_management.device_categories.with_url(
                response.odata_next_link
            ).get()
            if response and getattr(response, "value", None):
                items.extend(response.value)
        return [
            {
                "id": getattr(c, "id", None),
                "displayName": getattr(c, "display_name", None),
                "description": getattr(c, "description", None),
                "roleScopeTagIds": getattr(c, "role_scope_tag_ids", None),
            }
            for c in items
        ]
    except Exception as e:
        logger.error(f"Error fetching device categories: {str(e)}")
        raise
