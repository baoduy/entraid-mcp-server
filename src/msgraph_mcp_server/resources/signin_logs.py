"""Sign-in logs resource module for Microsoft Graph.

This module provides access to Microsoft Graph sign-in logs.
"""

import logging
from datetime import datetime, timedelta, timezone
from typing import Any

from kiota_abstractions.base_request_configuration import RequestConfiguration
from msgraph.generated.audit_logs.sign_ins.sign_ins_request_builder import (
    SignInsRequestBuilder,
)

from ..utils.graph_client import GraphClient
from ..utils.odata import odata_string_literal

logger = logging.getLogger(__name__)


async def get_user_sign_in_logs(
    graph_client: GraphClient, user_id: str, days: int = 7
) -> list[dict[str, Any]]:
    """Get sign-in logs for a specific user within the last N days.

    Args:
        graph_client: GraphClient instance
        user_id: The unique identifier of the user.
        days: The number of past days to retrieve logs for (default: 7).

    Returns:
        A list of dictionaries, each representing a sign-in log event.
    """
    client = graph_client.get_client()

    # Calculate date range
    end_date = datetime.now(timezone.utc)
    start_date = end_date - timedelta(days=days)

    # Format dates for query using the exact format from documentation
    start_date_str = start_date.strftime("%Y-%m-%dT%H:%M:%SZ")
    end_date_str = end_date.strftime("%Y-%m-%dT%H:%M:%SZ")

    # Define the OData filter query with proper formatting
    filter_query = f"createdDateTime ge {start_date_str} and createdDateTime le {end_date_str} and userId eq {odata_string_literal(user_id)}"

    logger.info(f"Fetching sign-in logs for user ID: {user_id}")
    logger.info(f"Date range: {start_date_str} to {end_date_str}")
    logger.debug(f"Filter query: {filter_query}")

    # Set up query parameters using SignInsRequestBuilder
    query_params = SignInsRequestBuilder.SignInsRequestBuilderGetQueryParameters(
        filter=filter_query,
        orderby=["createdDateTime desc"],
        top=1000,  # Increased from default to get more logs
    )

    # Create request configuration
    request_configuration = RequestConfiguration(query_parameters=query_params)
    request_configuration.headers.add("ConsistencyLevel", "eventual")

    # Execute the request
    response = await client.audit_logs.sign_ins.get(
        request_configuration=request_configuration
    )
    sign_ins = []
    if response and response.value:
        sign_ins.extend(response.value)
    while response is not None and getattr(response, "odata_next_link", None):
        response = await client.audit_logs.sign_ins.with_url(
            response.odata_next_link
        ).get(request_configuration=request_configuration)
        if response and response.value:
            sign_ins.extend(response.value)

    formatted_logs = []
    if sign_ins:
        logger.info(f"Found {len(sign_ins)} sign-in records")

        for log in sign_ins:
            # Format each log entry with comprehensive fields
            log_data = {
                "id": log.id,
                "createdDateTime": log.created_date_time.isoformat()
                if log.created_date_time
                else None,
                "userId": log.user_id,
                "userDisplayName": log.user_display_name,
                "userPrincipalName": log.user_principal_name,
                "appDisplayName": log.app_display_name,
                "appId": log.app_id,
                "ipAddress": log.ip_address,
                "clientAppUsed": log.client_app_used,
                "correlationId": log.correlation_id,
                "isInteractive": log.is_interactive,
                "resourceDisplayName": log.resource_display_name,
                "status": {
                    "errorCode": log.status.error_code if log.status else None,
                    "failureReason": log.status.failure_reason if log.status else None,
                    "additionalDetails": log.status.additional_details
                    if log.status
                    else None,
                },
                "riskInformation": {
                    "riskDetail": log.risk_detail,
                    "riskLevelAggregated": log.risk_level_aggregated,
                    "riskLevelDuringSignIn": log.risk_level_during_sign_in,
                    "riskState": log.risk_state,
                    "riskEventTypes": log.risk_event_types_v2
                    if hasattr(log, "risk_event_types_v2")
                    else [],
                },
            }

            # Add device details if available
            if hasattr(log, "device_detail") and log.device_detail:
                device = log.device_detail
                log_data["deviceDetail"] = {
                    "deviceId": device.device_id,
                    "displayName": device.display_name,
                    "operatingSystem": device.operating_system,
                    "browser": device.browser,
                    "isCompliant": device.is_compliant,
                    "isManaged": device.is_managed,
                    "trustType": device.trust_type,
                }

            # Add location if available
            if hasattr(log, "location") and log.location:
                location = log.location
                log_data["location"] = {
                    "city": location.city,
                    "state": location.state,
                    "countryOrRegion": location.country_or_region,
                    "coordinates": None,
                }

                # Add coordinates if available
                if hasattr(location, "geo_coordinates") and location.geo_coordinates:
                    log_data["location"]["coordinates"] = {
                        "latitude": location.geo_coordinates.latitude,
                        "longitude": location.geo_coordinates.longitude,
                    }

            formatted_logs.append(log_data)
    else:
        logger.info(
            f"No sign-in logs found for user {user_id} in the last {days} days."
        )

    return formatted_logs
