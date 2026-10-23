"""MFA resource module for Microsoft Graph.

This module provides access to Microsoft Graph MFA-related resources.
"""

import logging
from typing import Any

from kiota_abstractions.base_request_configuration import RequestConfiguration
from msgraph.generated.users.item.user_item_request_builder import (
    UserItemRequestBuilder,
)

from ..utils.graph_client import GraphClient

logger = logging.getLogger(__name__)

USER_ODATA_TYPE = "#microsoft.graph.user"


async def get_mfa_status(graph_client: GraphClient, user_id: str) -> dict[str, Any]:
    """Get MFA status and methods for a specific user.

    Args:
        graph_client: GraphClient instance
        user_id: The unique identifier of the user.

    Returns:
        A dictionary containing MFA status and methods information.
    """
    client = graph_client.get_client()

    # Get user's company name
    query_params = UserItemRequestBuilder.UserItemRequestBuilderGetQueryParameters(
        select=["companyName"]
    )
    request_configuration = RequestConfiguration(query_parameters=query_params)
    user = await client.users.by_user_id(user_id).get(
        request_configuration=request_configuration
    )

    # Get MFA methods
    mfa_data = await client.users.by_user_id(user_id).authentication.methods.get()

    if not mfa_data:
        logger.warning(f"No MFA data found for user {user_id}")
        return None

    # Initialize MFA status object
    mfa_status = {
        "userPrincipalName": user_id,
        "mail": user.mail if user else None,
        "companyName": user.company_name if user else None,
        "mfaStatus": "Disabled",
        "methods": {
            "email": False,
            "fido2": False,
            "authenticatorApp": False,
            "password": False,
            "phone": False,
            "softwareOath": False,
            "temporaryAccessPass": False,
            "windowsHelloForBusiness": False,
        },
    }

    # Process each authentication method
    for method in mfa_data.value:
        method_type = method.odata_type

        if method_type == "#microsoft.graph.emailAuthenticationMethod":
            mfa_status["methods"]["email"] = True
            mfa_status["mfaStatus"] = "Enabled"
        elif method_type == "#microsoft.graph.fido2AuthenticationMethod":
            mfa_status["methods"]["fido2"] = True
            mfa_status["mfaStatus"] = "Enabled"
        elif (
            method_type == "#microsoft.graph.microsoftAuthenticatorAuthenticationMethod"
        ):
            mfa_status["methods"]["authenticatorApp"] = True
            mfa_status["mfaStatus"] = "Enabled"
        elif method_type == "#microsoft.graph.passwordAuthenticationMethod":
            mfa_status["methods"]["password"] = True
            if mfa_status["mfaStatus"] != "Enabled":
                mfa_status["mfaStatus"] = "Disabled"
        elif method_type == "#microsoft.graph.phoneAuthenticationMethod":
            mfa_status["methods"]["phone"] = True
            mfa_status["mfaStatus"] = "Enabled"
        elif method_type == "#microsoft.graph.softwareOathAuthenticationMethod":
            mfa_status["methods"]["softwareOath"] = True
            mfa_status["mfaStatus"] = "Enabled"
        elif method_type == "#microsoft.graph.temporaryAccessPassAuthenticationMethod":
            mfa_status["methods"]["temporaryAccessPass"] = True
            mfa_status["mfaStatus"] = "Enabled"
        elif (
            method_type
            == "#microsoft.graph.windowsHelloForBusinessAuthenticationMethod"
        ):
            mfa_status["methods"]["windowsHelloForBusiness"] = True
            mfa_status["mfaStatus"] = "Enabled"

    return mfa_status


async def get_group_mfa_status(
    graph_client: GraphClient, group_id: str
) -> tuple[list[dict[str, Any]], int, int]:
    """Get MFA status for every user member of a group.

    Reads every page of members. Members that are not users (devices, service
    principals, groups, unknown types) are skipped; nested groups are not expanded.

    Args:
        graph_client: GraphClient instance
        group_id: The unique identifier of the group.

    Returns:
        A tuple of the MFA statuses that could be read, the number of user
        members whose status could not be read, and the number of user members.
    """
    client = graph_client.get_client()
    members_builder = client.groups.by_group_id(group_id).members

    users = []
    page = await members_builder.get()
    while page:
        users += [m for m in page.value or [] if m.odata_type == USER_ODATA_TYPE]
        if not page.odata_next_link:
            break
        page = await members_builder.with_url(page.odata_next_link).get()

    if not users:
        logger.warning(f"No user members found in group {group_id}")
        return [], 0, 0

    mfa_statuses = []
    unreadable = 0
    for member in users:
        try:
            mfa_status = await get_mfa_status(graph_client, member.id)
        except Exception as e:  # noqa: BLE001 — an unreadable member is counted, the report goes on
            logger.error(f"Error processing member {member.id}: {e!s}")
            mfa_status = None
        if mfa_status:
            mfa_statuses.append(mfa_status)
        else:
            unreadable += 1

    return mfa_statuses, unreadable, len(users)
