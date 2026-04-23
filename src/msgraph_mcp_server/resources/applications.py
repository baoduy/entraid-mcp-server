"""Applications resource module for Microsoft Graph.

This module provides access to Microsoft Graph application resources (app registrations).
"""

import logging
from typing import Dict, List, Any, Optional
from ..utils.graph_client import GraphClient
from msgraph.generated.models.application import Application
from .service_principals import get_service_principal_by_app_id

logger = logging.getLogger(__name__)

async def list_applications(graph_client: GraphClient, limit: int = 100) -> List[Dict[str, Any]]:
    """List all applications (app registrations) in the tenant, with paging."""
    try:
        client = graph_client.get_client()
        response = await client.applications.get()
        applications = []
        if response and response.value:
            applications.extend(response.value)
        # Paging: fetch more if odata_next_link is present
        while response is not None and getattr(response, 'odata_next_link', None) and len(applications) < limit:
            response = await client.applications.with_url(response.odata_next_link).get()
            if response and response.value:
                applications.extend(response.value)
        formatted_apps = []
        for app in applications[:limit]:
            app_data = {
                'id': getattr(app, 'id', None),
                'appId': getattr(app, 'app_id', None),
                'displayName': getattr(app, 'display_name', None),
                'createdDateTime': app.created_date_time.isoformat() if getattr(app, 'created_date_time', None) else None,
                'signInAudience': getattr(app, 'sign_in_audience', None),
                'publisherDomain': getattr(app, 'publisher_domain', None),
                'tags': getattr(app, 'tags', None),
            }
            formatted_apps.append(app_data)
        return formatted_apps
    except Exception as e:
        logger.error(f"Error listing applications: {str(e)}")
        raise

async def get_application_by_id(graph_client: GraphClient, app_id: str) -> Optional[Dict[str, Any]]:
    """Get a specific application by its object ID, including appRoleAssignments and oauth2PermissionGrants from the corresponding service principal."""
    try:
        client = graph_client.get_client()
        app = await client.applications.by_application_id(app_id).get()
        if app:
            app_data = {
                'id': getattr(app, 'id', None),
                'appId': getattr(app, 'app_id', None),
                'displayName': getattr(app, 'display_name', None),
                'createdDateTime': app.created_date_time.isoformat() if getattr(app, 'created_date_time', None) else None,
                'signInAudience': getattr(app, 'sign_in_audience', None),
                'publisherDomain': getattr(app, 'publisher_domain', None),
                'tags': getattr(app, 'tags', None),
            }
            # Find the corresponding service principal by appId
            sp = await get_service_principal_by_app_id(graph_client, getattr(app, 'app_id', None))
            if sp:
                sp_id = getattr(sp, 'id', None)
                # Fetch appRoleAssignments and oauth2PermissionGrants using the same logic as in service_principals.py
                # Fetch appRoleAssignments
                app_role_assignments = []
                try:
                    response = await client.service_principals.by_service_principal_id(sp_id).app_role_assignments.get()
                    while response:
                        if response.value:
                            for assignment in response.value:
                                app_role_assignments.append({
                                    'id': getattr(assignment, 'id', None),
                                    'createdDateTime': getattr(assignment, 'created_date_time', None),
                                    'appRoleId': getattr(assignment, 'app_role_id', None),
                                    'principalDisplayName': getattr(assignment, 'principal_display_name', None),
                                    'principalId': getattr(assignment, 'principal_id', None),
                                    'principalType': getattr(assignment, 'principal_type', None),
                                    'resourceDisplayName': getattr(assignment, 'resource_display_name', None),
                                    'resourceId': getattr(assignment, 'resource_id', None),
                                })
                        if getattr(response, 'odata_next_link', None):
                            response = await client.service_principals.by_service_principal_id(sp_id).app_role_assignments.with_url(response.odata_next_link).get()
                        else:
                            break
                except Exception as e:
                    logger.warning(f"Error fetching appRoleAssignments for service principal {sp_id}: {str(e)}")
                app_data['appRoleAssignments'] = app_role_assignments

                # Fetch oauth2PermissionGrants
                oauth2_permission_grants = []
                try:
                    response = await client.service_principals.by_service_principal_id(sp_id).oauth2_permission_grants.get()
                    while response:
                        if response.value:
                            for grant in response.value:
                                oauth2_permission_grants.append({
                                    'id': getattr(grant, 'id', None),
                                    'clientId': getattr(grant, 'client_id', None),
                                    'consentType': getattr(grant, 'consent_type', None),
                                    'principalId': getattr(grant, 'principal_id', None),
                                    'resourceId': getattr(grant, 'resource_id', None),
                                    'scope': getattr(grant, 'scope', None),
                                })
                        if getattr(response, 'odata_next_link', None):
                            response = await client.service_principals.by_service_principal_id(sp_id).oauth2_permission_grants.with_url(response.odata_next_link).get()
                        else:
                            break
                except Exception as e:
                    logger.warning(f"Error fetching oauth2PermissionGrants for service principal {sp_id}: {str(e)}")
                app_data['oauth2PermissionGrants'] = oauth2_permission_grants
            else:
                app_data['appRoleAssignments'] = []
                app_data['oauth2PermissionGrants'] = []
            return app_data
        return None
    except Exception as e:
        logger.error(f"Error getting application {app_id}: {str(e)}")
        raise

async def create_application(graph_client: GraphClient, app_data: Dict[str, Any]) -> Dict[str, Any]:
    """Create a new application (app registration)."""
    try:
        client = graph_client.get_client()
        app = Application()
        # Set properties from app_data
        if 'displayName' in app_data:
            app.display_name = app_data['displayName']
        if 'signInAudience' in app_data:
            app.sign_in_audience = app_data['signInAudience']
        if 'tags' in app_data:
            app.tags = app_data['tags']
        if 'identifierUris' in app_data:
            app.identifier_uris = app_data['identifierUris']
        if 'web' in app_data:
            app.web = app_data['web']
        if 'api' in app_data:
            app.api = app_data['api']
        if 'requiredResourceAccess' in app_data:
            app.required_resource_access = app_data['requiredResourceAccess']
        new_app = await client.applications.post(app)
        if new_app:
            return {
                'id': getattr(new_app, 'id', None),
                'appId': getattr(new_app, 'app_id', None),
                'displayName': getattr(new_app, 'display_name', None),
                'createdDateTime': new_app.created_date_time.isoformat() if getattr(new_app, 'created_date_time', None) else None,
                'signInAudience': getattr(new_app, 'sign_in_audience', None),
                'publisherDomain': getattr(new_app, 'publisher_domain', None),
                'tags': getattr(new_app, 'tags', None),
            }
        raise Exception("Failed to create application")
    except Exception as e:
        logger.error(f"Error creating application: {str(e)}")
        raise

async def update_application(graph_client: GraphClient, app_id: str, app_data: Dict[str, Any]) -> Dict[str, Any]:
    """Update an existing application (app registration)."""
    try:
        client = graph_client.get_client()
        app = Application()
        # Set updatable properties from app_data
        if 'displayName' in app_data:
            app.display_name = app_data['displayName']
        if 'signInAudience' in app_data:
            app.sign_in_audience = app_data['signInAudience']
        if 'tags' in app_data:
            app.tags = app_data['tags']
        if 'identifierUris' in app_data:
            app.identifier_uris = app_data['identifierUris']
        if 'web' in app_data:
            app.web = app_data['web']
        if 'api' in app_data:
            app.api = app_data['api']
        if 'requiredResourceAccess' in app_data:
            app.required_resource_access = app_data['requiredResourceAccess']
        await client.applications.by_application_id(app_id).patch(app)
        # Return the updated application
        return await get_application_by_id(graph_client, app_id)
    except Exception as e:
        logger.error(f"Error updating application {app_id}: {str(e)}")
        raise

async def delete_application(graph_client: GraphClient, app_id: str) -> bool:
    """Delete an application (app registration) by its object ID."""
    try:
        client = graph_client.get_client()
        await client.applications.by_application_id(app_id).delete()
        return True
    except Exception as e:
        logger.error(f"Error deleting application {app_id}: {str(e)}")
        raise 