"""Unit facts the DRK-2430 Build adds beyond the frozen acceptance tests (brief §6a D4–D7, rows 6–8).

The start-up acceptance tests run the server in a fresh interpreter, so these tests also pin the
same rules in-process: the order inside ``main``, when the credential settings are read, and the
version fallback.
"""

import importlib.metadata
import logging
import runpy
from types import SimpleNamespace

import pytest
from conftest import GraphError
from msgraph.generated.models.application import Application
from msgraph.generated.models.directory_role import DirectoryRole
from msgraph.generated.models.group import Group
from msgraph.generated.models.service_principal import ServicePrincipal
from msgraph.generated.models.user import User
from msgraph_beta.generated.models.detected_app import DetectedApp
from msgraph_beta.generated.models.device_category import DeviceCategory
from msgraph_beta.generated.models.device_compliance_policy_state import (
    DeviceCompliancePolicyState,
)
from msgraph_beta.generated.models.device_configuration_state import (
    DeviceConfigurationState,
)
from msgraph_beta.generated.models.managed_device import ManagedDevice
from msgraph_beta.generated.models.windows10_compliance_policy import (
    Windows10CompliancePolicy,
)
from msgraph_beta.generated.models.windows10_general_configuration import (
    Windows10GeneralConfiguration,
)
from test_group_members import ALL_STAFF, AOIFE, all_staff, use_tool
from test_tool_errors import call_once

import msgraph_mcp_server
from msgraph_mcp_server import server
from msgraph_mcp_server.auth import graph_auth
from msgraph_mcp_server.resources import conditional_access, groups
from msgraph_mcp_server.utils.graph_client import GraphClient

# -- D5: create_group adds its listed members through the member add -----------------------------


async def test_a_listed_member_already_in_the_new_group_does_not_fail_the_creation(
    graph,
):
    result = await groups.create_group(
        graph,
        {
            "displayName": "All Staff",
            "mailNickname": "allstaff",
            "members": ["user-0001", AOIFE, "user-0001"],
        },
    )

    assert result["id"] == "new-0001"
    members = graph.groups.by_group_id("new-0001").members
    assert members.ids() == ["user-0001", AOIFE]
    assert len(graph.requests_to("groups/new-0001/members/$ref")) == 2
    member_checks = [
        r for r in graph.requests_to("groups/new-0001/members") if r.method == "get"
    ]
    assert len(member_checks) == 2


# -- D4: the membership check -------------------------------------------------------------------


@pytest.mark.parametrize(
    "tool, notification",
    [
        (
            "add_group_member",
            f"Error adding member {AOIFE} to group {ALL_STAFF}: Service unavailable",
        ),
        (
            "remove_group_member",
            f"Error removing member {AOIFE} from group {ALL_STAFF}: Service unavailable",
        ),
    ],
)
async def test_a_failed_membership_check_fails_the_call(
    graph, monkeypatch, tool, notification
):
    members = all_staff(graph, ["user-0001"])

    async def unavailable(request_configuration=None):
        raise GraphError("Service unavailable")

    monkeypatch.setattr(members, "get", unavailable)

    notifications, error_result = await call_once(
        tool, {"group_id": ALL_STAFF, "member_id": AOIFE}
    )

    assert notifications == [notification]
    assert error_result == f"Error calling tool '{tool}': Service unavailable"
    assert members.ids() == ["user-0001"]


async def test_the_membership_check_asks_graph_for_that_member_only(graph, monkeypatch):
    """Graph evaluates ``$filter`` on a member list only as an advanced query."""
    members = all_staff(graph, ["user-0001"])
    configurations = []
    real_get = members.get

    async def recording_get(request_configuration=None):
        configurations.append(request_configuration)
        return await real_get(request_configuration)

    monkeypatch.setattr(members, "get", recording_get)

    await use_tool(
        "add_group_member", {"group_id": ALL_STAFF, "member_id": "o'brien@contoso.com"}
    )

    [check] = configurations
    assert check.query_parameters.filter == "id eq 'o''brien@contoso.com'"
    assert check.query_parameters.count is True
    assert check.headers.get("ConsistencyLevel") == {"eventual"}


# -- D7: a ValueError on the member tools is sent as its own text --------------------------------


@pytest.mark.parametrize("tool", ["add_group_member", "remove_group_member"])
async def test_a_missing_group_is_reported_with_its_own_text(graph, monkeypatch, tool):
    async def no_group(graph_client, group_id):
        return None

    monkeypatch.setattr(groups, "get_group_by_id", no_group)

    notifications, error_result = await call_once(
        tool, {"group_id": "g-1", "member_id": AOIFE}
    )

    assert notifications == ["Group with ID g-1 not found"]
    assert error_result == f"Error calling tool '{tool}': Group with ID g-1 not found"


# -- D6: the version when the package is not installed -------------------------------------------


def test_the_version_is_0_0_0_when_the_package_is_not_installed(monkeypatch):
    def not_installed(name):
        raise importlib.metadata.PackageNotFoundError(name)

    monkeypatch.setattr(importlib.metadata, "version", not_installed)

    namespace = runpy.run_path(msgraph_mcp_server.__file__)

    assert namespace["__version__"] == "0.0.0"


def test_the_version_is_the_installed_release(monkeypatch):
    monkeypatch.setattr(
        importlib.metadata,
        "version",
        lambda name: {"entraid-mcp-server": "0.3.1"}[name],
    )

    namespace = runpy.run_path(msgraph_mcp_server.__file__)

    assert namespace["__version__"] == "0.3.1"


# -- Row 6 / R2: main sets up logging, then loads the settings file, then serves ------------------


def test_main_sets_up_logging_then_loads_the_settings_file_then_serves(monkeypatch):
    steps = []
    monkeypatch.setattr(
        logging, "basicConfig", lambda **kwargs: steps.append(("logging", kwargs))
    )
    monkeypatch.setattr(
        server, "load_env_files", lambda: steps.append(("settings", {}))
    )
    monkeypatch.setattr(server.mcp, "run", lambda: steps.append(("serve", {})))

    server.main()

    assert steps == [
        (
            "logging",
            {
                "level": logging.INFO,
                "format": "%(asctime)s - %(name)s - %(levelname)s - %(message)s",
            },
        ),
        ("settings", {}),
        ("serve", {}),
    ]


# -- Row 7 / R2: credential settings are read when the credential is first built -----------------

CREDENTIAL_VARIABLES = [
    "TENANT_ID",
    "AZURE_TENANT_ID",
    "CLIENT_ID",
    "AZURE_CLIENT_ID",
    "CLIENT_SECRET",
    "AZURE_CLIENT_SECRET",
    "CERTIFICATE_PATH",
    "AZURE_CLIENT_CERTIFICATE_PATH",
    "CERTIFICATE_PWD",
    "AZURE_CLIENT_CERTIFICATE_PASSWORD",
]


class _Credential:
    """Stands in for an azure-identity credential: records how it was built."""

    def __init__(self, **kwargs):
        self.kwargs = kwargs

    def get_token(self, *scopes, **kwargs):
        raise AssertionError("no sign-in in a unit test")


@pytest.fixture
def credentials(monkeypatch):
    """An environment with no credential settings, and recording credential classes."""
    for name in CREDENTIAL_VARIABLES:
        monkeypatch.delenv(name, raising=False)
    kinds = {}
    for kind in (
        "ClientSecretCredential",
        "CertificateCredential",
        "DefaultAzureCredential",
    ):
        kinds[kind] = type(kind, (_Credential,), {})
        monkeypatch.setattr(graph_auth, kind, kinds[kind])
    return kinds


def test_settings_set_after_construction_reach_the_credential(monkeypatch, credentials):
    manager = graph_auth.GraphAuthManager()
    monkeypatch.setenv("TENANT_ID", "contoso")
    monkeypatch.setenv("CLIENT_ID", "payroll-sync-app")
    monkeypatch.setenv("CLIENT_SECRET", "contoso-secret-0042")

    credential = manager.get_credential()

    assert type(credential).__name__ == "ClientSecretCredential"
    assert credential.kwargs == {
        "tenant_id": "contoso",
        "client_id": "payroll-sync-app",
        "client_secret": "contoso-secret-0042",
    }
    assert manager.get_credential() is credential


@pytest.mark.parametrize(
    "variables, kind, kwargs",
    [
        pytest.param(
            {
                "AZURE_TENANT_ID": "contoso",
                "AZURE_CLIENT_ID": "payroll-sync-app",
                "AZURE_CLIENT_SECRET": "contoso-secret-0042",
                "AZURE_CLIENT_CERTIFICATE_PATH": "/certs/payroll-sync.pem",
            },
            "ClientSecretCredential",
            {
                "tenant_id": "contoso",
                "client_id": "payroll-sync-app",
                "client_secret": "contoso-secret-0042",
            },
            id="a client secret and a certificate",
        ),
        pytest.param(
            {
                "TENANT_ID": "contoso",
                "CLIENT_ID": "payroll-sync-app",
                "CERTIFICATE_PATH": "/certs/payroll-sync.pem",
                "CERTIFICATE_PWD": "pem-pass",
            },
            "CertificateCredential",
            {
                "tenant_id": "contoso",
                "client_id": "payroll-sync-app",
                "certificate_path": "/certs/payroll-sync.pem",
                "password": "pem-pass",
            },
            id="only a certificate",
        ),
        pytest.param(
            {
                "AZURE_TENANT_ID": "contoso",
                "AZURE_CLIENT_ID": "payroll-sync-app",
                "AZURE_CLIENT_CERTIFICATE_PATH": "/certs/payroll-sync.pem",
                "AZURE_CLIENT_CERTIFICATE_PASSWORD": "pem-pass",
            },
            "CertificateCredential",
            {
                "tenant_id": "contoso",
                "client_id": "payroll-sync-app",
                "certificate_path": "/certs/payroll-sync.pem",
                "password": "pem-pass",
            },
            id="only a certificate, Azure variable names",
        ),
        pytest.param(
            {"TENANT_ID": "contoso", "CLIENT_ID": "payroll-sync-app"},
            "DefaultAzureCredential",
            {},
            id="no secret and no certificate",
        ),
    ],
)
def test_sign_in_tries_credentials_in_the_same_order(
    monkeypatch, credentials, variables, kind, kwargs
):
    for name, value in variables.items():
        monkeypatch.setenv(name, value)

    credential = graph_auth.GraphAuthManager().get_credential()

    assert type(credential).__name__ == kind
    assert credential.kwargs == kwargs


def test_constructor_settings_win_over_the_environment(monkeypatch, credentials):
    monkeypatch.setenv("TENANT_ID", "fabrikam")
    monkeypatch.setenv("CLIENT_ID", "fabrikam-app")
    monkeypatch.setenv("CLIENT_SECRET", "fabrikam-secret")

    credential = graph_auth.GraphAuthManager(
        tenant_id="contoso",
        client_id="payroll-sync-app",
        client_secret="contoso-secret-0042",
    ).get_credential()

    assert credential.kwargs == {
        "tenant_id": "contoso",
        "client_id": "payroll-sync-app",
        "client_secret": "contoso-secret-0042",
    }


def test_a_supplied_credential_is_used_as_is(credentials):
    supplied = _Credential()

    assert graph_auth.GraphAuthManager(credential=supplied).get_credential() is supplied


@pytest.mark.parametrize(
    "build, client_class",
    [
        ("get_graph_client", "GraphServiceClient"),
        ("get_beta_graph_client", "GraphBetaServiceClient"),
    ],
)
def test_each_graph_client_is_built_once_with_the_credential(
    monkeypatch, build, client_class
):
    built = []
    monkeypatch.setattr(
        graph_auth,
        client_class,
        lambda **kwargs: built.append(kwargs) or SimpleNamespace(**kwargs),
    )
    supplied = _Credential()
    manager = graph_auth.GraphAuthManager(credential=supplied)

    first = getattr(manager, build)()

    assert getattr(manager, build)() is first
    assert built == [
        {"credentials": supplied, "scopes": ["https://graph.microsoft.com/.default"]}
    ]


@pytest.mark.parametrize(
    "build, client_class, reason",
    [
        (
            "get_graph_client",
            "GraphServiceClient",
            "Failed to create Graph client: bad scope",
        ),
        (
            "get_beta_graph_client",
            "GraphBetaServiceClient",
            "Failed to create Graph beta client: bad scope",
        ),
    ],
)
def test_a_graph_client_that_cannot_be_built_is_an_authentication_error(
    monkeypatch, build, client_class, reason
):
    def refuse(**kwargs):
        raise ValueError("bad scope")

    monkeypatch.setattr(graph_auth, client_class, refuse)
    manager = graph_auth.GraphAuthManager(credential=_Credential())

    with pytest.raises(graph_auth.AuthenticationError, match=f"^{reason}$"):
        getattr(manager, build)()


@pytest.mark.parametrize("build", ["get_graph_client", "get_beta_graph_client"])
def test_an_authentication_error_passes_through_unchanged(monkeypatch, build):
    manager = graph_auth.GraphAuthManager()

    def no_login():
        raise graph_auth.AuthenticationError("no Azure login found")

    monkeypatch.setattr(manager, "get_credential", no_login)

    with pytest.raises(graph_auth.AuthenticationError, match="^no Azure login found$"):
        getattr(manager, build)()


@pytest.mark.parametrize(
    "get, build",
    [("get_client", "get_graph_client"), ("get_beta_client", "get_beta_graph_client")],
)
def test_the_graph_client_wrapper_builds_each_client_once(get, build):
    calls = []
    auth_manager = SimpleNamespace(
        **{build: lambda: calls.append(build) or f"{build} client"}
    )
    client = GraphClient(auth_manager)

    assert getattr(client, get)() == f"{build} client"
    assert getattr(client, get)() == f"{build} client"
    assert calls == [build]


# -- MFA method types (mfa.py lines touched by row 5) ---------------------------------------------

NO_METHODS = {
    "email": False,
    "fido2": False,
    "authenticatorApp": False,
    "password": False,
    "phone": False,
    "softwareOath": False,
    "temporaryAccessPass": False,
    "windowsHelloForBusiness": False,
}


@pytest.mark.parametrize(
    "odata_type, method, status",
    [
        ("#microsoft.graph.emailAuthenticationMethod", "email", "Enabled"),
        ("#microsoft.graph.fido2AuthenticationMethod", "fido2", "Enabled"),
        (
            "#microsoft.graph.microsoftAuthenticatorAuthenticationMethod",
            "authenticatorApp",
            "Enabled",
        ),
        ("#microsoft.graph.passwordAuthenticationMethod", "password", "Disabled"),
        ("#microsoft.graph.phoneAuthenticationMethod", "phone", "Enabled"),
        (
            "#microsoft.graph.softwareOathAuthenticationMethod",
            "softwareOath",
            "Enabled",
        ),
        (
            "#microsoft.graph.temporaryAccessPassAuthenticationMethod",
            "temporaryAccessPass",
            "Enabled",
        ),
        (
            "#microsoft.graph.windowsHelloForBusinessAuthenticationMethod",
            "windowsHelloForBusiness",
            "Enabled",
        ),
    ],
)
async def test_each_authentication_method_is_reported(
    graph, ctx, odata_type, method, status
):
    graph.add_user(AOIFE)
    graph.users.by_user_id(AOIFE).authentication.methods.items = [
        SimpleNamespace(odata_type=odata_type)
    ]

    report = await server.get_user_mfa_status(AOIFE, ctx)

    assert report == {
        "userPrincipalName": AOIFE,
        "mail": AOIFE,
        "companyName": "Contoso",
        "mfaStatus": status,
        "methods": {**NO_METHODS, method: True},
    }


async def test_a_password_after_a_strong_method_keeps_mfa_enabled(graph, ctx):
    graph.add_user(AOIFE)
    graph.users.by_user_id(AOIFE).authentication.methods.items = [
        SimpleNamespace(odata_type="#microsoft.graph.fido2AuthenticationMethod"),
        SimpleNamespace(odata_type="#microsoft.graph.passwordAuthenticationMethod"),
    ]

    report = await server.get_user_mfa_status(AOIFE, ctx)

    assert report["mfaStatus"] == "Enabled"
    assert report["methods"] == {**NO_METHODS, "fido2": True, "password": True}


async def test_no_authentication_methods_answer_reports_no_mfa_data(graph, ctx):
    graph.add_user(AOIFE)
    graph.users.by_user_id(AOIFE).authentication.methods.respond_none = True

    report = await server.get_user_mfa_status(AOIFE, ctx)

    assert report is None
    assert ctx.warnings == [f"No MFA data found for user {AOIFE}"]


# -- Paging: every list tool follows Graph's next links (lines touched by rows 5 and 14) ----------


def device(graph):
    return graph.device_management.managed_devices.by_managed_device_id("device-1")


PAGED_LISTS = [
    (
        "get_all_managed_devices",
        {},
        lambda g: g.device_management.managed_devices,
        ManagedDevice,
    ),
    (
        "get_managed_devices_by_user",
        {"user_id": AOIFE},
        lambda g: g.device_management.managed_devices,
        ManagedDevice,
    ),
    (
        "get_detected_apps_for_device",
        {"device_id": "device-1"},
        lambda g: device(g).detected_apps,
        DetectedApp,
    ),
    (
        "get_device_compliance_policy_states",
        {"device_id": "device-1"},
        lambda g: device(g).device_compliance_policy_states,
        DeviceCompliancePolicyState,
    ),
    (
        "get_device_configuration_states",
        {"device_id": "device-1"},
        lambda g: device(g).device_configuration_states,
        DeviceConfigurationState,
    ),
    (
        "get_device_compliance_policies",
        {},
        lambda g: g.device_management.device_compliance_policies,
        Windows10CompliancePolicy,
    ),
    (
        "get_device_configurations",
        {},
        lambda g: g.device_management.device_configurations,
        Windows10GeneralConfiguration,
    ),
    (
        "get_device_categories",
        {},
        lambda g: g.device_management.device_categories,
        DeviceCategory,
    ),
    ("list_applications", {}, lambda g: g.applications, Application),
    ("list_service_principals", {}, lambda g: g.service_principals, ServicePrincipal),
]


@pytest.mark.parametrize(
    "tool, arguments, collection, model",
    PAGED_LISTS,
    ids=[case[0] for case in PAGED_LISTS],
)
async def test_a_list_tool_collects_every_page(
    graph, ctx, tool, arguments, collection, model
):
    items = collection(graph)
    items.items = [model(id=f"item-{i}") for i in (1, 2, 3)]
    items.page_size = 1

    result = await getattr(server, tool)(ctx=ctx, **arguments)

    assert [entry["id"] for entry in result] == ["item-1", "item-2", "item-3"]
    assert [r.method for r in graph.requests_to(items.path)] == [
        "get",
        "with_url",
        "with_url",
    ]


async def test_privileged_users_collect_every_page_of_role_members(graph, ctx):
    graph.directory_roles.items = [
        DirectoryRole(id="role-1", display_name="Global Administrator"),
        DirectoryRole(id="role-2", display_name="User Administrator"),
    ]
    for role_id in ("role-1", "role-2"):
        members = graph.directory_roles.by_directory_role_id(role_id).members
        members.items = [
            User(id=AOIFE, odata_type="#microsoft.graph.user"),
            ServicePrincipal(
                id="payroll-sync-sp", odata_type="#microsoft.graph.servicePrincipal"
            ),
            User(id=f"user-{role_id}", odata_type="#microsoft.graph.user"),
        ]
        members.page_size = 1

    result = await server.get_privileged_users(ctx)

    assert [(user["id"], sorted(user["roles"])) for user in result] == [
        (AOIFE, ["Global Administrator", "User Administrator"]),
        ("user-role-1", ["Global Administrator"]),
        ("user-role-2", ["User Administrator"]),
    ]


async def test_user_groups_collect_every_page_and_skip_other_objects(graph, ctx):
    memberships = graph.users.by_user_id(AOIFE).transitive_member_of
    memberships.items = [
        Group(id="group-1", odata_type="#microsoft.graph.group"),
        DirectoryRole(id="role-1", odata_type="#microsoft.graph.directoryRole"),
        Group(id="group-2", odata_type="#microsoft.graph.group"),
    ]
    memberships.page_size = 1
    for group_id in ("group-1", "group-2"):
        graph.groups.by_group_id(group_id).value = Group(
            id=group_id, display_name=group_id
        )

    result = await server.get_user_groups(AOIFE, ctx)

    assert [group["id"] for group in result] == ["group-1", "group-2"]
    assert len(graph.requests_to(memberships.path)) == 3


async def test_user_roles_collect_every_page_and_skip_other_objects(graph, ctx):
    memberships = graph.users.by_user_id(AOIFE).member_of
    memberships.items = [
        DirectoryRole(id="role-1", odata_type="#microsoft.graph.directoryRole"),
        Group(id="group-1", odata_type="#microsoft.graph.group"),
        DirectoryRole(id="role-2", odata_type="#microsoft.graph.directoryRole"),
    ]
    memberships.page_size = 1
    for role_id in ("role-1", "role-2"):
        graph.directory_roles.by_directory_role_id(role_id).value = DirectoryRole(
            id=role_id, display_name=role_id
        )

    result = await server.get_user_roles(AOIFE, ctx)

    assert [role["id"] for role in result] == ["role-1", "role-2"]
    assert len(graph.requests_to(memberships.path)) == 3


# -- Create and update pass every supported field to Graph ----------------------------------------

APP_FIELDS = {
    "displayName": "Payroll Sync",
    "signInAudience": "AzureADMyOrg",
    "tags": ["payroll"],
    "identifierUris": ["api://payroll-sync"],
    "web": None,
    "api": None,
    "requiredResourceAccess": [],
}
APP_BODY = {
    "display_name": "Payroll Sync",
    "sign_in_audience": "AzureADMyOrg",
    "tags": ["payroll"],
    "identifier_uris": ["api://payroll-sync"],
    "web": None,
    "api": None,
    "required_resource_access": [],
}
SP_FIELDS = {
    "accountEnabled": False,
    "tags": ["payroll"],
    "appRoleAssignmentRequired": True,
    "displayName": "Payroll Sync",
}
SP_BODY = {
    "account_enabled": False,
    "tags": ["payroll"],
    "app_role_assignment_required": True,
    "display_name": "Payroll Sync",
}


def sent(body, fields):
    return {field: getattr(body, field) for field in fields}


async def test_an_application_is_created_with_every_field(graph, ctx):
    await server.create_application(ctx, APP_FIELDS)

    [post] = graph.requests_to("applications")
    assert sent(post.body, APP_BODY) == APP_BODY


async def test_an_application_is_updated_with_every_field(graph, ctx):
    graph.applications.by_application_id("payroll-sync-object").value = Application(
        id="payroll-sync-object", app_id="payroll-sync-app"
    )

    await server.update_application("payroll-sync-object", ctx, APP_FIELDS)

    [patch] = [
        r
        for r in graph.requests_to("applications/payroll-sync-object")
        if r.method == "patch"
    ]
    assert sent(patch.body, APP_BODY) == APP_BODY


async def test_a_service_principal_is_created_with_every_field(graph, ctx):
    await server.create_service_principal(
        ctx, {"appId": "payroll-sync-app", **SP_FIELDS}
    )

    [post] = graph.requests_to("servicePrincipals")
    assert sent(post.body, {"app_id": "payroll-sync-app", **SP_BODY}) == {
        "app_id": "payroll-sync-app",
        **SP_BODY,
    }


async def test_a_service_principal_is_updated_with_every_field(graph, ctx):
    sp = graph.service_principals.by_service_principal_id("payroll-sync-sp")
    sp.value = ServicePrincipal(id="payroll-sync-sp")

    await server.update_service_principal("payroll-sync-sp", ctx, SP_FIELDS)

    [patch] = [
        r
        for r in graph.requests_to("servicePrincipals/payroll-sync-sp")
        if r.method == "patch"
    ]
    assert sent(patch.body, SP_BODY) == SP_BODY


# -- Conditional access parsing: missing parts and unreadable groups (lines touched by row 14) -----

NO_CONDITIONS = dict.fromkeys(
    [
        "Users_Include",
        "Users_Exclude",
        "Groups_Include",
        "Groups_Include_Names",
        "Groups_Exclude",
        "Groups_Exclude_Names",
        "Roles_Include",
        "Roles_Exclude",
        "Include_Guest_Or_External_Users",
        "Exclude_Guest_Or_External_Users",
        "Apps_Include",
        "Apps_Exclude",
        "User_Actions",
        "Authentication_Context_References",
        "Application_Filter",
        "User_Risk_Levels",
        "Sign_In_Risk_Levels",
        "Service_Principal_Risk_Levels",
        "Insider_Risk_Levels",
        "Client_App_Types",
        "Platforms",
        "Locations",
        "Devices",
        "Client_Applications",
    ],
    "",
)
NO_DEVICE_CONDITIONS = {
    "platforms": None,
    "locations": None,
    "devices": None,
    "client_applications": None,
}


@pytest.mark.parametrize(
    "conditions",
    [
        pytest.param(
            SimpleNamespace(**NO_DEVICE_CONDITIONS), id="no applications and no users"
        ),
        pytest.param(
            SimpleNamespace(
                applications=SimpleNamespace(),
                users=SimpleNamespace(),
                **NO_DEVICE_CONDITIONS,
            ),
            id="empty applications and users",
        ),
    ],
)
async def test_conditions_without_their_parts_parse_as_empty(graph, conditions):
    assert await conditional_access.parse_conditions(graph, conditions) == NO_CONDITIONS


async def test_conditions_that_cannot_be_parsed_leave_one_warning(graph, caplog):
    caplog.set_level(logging.WARNING)

    parsed = await conditional_access.parse_conditions(graph, SimpleNamespace())

    assert parsed == NO_CONDITIONS
    assert [r.getMessage() for r in caplog.records] == [
        "Error parsing conditions: 'types.SimpleNamespace' object has no attribute 'platforms'"
    ]


async def test_group_names_cover_well_known_values_and_unreadable_groups(graph, caplog):
    graph.groups.by_group_id("group-1").value = Group(
        id="group-1", display_name="All Staff"
    )
    caplog.set_level(logging.WARNING)

    names = await conditional_access.get_group_details(
        graph, ["", "All", "GuestsOrExternalUsers", "group-1", "group-404"]
    )

    assert names == {
        "All": "All",
        "GuestsOrExternalUsers": "GuestsOrExternalUsers",
        "group-1": "All Staff (group-1)",
        "group-404": "Unknown Group (group-404)",
    }
    assert [r.getMessage() for r in caplog.records] == [
        (
            "Could not fetch details for group group-404: Request_ResourceNotFound: "
            "Resource 'groups/group-404' does not exist."
        )
    ]
