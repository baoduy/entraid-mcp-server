"""Import and start-up (DRK-2429 §5; brief §3 rows 6–7, §6 R2–R4, §6a D1–D3).

Each test runs the server in a fresh interpreter on its own throwaway machine
(``server_process.Machine``): ``conftest`` has already imported the server in the test process.
"""

import json
import logging
import re

import pytest
from server_process import START_PATHS, Machine

WORKING_CONFIG = "the working folder's config folder"
CONTOSO_SECRET = {
    "TENANT_ID": "contoso",
    "CLIENT_ID": "payroll-sync-app",
    "CLIENT_SECRET": "contoso-secret-0042",
}
FABRIKAM_SECRET = {
    "TENANT_ID": "fabrikam",
    "CLIENT_ID": "payroll-sync-app",
    "CLIENT_SECRET": "fabrikam-secret-0042",
}


@pytest.fixture
def machine(tmp_path):
    return Machine(tmp_path)


async def first_graph_call(machine, start_path="Python's module runner"):
    """Starts the server, has an AI client call one tool, and returns the server's standard error."""
    running = await machine.start(start_path)
    try:
        await running.initialize()
        await running.call_tool("search_users", {"query": "Aoife"})
    finally:
        stderr = await running.stop()
    machine.stdout_lines = running.stdout_lines
    return stderr


def tenants_signed_in_to(machine):
    return [sign_in["tenant_id"] for sign_in in machine.sign_ins()]


# -- Importing the server -------------------------------------------------------------------


def test_importing_the_server_leaves_the_process_environment_alone(machine):
    machine.settings_file(WORKING_CONFIG, TENANT_ID="contoso")

    changed = machine.run_host(
        "import json, os\n"
        "before = dict(os.environ)\n"
        "import msgraph_mcp_server.server\n"
        "print(json.dumps({k: v for k, v in os.environ.items() if before.get(k) != v}))\n"
    )

    assert changed == {}


def test_importing_the_server_leaves_the_logging_setup_alone(machine):
    setup = machine.run_host(
        "import json, logging\n"
        "root = logging.getLogger()\n"
        "root.setLevel(logging.WARNING)  # the host's own setup: warnings and up, handlers added later\n"
        "def setup():\n"
        "    return {'level': logging.getLevelName(root.level), 'handlers': [repr(h) for h in root.handlers]}\n"
        "before = setup()\n"
        "import msgraph_mcp_server.server\n"
        "print(json.dumps([before, setup()]))\n"
    )

    before, after = setup
    assert before == {"level": "WARNING", "handlers": []}
    assert after == before


def test_importing_the_server_writes_no_log_line(machine):
    """A record at INFO or above from any logger fails, as does any record from the server's own
    loggers (``msgraph_mcp_server.*``) or the root logger.

    Only ``mcp.*`` and ``fastmcp.*`` records below INFO are tolerated: building the ``FastMCP``
    object writes them, for example mcp's DEBUG ``Initializing server 'EntraID MCP Server'`` and
    ``Registering handler for ListToolsRequest``.
    """
    machine.settings_file(WORKING_CONFIG, TENANT_ID="contoso")

    records = machine.run_host(
        "import json, logging\n"
        "class Recorder(logging.Handler):\n"
        "    records = []\n"
        "    def emit(self, record):\n"
        "        self.records.append([record.name, record.levelno, record.getMessage()])\n"
        "root = logging.getLogger()\n"
        "root.setLevel(logging.DEBUG)\n"
        "root.addHandler(Recorder())\n"
        "import msgraph_mcp_server.server\n"
        "print(json.dumps(Recorder.records))\n"
    )

    def tolerated(name, level):
        library = name.split(".")[0] in ("mcp", "fastmcp")
        return library and level < logging.INFO

    assert [r for r in records if not tolerated(r[0], r[1])] == []


# -- Starting the server ------------------------------------------------------------------------


@pytest.mark.parametrize("start_path", START_PATHS)
async def test_every_start_path_loads_the_settings_file(machine, start_path):
    machine.settings_file(WORKING_CONFIG, **CONTOSO_SECRET)

    await first_graph_call(machine, start_path)

    assert tenants_signed_in_to(machine) == ["contoso"]


LOG_LINE = re.compile(r"^\d{4}-\d\d-\d\d \d\d:\d\d:\d\d,\d{3} - [\w.]+ - (\w+) - (.*)$")


@pytest.mark.parametrize("start_path", START_PATHS)
async def test_every_start_path_writes_the_server_log_to_standard_error(
    machine, start_path
):
    machine.settings_file(WORKING_CONFIG, **CONTOSO_SECRET)

    stderr = await first_graph_call(machine, start_path)

    entries = [
        match.groups() for match in map(LOG_LINE.match, stderr.splitlines()) if match
    ]
    assert ("INFO", "Successfully created Microsoft Graph client") in entries
    assert machine.stdout_lines
    assert all(json.loads(line)["jsonrpc"] == "2.0" for line in machine.stdout_lines)


@pytest.mark.parametrize(
    "first_place, second_place",
    [
        ("the repo's config folder", "the working folder's config folder"),
        ("the working folder's config folder", "the home folder's entraid folder"),
        ("the home folder's entraid folder", "the system entraid folder"),
    ],
)
async def test_the_first_settings_file_found_wins(machine, first_place, second_place):
    machine.settings_file(first_place, **CONTOSO_SECRET)
    machine.settings_file(second_place, **FABRIKAM_SECRET)

    await first_graph_call(machine)

    assert tenants_signed_in_to(machine) == ["contoso"]


async def test_the_second_place_alone_is_read(machine):
    """Presence sibling of the rows above: the system folder stand-in is a place the server reads."""
    machine.settings_file("the system entraid folder", **FABRIKAM_SECRET)

    await first_graph_call(machine)

    assert tenants_signed_in_to(machine) == ["fabrikam"]


async def test_a_variable_already_set_wins_over_the_settings_file(machine):
    machine.env.update(FABRIKAM_SECRET)
    machine.settings_file(WORKING_CONFIG, **CONTOSO_SECRET)

    await first_graph_call(machine)

    assert tenants_signed_in_to(machine) == ["fabrikam"]


async def test_the_server_starts_without_a_settings_file(machine):
    machine.env.update(FABRIKAM_SECRET)

    await first_graph_call(machine)

    assert tenants_signed_in_to(machine) == ["fabrikam"]


async def test_credentials_from_the_settings_file_reach_sign_in(machine):
    machine.settings_file(WORKING_CONFIG, **CONTOSO_SECRET)

    await first_graph_call(machine)

    assert machine.sign_ins() == [
        {
            "kind": "client secret",
            "tenant_id": "contoso",
            "client_id": "payroll-sync-app",
            "client_secret": "contoso-secret-0042",
            "certificate_path": None,
        }
    ]


async def test_sign_in_waits_for_the_first_microsoft_graph_call(machine):
    machine.settings_file(WORKING_CONFIG, **CONTOSO_SECRET)

    running = await machine.start()
    try:
        await running.initialize()
        await running.request("tools/list")
    finally:
        await running.stop()

    assert machine.sign_ins() == []


@pytest.mark.parametrize(
    "credentials, kind",
    [
        pytest.param(
            {
                "CLIENT_SECRET": "contoso-secret-0042",
                "CERTIFICATE_PATH": "/certs/payroll-sync.pem",
            },
            "client secret",
            id="a client secret and a certificate",
        ),
        pytest.param(
            {"CERTIFICATE_PATH": "/certs/payroll-sync.pem"},
            "certificate",
            id="only a certificate",
        ),
        pytest.param({}, "current Azure login", id="no secret and no certificate"),
    ],
)
async def test_sign_in_tries_credentials_in_the_same_order(machine, credentials, kind):
    machine.settings_file(
        WORKING_CONFIG, TENANT_ID="contoso", CLIENT_ID="payroll-sync-app", **credentials
    )

    await first_graph_call(machine)

    assert [sign_in["kind"] for sign_in in machine.sign_ins()] == [kind]
