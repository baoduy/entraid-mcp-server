"""What an AI client sees when it lists the server (DRK-2429 §5, §3a)."""

import json
from pathlib import Path

import pytest
from fastmcp import Client
from server_process import MODERN_ENVELOPE, Machine

from msgraph_mcp_server import server

TOOLS_0_3_0 = json.loads(
    (Path(__file__).parent / "fixtures" / "tools_v0_3_0.json").read_text()
)


async def list_tools(protocol, machine):
    """Lists the tools over standard input and output the way a client on ``protocol`` does."""
    running = await machine.start()
    try:
        if protocol == "2026-07-28":
            # 2026-07-28 has no handshake: every request carries its own protocol envelope.
            return await running.request("tools/list", {"_meta": MODERN_ENVELOPE})
        initialized = await running.initialize(protocol)
        assert initialized["protocolVersion"] == protocol
        return await running.request("tools/list")
    finally:
        await running.stop()


@pytest.mark.parametrize("protocol", ["2025-11-25", "2026-07-28"])
async def test_ai_clients_on_both_mcp_protocol_versions_get_the_same_tools(
    tmp_path, protocol
):
    response = await list_tools(protocol, Machine(tmp_path))

    assert "error" not in response
    assert sorted(tool["name"] for tool in response["result"]["tools"]) == sorted(
        TOOLS_0_3_0
    )


def may_be_left_empty(schema):
    return {"type": "null"} in schema.get("anyOf", []) or "null" in schema.get(
        "type", []
    )


@pytest.mark.parametrize(
    "tool, input_name",
    [
        pytest.param(
            "get_all_managed_devices",
            "filter_os",
            id="managed device list · operating system filter",
        ),
        pytest.param(
            "reset_user_password_direct", "password", id="password reset · password"
        ),
        pytest.param(
            "search_permissions",
            "permission_type",
            id="permission search · permission type",
        ),
    ],
)
async def test_optional_inputs_can_still_be_left_empty(tool, input_name):
    async with Client(server.mcp) as client:
        listed = {t.name: t for t in await client.list_tools()}

    schema = getattr(listed[tool], "input_schema", None) or listed[tool].inputSchema
    assert input_name not in schema.get("required", [])
    assert may_be_left_empty(schema["properties"][input_name])


async def test_the_demo_greeting_is_gone():
    async with Client(server.mcp) as client:
        resources = await client.list_resources()
        templates = await client.list_resource_templates()

    assert resources == []
    assert templates == []


@pytest.mark.parametrize("reader", ["an AI client that connects", "a host program"])
async def test_the_release_number_is_reported_as_the_version(tmp_path, reader):
    machine = Machine(tmp_path)
    machine.install_release("0.3.1")

    if reader == "a host program":
        version = machine.run_host(
            "import json, msgraph_mcp_server; print(json.dumps(msgraph_mcp_server.__version__))"
        )
    else:
        running = await machine.start()
        try:
            version = (await running.initialize())["serverInfo"]["version"]
        finally:
            await running.stop()

    assert version == "0.3.1"
