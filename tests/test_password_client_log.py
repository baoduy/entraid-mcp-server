"""A refused password reset called through an MCP client keeps the password out of every log and the client error (DRK-2424 §3 row 6)."""

import pytest
from fastmcp import Client
from fastmcp.exceptions import ToolError

from msgraph_mcp_server import server

AOIFE = "aoife@contoso.com"
SUPPLIED = "Sp!ce-Harbor-2026"


async def test_a_refused_reset_through_the_mcp_client_keeps_the_password_out_of_the_log_and_the_error(
    graph, caplog
):
    caplog.set_level("DEBUG")
    graph.add_user(AOIFE)
    graph.users.by_user_id(AOIFE).refuse_patch = True

    async with Client(server.mcp) as client:
        with pytest.raises(ToolError) as refusal:
            await client.call_tool(
                "reset_user_password_direct", {"user_id": AOIFE, "password": SUPPLIED}
            )

    assert graph.password_changes(AOIFE) == [SUPPLIED]
    # The refusal reaches the client and fastmcp's own error log, masked.
    assert str(refusal.value) == (
        "Error calling tool 'reset_user_password_direct': "
        "Request_BadRequest: The password '***' does not comply with the password policy."
    )
    assert "Error calling tool 'reset_user_password_direct'" in [
        r.getMessage() for r in caplog.records
    ]
    assert SUPPLIED not in caplog.text
    assert SUPPLIED not in str(refusal.value)
