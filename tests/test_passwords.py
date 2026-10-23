"""Password reset through the `reset_user_password_direct` tool (DRK-2421 §5, DRK-2424 §7)."""

import random
import string

import pytest

from msgraph_mcp_server import server

AOIFE = "aoife@contoso.com"
SUPPLIED = "Sp!ce-Harbor-2026"


@pytest.fixture
def aoife(graph):
    graph.add_user(AOIFE)
    return graph.users.by_user_id(AOIFE)


async def reset(ctx, **kwargs):
    return await server.reset_user_password_direct(AOIFE, ctx, **kwargs)


async def reset_expecting(graph_refuses, ctx, caplog, **kwargs):
    if graph_refuses:
        with pytest.raises(Exception, match="does not comply with the password policy"):
            await reset(ctx, **kwargs)
        # The refusal is logged, so the absence check below reads a real error line.
        assert [
            r
            for r in caplog.records
            if r.levelname == "ERROR" and AOIFE in r.getMessage()
        ]
    else:
        await reset(ctx, **kwargs)


async def test_a_generated_password_does_not_follow_pythons_ordinary_random_generator(
    graph, ctx, aoife
):
    random.seed(42)
    first = await reset(ctx, generate_password=True, password_length=16)
    random.seed(42)
    second = await reset(ctx, generate_password=True, password_length=16)

    assert first["generated_password"] != second["generated_password"]


@pytest.mark.parametrize("length", [8, 16, 256])
async def test_a_generated_password_has_the_asked_length_and_every_character_type(
    graph, ctx, aoife, length
):
    await reset(ctx, generate_password=True, password_length=length)

    [password] = graph.password_changes(AOIFE)
    assert len(password) == length
    assert any(c in string.digits for c in password)
    assert any(c in string.ascii_uppercase for c in password)
    assert any(c in string.ascii_lowercase for c in password)
    assert any(not c.isalnum() for c in password)


async def test_a_generated_password_has_12_characters_when_no_length_is_asked(
    graph, ctx, aoife
):
    await reset(ctx, generate_password=True)

    [password] = graph.password_changes(AOIFE)
    assert len(password) == 12


@pytest.mark.parametrize("length", [3, 7, 257])
async def test_a_generated_length_outside_8_to_256_is_refused(
    graph, ctx, aoife, length
):
    with pytest.raises(ValueError) as refusal:
        await reset(ctx, generate_password=True, password_length=length)

    # The spec names no exact wording, only that the refusal names both of these.
    assert "password length" in str(refusal.value).lower()
    assert "8 to 256" in str(refusal.value)
    assert graph.password_changes(AOIFE) == []


async def test_the_generated_password_comes_back_to_the_admin(graph, ctx, aoife):
    result = await reset(ctx, generate_password=True, password_length=16)

    [sent] = graph.password_changes(AOIFE)
    assert len(sent) == 16
    assert result["generated_password"] == sent


@pytest.mark.parametrize(
    "graph_refuses", [False, True], ids=["accepts the change", "refuses the change"]
)
async def test_no_generated_password_reaches_the_server_log(
    graph, ctx, aoife, caplog, graph_refuses
):
    caplog.set_level("DEBUG")
    aoife.refuse_patch = graph_refuses

    await reset_expecting(
        graph_refuses, ctx, caplog, generate_password=True, password_length=16
    )

    [password] = graph.password_changes(AOIFE)
    assert password not in caplog.text
    assert not [e for e in ctx.errors if password in e]


@pytest.mark.parametrize(
    "graph_refuses", [False, True], ids=["accepts the change", "refuses the change"]
)
async def test_no_supplied_password_reaches_the_server_log(
    graph, ctx, aoife, caplog, graph_refuses
):
    caplog.set_level("DEBUG")
    aoife.refuse_patch = graph_refuses

    await reset_expecting(graph_refuses, ctx, caplog, password=SUPPLIED)

    assert graph.password_changes(AOIFE) == [SUPPLIED]
    assert SUPPLIED not in caplog.text
    assert not [e for e in ctx.errors if SUPPLIED in e]


async def test_a_supplied_password_goes_to_entra_id_unchanged(graph, ctx, aoife):
    # "Ab1!xy" has 6 characters: the 8 to 256 range does not apply to a supplied password.
    await reset(ctx, password="Ab1!xy")

    assert graph.password_changes(AOIFE) == ["Ab1!xy"]


async def test_a_generated_password_wins_over_a_supplied_one(graph, ctx, aoife):
    result = await reset(
        ctx, password="Ab1!xy", generate_password=True, password_length=16
    )

    [sent] = graph.password_changes(AOIFE)
    assert sent != "Ab1!xy"
    assert len(sent) == 16
    assert result["generated_password"] == sent
