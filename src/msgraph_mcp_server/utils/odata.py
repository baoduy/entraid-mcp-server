"""Escaping of user text placed in Microsoft Graph query parameters."""


def odata_string_literal(value: str) -> str:
    """Return ``value`` as one quoted OData string literal for a ``$filter``.

    Each single quote is doubled, so text such as ``Finance' or 1 eq 1`` stays
    part of the value instead of closing the literal and adding a condition.
    """
    return "'" + value.replace("'", "''") + "'"


def graph_search_escape(value: str) -> str:
    """Escape ``value`` for use inside a double-quoted ``$search`` clause.

    Each backslash, then each double quote, is prefixed with a backslash.
    """
    return value.replace("\\", "\\\\").replace('"', '\\"')
