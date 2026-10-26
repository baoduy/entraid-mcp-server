"""Loaded first by every server the start-up tests launch (``sitecustomize`` on ``PYTHONPATH``).

It replaces the two outbound ports a start-up test cannot reach:

- Azure sign-in. Each azure-identity credential class records how it was built, one JSON line in
  ``$ENTRAID_TEST_SIGN_INS``, and refuses every token, so no request leaves the machine.
- The system settings folder ``/etc/entraid``. Reads of that folder go to
  ``$ENTRAID_TEST_SYSTEM_DIR`` instead, so a test can place a settings file there.
"""

import builtins
import io
import json
import os
import sys

_SYSTEM_DIR = "/etc/entraid"
_system_stand_in = os.environ.get("ENTRAID_TEST_SYSTEM_DIR")


def _redirect(path):
    if not isinstance(path, (str, os.PathLike)):
        return path
    text = os.fspath(path)
    if isinstance(text, str) and (
        text == _SYSTEM_DIR or text.startswith(_SYSTEM_DIR + "/")
    ):
        return _system_stand_in + text[len(_SYSTEM_DIR) :]
    return path


class _Redirecting:
    """A callable object, not a function, so a class that stores it (Python 3.10 ``pathlib``) does not bind it."""

    def __init__(self, function):
        self._function = function

    def __call__(self, path, *args, **kwargs):
        return self._function(_redirect(path), *args, **kwargs)


if _system_stand_in:
    os.stat = _Redirecting(os.stat)
    builtins.open = io.open = _Redirecting(io.open)
    # Python 3.10 pathlib, when imported before this file, holds its own os.stat and io.open.
    accessor = getattr(sys.modules.get("pathlib"), "_NormalAccessor", None)
    if accessor is not None:
        accessor.stat, accessor.open = os.stat, io.open

# Imported only now, so the redirect above is in place first.
import azure.identity
from azure.core.exceptions import ClientAuthenticationError


class _OfflineCredential:
    kind = ""

    def __init__(self, *args, **kwargs):
        sign_in = {
            "kind": self.kind,
            "tenant_id": kwargs.get("tenant_id"),
            "client_id": kwargs.get("client_id"),
            "client_secret": kwargs.get("client_secret"),
            "certificate_path": kwargs.get("certificate_path"),
        }
        with open(os.environ["ENTRAID_TEST_SIGN_INS"], "a", encoding="utf-8") as record:
            record.write(json.dumps(sign_in) + "\n")

    def get_token(self, *scopes, **kwargs):
        raise ClientAuthenticationError("offline test credential: no sign-in")

    get_token_info = get_token

    def close(self):
        pass


class ClientSecretCredential(_OfflineCredential):
    kind = "client secret"


class CertificateCredential(_OfflineCredential):
    kind = "certificate"


class DefaultAzureCredential(_OfflineCredential):
    kind = "current Azure login"


azure.identity.ClientSecretCredential = ClientSecretCredential
azure.identity.CertificateCredential = CertificateCredential
azure.identity.DefaultAzureCredential = DefaultAzureCredential
