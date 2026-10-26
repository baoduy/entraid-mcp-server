"""The installed package's requirements (DRK-2429 §5, §3b)."""

from importlib.metadata import metadata, requires

import pytest
from packaging.requirements import Requirement
from packaging.specifiers import SpecifierSet


def installed_requirements():
    return {
        Requirement(line).name: Requirement(line).specifier
        for line in requires("entraid-mcp-server")
    }


@pytest.mark.parametrize(
    "package, minimum, next_major",
    [
        ("fastmcp", "4.1.3", "5"),
        ("azure-core", "1.41.0", "2"),
        ("azure-identity", "1.26.0", "2"),
        ("msgraph-core", "1.5.2", "2"),
        ("msgraph-sdk", "1.65.0", "2"),
        ("msgraph-beta-sdk", "1.66.0", "2"),
        ("python-dotenv", "1.2.4", "2"),
    ],
)
def test_the_package_requires_the_latest_release_of_each_runtime_package(
    package, minimum, next_major
):
    assert installed_requirements()[package] == SpecifierSet(
        f">={minimum},<{next_major}"
    )


def test_the_minimum_python_version_stays_3_10():
    assert metadata("entraid-mcp-server")["Requires-Python"] == ">=3.10"
