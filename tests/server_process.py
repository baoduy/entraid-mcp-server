"""Runs the server in a fresh interpreter, the way a host program or an AI client starts it.

Each ``Machine`` is a throwaway computer for one test: its own copy of the package (so "the
repo's config folder" is a folder the test owns), working folder, home folder and system
settings folder, and a process environment that holds only ``PATH`` plus what the test sets.
``start_shim/sitecustomize.py`` stands in for Azure sign-in and for ``/etc/entraid``.
"""

import asyncio
import json
import os
import shutil
import subprocess
import sys
from pathlib import Path
from typing import Any

PACKAGE = Path(__file__).resolve().parents[1] / "src" / "msgraph_mcp_server"
SHIM = Path(__file__).resolve().parent / "start_shim"
TIMEOUT = 60

START_PATHS = {
    "the entraid-mcp-server command": [
        str(Path(sys.executable).parent / "entraid-mcp-server")
    ],
    "Python's module runner": [sys.executable, "-m", "msgraph_mcp_server"],
}

CLIENT_INFO = {"name": "ai-client", "version": "1.0.0"}
MODERN_ENVELOPE = {
    "io.modelcontextprotocol/protocolVersion": "2026-07-28",
    "io.modelcontextprotocol/clientCapabilities": {},
    "io.modelcontextprotocol/clientInfo": CLIENT_INFO,
}


class Machine:
    def __init__(self, root: Path):
        self.root = root
        self.repo = root / "repo"
        shutil.copytree(
            PACKAGE,
            self.repo / "src" / "msgraph_mcp_server",
            ignore=shutil.ignore_patterns("__pycache__"),
        )
        self.work = root / "work"
        self.home = root / "home"
        self.system = root / "etc-entraid"
        for folder in (self.work, self.home, self.system):
            folder.mkdir()
        self.sign_in_record = root / "sign-ins.jsonl"
        self.python_path = [str(SHIM), str(self.repo / "src")]
        self.env = {
            "PATH": os.environ["PATH"],
            "HOME": str(self.home),
            "ENTRAID_TEST_SIGN_INS": str(self.sign_in_record),
            "ENTRAID_TEST_SYSTEM_DIR": str(self.system),
        }
        self.places = {
            "the repo's config folder": self.repo / "config",
            "the working folder's config folder": self.work / "config",
            "the home folder's entraid folder": self.home / ".entraid",
            "the system entraid folder": self.system,
        }

    def settings_file(self, place: str, **variables: str) -> None:
        folder = self.places[place]
        folder.mkdir(exist_ok=True)
        (folder / ".env").write_text(
            "".join(f"{name}={value}\n" for name, value in variables.items())
        )

    def install_release(self, version: str) -> None:
        """Makes ``entraid-mcp-server <version>`` the installed release this machine's Python finds."""
        dist_info = self.root / "site" / f"entraid_mcp_server-{version}.dist-info"
        dist_info.mkdir(parents=True)
        (dist_info / "METADATA").write_text(
            f"Metadata-Version: 2.1\nName: entraid-mcp-server\nVersion: {version}\n"
        )
        self.python_path.insert(1, str(dist_info.parent))

    def process_env(self) -> dict[str, str]:
        return {**self.env, "PYTHONPATH": os.pathsep.join(self.python_path)}

    def sign_ins(self) -> list[dict[str, Any]]:
        if not self.sign_in_record.exists():
            return []
        return [
            json.loads(line) for line in self.sign_in_record.read_text().splitlines()
        ]

    def run_host(self, program: str) -> Any:
        """Runs a host program in a fresh interpreter; it prints one JSON value, which is returned."""
        done = subprocess.run(
            [sys.executable, "-c", program],
            cwd=self.work,
            env=self.process_env(),
            capture_output=True,
            text=True,
            timeout=TIMEOUT,
            check=False,
        )
        assert done.returncode == 0, done.stderr
        return json.loads(done.stdout.splitlines()[-1])

    async def start(
        self, start_path: str = "Python's module runner"
    ) -> "RunningServer":
        stderr = (self.root / "stderr.log").open("w")
        process = await asyncio.create_subprocess_exec(
            *START_PATHS[start_path],
            cwd=self.work,
            env=self.process_env(),
            stdin=asyncio.subprocess.PIPE,
            stdout=asyncio.subprocess.PIPE,
            stderr=stderr,
        )
        return RunningServer(process, stderr, self.root / "stderr.log")


class RunningServer:
    """An AI client's raw JSON-RPC line protocol over the server's standard input and output."""

    def __init__(self, process, stderr, stderr_path: Path):
        self._process = process
        self._stderr = stderr
        self._stderr_path = stderr_path
        self._next_id = 0
        self.stdout_lines: list[str] = []

    async def send(self, message: dict[str, Any]) -> None:
        self._process.stdin.write((json.dumps(message) + "\n").encode())
        await self._process.stdin.drain()

    async def request(
        self, method: str, params: dict[str, Any] | None = None
    ) -> dict[str, Any]:
        """Sends a request and returns the whole response message (``result`` or ``error``)."""
        self._next_id += 1
        await self.send(
            {
                "jsonrpc": "2.0",
                "id": self._next_id,
                "method": method,
                "params": params or {},
            }
        )
        while True:
            line = await asyncio.wait_for(self._process.stdout.readline(), TIMEOUT)
            assert line, (
                f"the server closed its output; standard error:\n{self._stderr_path.read_text()}"
            )
            self.stdout_lines.append(line.decode())
            message = json.loads(line)
            if message.get("id") == self._next_id and (
                "result" in message or "error" in message
            ):
                return message

    async def initialize(self, protocol: str = "2025-11-25") -> dict[str, Any]:
        """The handshake of the protocol versions before 2026-07-28; returns the initialize result."""
        response = await self.request(
            "initialize",
            {
                "protocolVersion": protocol,
                "capabilities": {},
                "clientInfo": CLIENT_INFO,
            },
        )
        await self.send({"jsonrpc": "2.0", "method": "notifications/initialized"})
        return response["result"]

    async def call_tool(self, name: str, arguments: dict[str, Any]) -> dict[str, Any]:
        return await self.request("tools/call", {"name": name, "arguments": arguments})

    async def stop(self) -> str:
        """Closes the client side, waits for the server to end and returns its standard error."""
        self._process.stdin.close()
        try:
            rest = await asyncio.wait_for(self._process.stdout.read(), TIMEOUT)
            await asyncio.wait_for(self._process.wait(), TIMEOUT)
        finally:
            if self._process.returncode is None:
                self._process.kill()
            self._stderr.close()
        self.stdout_lines.extend(rest.decode().splitlines(keepends=True))
        return self._stderr_path.read_text()
