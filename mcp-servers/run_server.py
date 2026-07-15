#!/usr/bin/env python3
"""Bootstrap a China finance MCP server with its Python dependencies.

Claude plugins do not install Python requirements during plugin installation.
This launcher creates a per-server virtualenv in the user's cache directory,
installs the server's requirements.txt, then execs the real server.py so stdio
belongs to the MCP server process.
"""

from __future__ import annotations

import hashlib
import os
import subprocess
import sys
import venv
from pathlib import Path


PLUGIN_DIR = Path(__file__).resolve().parent
DEFAULT_CACHE_DIR = Path.home() / ".cache" / "china-finance-mcp"


def _die(message: str, code: int = 1) -> None:
    print(f"china-finance-mcp bootstrap: {message}", file=sys.stderr)
    raise SystemExit(code)


def _requirements_hash(requirements: Path) -> str:
    return hashlib.sha256(requirements.read_bytes()).hexdigest()


def _venv_python(venv_dir: Path) -> Path:
    if os.name == "nt":
        return venv_dir / "Scripts" / "python.exe"
    return venv_dir / "bin" / "python"


def _ensure_venv(server_name: str, requirements: Path) -> Path:
    cache_root = Path(os.environ.get("CHINA_FINANCE_MCP_VENV_DIR", DEFAULT_CACHE_DIR))
    venv_dir = cache_root / server_name
    marker = venv_dir / ".requirements.sha256"
    current_hash = _requirements_hash(requirements)
    python = _venv_python(venv_dir)

    if python.exists() and marker.exists() and marker.read_text(encoding="utf-8").strip() == current_hash:
        return python

    print(f"china-finance-mcp bootstrap: preparing dependencies for {server_name}", file=sys.stderr)
    venv_dir.parent.mkdir(parents=True, exist_ok=True)
    venv.EnvBuilder(with_pip=True, clear=venv_dir.exists()).create(venv_dir)

    env = os.environ.copy()
    env.setdefault("PIP_DISABLE_PIP_VERSION_CHECK", "1")
    subprocess.run(
        [str(python), "-m", "pip", "install", "-r", str(requirements)],
        check=True,
        env=env,
        stdout=sys.stderr,
        stderr=sys.stderr,
    )
    marker.write_text(current_hash, encoding="utf-8")
    return python


def main() -> None:
    if len(sys.argv) < 2:
        _die("usage: run_server.py <server-dir> [server-args...]")

    server_name = sys.argv[1]
    if "/" in server_name or "\\" in server_name or server_name.startswith("."):
        _die(f"invalid server directory: {server_name}")

    server_dir = PLUGIN_DIR / server_name
    server_py = server_dir / "server.py"
    requirements = server_dir / "requirements.txt"

    if not server_py.exists():
        _die(f"missing server script: {server_py}")
    if not requirements.exists():
        _die(f"missing requirements file: {requirements}")

    python = _ensure_venv(server_name, requirements)
    os.execv(str(python), [str(python), str(server_py), *sys.argv[2:]])


if __name__ == "__main__":
    main()
