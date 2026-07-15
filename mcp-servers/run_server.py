#!/usr/bin/env python3
"""Bootstrap a China finance MCP server with its Python dependencies.

Claude plugins do not install Python requirements during plugin installation.
This launcher creates a shared virtualenv in the user's cache directory,
installs all MCP server requirements once, then execs the real server.py so
stdio belongs to the MCP server process.
"""

from __future__ import annotations

import hashlib
import os
import shutil
import subprocess
import sys
import time
from pathlib import Path
from typing import Optional


PLUGIN_DIR = Path(__file__).resolve().parent
DEFAULT_CACHE_DIR = Path.home() / ".cache" / "china-finance-mcp"
MIN_PYTHON = (3, 10)
PYTHON_CANDIDATES = (
    "python3.13",
    "python3.12",
    "python3.11",
    "python3.10",
    "python3",
    "python",
)


def _die(message: str, code: int = 1) -> None:
    print(f"china-finance-mcp bootstrap: {message}", file=sys.stderr)
    raise SystemExit(code)


def _requirements_hash(requirements: Path) -> str:
    hasher = hashlib.sha256()
    for req_file in sorted((PLUGIN_DIR).glob("*-mcp/requirements.txt")):
        hasher.update(req_file.relative_to(PLUGIN_DIR).as_posix().encode("utf-8"))
        hasher.update(b"\0")
        hasher.update(req_file.read_bytes())
        hasher.update(b"\0")
    hasher.update(requirements.read_bytes())
    return hasher.hexdigest()


def _venv_python(venv_dir: Path) -> Path:
    if os.name == "nt":
        return venv_dir / "Scripts" / "python.exe"
    return venv_dir / "bin" / "python"


def _python_version(executable: str) -> Optional[tuple[int, int, int]]:
    try:
        result = subprocess.run(
            [
                executable,
                "-c",
                "import sys; print(f'{sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}')",
            ],
            check=True,
            capture_output=True,
            text=True,
        )
    except (OSError, subprocess.CalledProcessError):
        return None

    try:
        major, minor, micro = result.stdout.strip().split(".", 2)
        return int(major), int(minor), int(micro)
    except ValueError:
        return None


def _find_python() -> str:
    requested = os.environ.get("CHINA_FINANCE_MCP_PYTHON", "").strip()
    candidates = (requested, *PYTHON_CANDIDATES) if requested else PYTHON_CANDIDATES
    seen: set[str] = set()

    for candidate in candidates:
        if not candidate or candidate in seen:
            continue
        seen.add(candidate)
        executable = shutil.which(candidate) if not os.path.isabs(candidate) else candidate
        if not executable:
            continue
        version = _python_version(executable)
        if version and version >= (*MIN_PYTHON, 0):
            return executable

    min_version = ".".join(str(part) for part in MIN_PYTHON)
    _die(
        f"Python {min_version}+ is required. Install python3.12, or set "
        "CHINA_FINANCE_MCP_PYTHON to a compatible Python executable."
    )


class _InstallLock:
    def __init__(self, lock_dir: Path, timeout: int = 600) -> None:
        self.lock_dir = lock_dir
        self.timeout = timeout

    def __enter__(self) -> None:
        start = time.monotonic()
        while True:
            try:
                self.lock_dir.mkdir(parents=True)
                return None
            except FileExistsError:
                if time.monotonic() - start > self.timeout:
                    _die(f"timed out waiting for dependency install lock: {self.lock_dir}")
                time.sleep(0.5)

    def __exit__(self, exc_type, exc, tb) -> None:
        try:
            self.lock_dir.rmdir()
        except OSError:
            pass


def _ensure_venv(server_name: str, requirements: Path) -> Path:
    cache_root = Path(os.environ.get("CHINA_FINANCE_MCP_VENV_DIR", DEFAULT_CACHE_DIR))
    venv_dir = cache_root / "venv"
    marker = venv_dir / ".requirements.sha256"
    current_hash = _requirements_hash(requirements)
    python = _venv_python(venv_dir)

    if python.exists() and marker.exists() and marker.read_text(encoding="utf-8").strip() == current_hash:
        return python

    venv_dir.parent.mkdir(parents=True, exist_ok=True)
    lock_dir = venv_dir.parent / ".install.lock"

    with _InstallLock(lock_dir):
        if python.exists() and marker.exists() and marker.read_text(encoding="utf-8").strip() == current_hash:
            return python

        base_python = _find_python()
        print(
            f"china-finance-mcp bootstrap: preparing shared dependencies for {server_name} with {base_python}",
            file=sys.stderr,
        )
        try:
            subprocess.run([base_python, "-m", "venv", "--clear", str(venv_dir)], check=True, stderr=sys.stderr)
        except subprocess.CalledProcessError as exc:
            _die(f"failed to create virtualenv with {base_python}: exit code {exc.returncode}")

        env = os.environ.copy()
        env.setdefault("PIP_DISABLE_PIP_VERSION_CHECK", "1")
        install_cmd = [str(python), "-m", "pip", "install"]
        for req_file in sorted((PLUGIN_DIR).glob("*-mcp/requirements.txt")):
            install_cmd.extend(["-r", str(req_file)])
        try:
            subprocess.run(
                install_cmd,
                check=True,
                env=env,
                stdout=sys.stderr,
                stderr=sys.stderr,
            )
        except subprocess.CalledProcessError as exc:
            _die(
                "failed to install Python dependencies. Check network/PyPI access, "
                "or set PIP_INDEX_URL to an available package index. "
                f"Exit code: {exc.returncode}"
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
