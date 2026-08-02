"""Reliable Windows launcher for the bundled Streamlit application."""

from __future__ import annotations

import os
from pathlib import Path
import socket
import sys

# Explicit imports help PyInstaller discover runtime dependencies.
import openpyxl  # noqa: F401
import pandas  # noqa: F401
import pyarrow  # noqa: F401
import streamlit  # noqa: F401

import data_utils  # noqa: F401
from streamlit.web import bootstrap


def bundled_path(filename: str) -> Path:
    """Return a file path in source mode or inside a PyInstaller bundle."""
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS) / filename
    return Path(__file__).resolve().parent / filename


def find_free_port() -> int:
    """Ask Windows for a currently unused local TCP port."""
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.bind(("127.0.0.1", 0))
        return int(sock.getsockname()[1])


def main() -> None:
    """Configure and launch Streamlit on one matching local port."""
    app_path = bundled_path("app.py")
    if not app_path.is_file():
        raise FileNotFoundError(f"Bundled application file was not found: {app_path}")

    if getattr(sys, "frozen", False):
        os.chdir(Path(sys.executable).resolve().parent)
    else:
        os.chdir(Path(__file__).resolve().parent)

    port = find_free_port()

    # Override any existing machine-level or environment-level Streamlit
    # settings that could point the browser to another port such as 3000.
    os.environ["STREAMLIT_SERVER_PORT"] = str(port)
    os.environ["STREAMLIT_BROWSER_SERVER_PORT"] = str(port)
    os.environ["STREAMLIT_SERVER_ADDRESS"] = "127.0.0.1"
    os.environ["STREAMLIT_BROWSER_SERVER_ADDRESS"] = "localhost"

    flag_options = {
        "global_developmentMode": False,
        "server_headless": False,
        "server_address": "127.0.0.1",
        "server_port": port,
        "browser_serverAddress": "localhost",
        "browser_serverPort": port,
        "server_fileWatcherType": "none",
        "browser_gatherUsageStats": False,
        "logger_hideWelcomeMessage": False,
    }

    # bootstrap.run does not itself apply CLI-style flag options when called
    # directly. Load them first so the listening port and browser URL match.
    bootstrap.load_config_options(flag_options)

    bootstrap.run(
        str(app_path),
        is_hello=False,
        args=[],
        flag_options=flag_options,
    )


if __name__ == "__main__":
    main()
