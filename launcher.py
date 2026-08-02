"""Windows launcher for the bundled Streamlit catalogue application."""

from __future__ import annotations

import os
from pathlib import Path
import socket
import sys

# These imports make the application dependencies visible to PyInstaller.
import openpyxl  # noqa: F401
import pandas  # noqa: F401
import pyarrow  # noqa: F401
import streamlit  # noqa: F401

import data_utils  # noqa: F401
from streamlit.web import bootstrap


def bundled_path(filename: str) -> Path:
    """Return a project file path in source or PyInstaller bundle mode."""
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS) / filename
    return Path(__file__).resolve().parent / filename


def available_local_port() -> int:
    """Ask Windows for an unused local TCP port."""
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.bind(("127.0.0.1", 0))
        return int(sock.getsockname()[1])


def main() -> None:
    """Start the bundled Streamlit application in the default browser."""
    app_path = bundled_path("app.py")
    if not app_path.is_file():
        raise FileNotFoundError(f"Bundled application file was not found: {app_path}")

    # Keep any relative paths and Streamlit settings local to the EXE folder.
    if getattr(sys, "frozen", False):
        os.chdir(Path(sys.executable).resolve().parent)
    else:
        os.chdir(Path(__file__).resolve().parent)

    port = available_local_port()

    flag_options = {
        "global_developmentMode": False,
        "server_headless": False,
        "server_address": "127.0.0.1",
        "server_port": port,
        "server_fileWatcherType": "none",
        "browser_gatherUsageStats": False,
        "logger_hideWelcomeMessage": True,
    }

    bootstrap.run(
        str(app_path),
        is_hello=False,
        args=[],
        flag_options=flag_options,
    )


if __name__ == "__main__":
    main()
