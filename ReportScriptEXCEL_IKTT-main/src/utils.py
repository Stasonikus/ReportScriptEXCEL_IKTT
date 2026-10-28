from __future__ import annotations

import sys
from pathlib import Path


def resource_path(relative_path: str) -> Path:
    """
    Returns a bundled resource path both in .py mode and in PyInstaller .exe mode.
    """
    try:
        base_path = Path(sys._MEIPASS)  # type: ignore[attr-defined]
    except Exception:
        base_path = Path(__file__).resolve().parent

    return base_path / relative_path


def app_base_path() -> Path:
    """
    Returns the folder where runtime files should live.

    For .exe this is the folder with the executable. For .py this is the project
    root, so in/out stay next to src instead of inside src.
    """
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent

    return Path(__file__).resolve().parent.parent


def app_path(relative_path: str) -> Path:
    return app_base_path() / relative_path
