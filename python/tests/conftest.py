from __future__ import annotations

from pathlib import Path

import pytest

from kpimeta.config import load_settings
from kpimeta.demo import make_demo
from kpimeta.schema import load_schema


@pytest.fixture()
def settings(tmp_path, monkeypatch):
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path / "appdata"))
    s = load_settings()
    s.backup_dir = tmp_path / "backups"
    return s


@pytest.fixture()
def schema(settings):
    return load_schema(settings.schema_path)


@pytest.fixture()
def demo(tmp_path, settings) -> dict:
    return make_demo(tmp_path / "demo")


@pytest.fixture()
def db_path(demo) -> Path:
    return demo["db"]
