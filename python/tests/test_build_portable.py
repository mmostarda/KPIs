"""build_portable.py without network: app copy, launchers, pruning (the full build is tried by hand)."""

try:
    import tomllib
except ModuleNotFoundError:  # Python 3.10
    import tomli as tomllib

import build_portable as bp


def test_app_copy_is_served_to_this_pc_only(tmp_path):
    bp.copy_app(tmp_path / "app")
    config = (tmp_path / "app" / ".streamlit" / "config.toml").read_text(encoding="utf-8")
    server = tomllib.loads(config)["server"]
    assert server["address"] == "127.0.0.1" and server["showEmailPrompt"] is False
    assert (tmp_path / "app" / "kpimeta" / "__init__.py").exists()
    assert (tmp_path / "app" / "config" / "schema_metadata.csv").exists()
    assert not list((tmp_path / "app").rglob("__pycache__"))


def test_windows_launchers(tmp_path, monkeypatch):
    monkeypatch.setattr(bp, "WINDOWS", True)
    bp.write_launchers(tmp_path, "KPI_metadata")
    bat = (tmp_path / "Avvia KPI metadata.bat").read_bytes()
    assert b"\r\n" in bat and b'pushd "%~dp0app"' in bat  # pushd: works on network shares too
    assert b'"%~dp0python\\python.exe" -m streamlit run app.py' in bat
    assert b'set "PYTHONPATH=%~dp0app"' in (tmp_path / "kpimeta.bat").read_bytes()
    assert (tmp_path / "LEGGIMI.txt").read_bytes().startswith(b"\xef\xbb\xbf")  # UTF-8 for Notepad


def test_library_tests_are_removed(tmp_path, monkeypatch):
    monkeypatch.setattr(bp, "WINDOWS", True)
    site = tmp_path / "Lib" / "site-packages"
    (site / "pandas" / "tests" / "io").mkdir(parents=True)
    (site / "pandas" / "tests" / "io" / "test_x.py").write_text("x" * 100)
    (site / "numpy" / "_core" / "tests").mkdir(parents=True)
    (site / "pandas" / "core").mkdir(parents=True)
    (site / "streamlit" / "testing").mkdir(parents=True)
    assert bp.remove_package_tests(tmp_path) == 100
    assert not (site / "pandas" / "tests").exists() and not (site / "numpy" / "_core" / "tests").exists()
    assert (site / "pandas" / "core").exists() and (site / "streamlit" / "testing").exists()
