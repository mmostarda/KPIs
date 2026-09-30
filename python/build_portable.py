"""Portable copy of the app for colleagues: no installation, no admin rights, no network.

The build copies the Python that runs this script (without its packages), installs the
libraries of requirements.txt into the copy and adds the app and a launcher:

    KPI_metadata/
        Avvia KPI metadata.bat   double click: the app opens in the browser (http://127.0.0.1:8501)
        kpimeta.bat              command line (python -m kpimeta ...)
        LEGGIMI.txt
        app/                     app.py, kpimeta/, config/ (schema, settings, LoLa mapping)
        python/                  Python + libraries

Build it on Windows from the python/ folder, with a Python installed from python.org
(not the Microsoft Store one), e.g.:

    py -3.12 build_portable.py

Result: dist/KPI_metadata/ and dist/KPI_metadata.zip. Each colleague runs the app on their own
PC and the browser talks only to that PC (127.0.0.1): nothing is reachable from the network.
"""

from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent
WINDOWS = os.name == "nt"
APP_ITEMS = ("app.py", "kpimeta", "config")
OPTIONAL_REQUIREMENTS = ("scipy",)  # only to convert the old .mat mapping

SERVER_CONFIG = """
[server]
# served to this PC only: not reachable from the network, no firewall prompt. 127.0.0.1 and
# not "localhost": on Windows localhost is tried first as IPv6 (::1), where nothing listens
address = "127.0.0.1"
showEmailPrompt = false
fileWatcherType = "none"
"""

LAUNCHER_BAT = r"""@echo off
setlocal
title KPI metadata
rem Portable copy: only the Python and the libraries of this folder are used
set "PYTHONHOME="
set "PYTHONPATH="
set "PYTHONNOUSERSITE=1"
rem pushd (not cd): it also works when the folder is on a network share (\\server\...)
pushd "%~dp0app"
echo KPI metadata: il browser si apre da solo tra qualche secondo.
echo Per chiudere l'app chiudi questa finestra.
echo.
"%~dp0python\python.exe" -m streamlit run app.py
if errorlevel 1 pause
"""

CLI_BAT = r"""@echo off
setlocal
set "PYTHONHOME="
set "PYTHONNOUSERSITE=1"
set "PYTHONPATH=%~dp0app"
"%~dp0python\python.exe" -m kpimeta %*
"""

LAUNCHER_SH = """#!/bin/sh
# Portable copy: only the Python and the libraries of this folder are used
unset PYTHONHOME PYTHONPATH
export PYTHONNOUSERSITE=1
cd "$(dirname "$0")/app" || exit 1
exec ../python/bin/python3 -m streamlit run app.py
"""

README = """KPI metadata - versione portatile
=================================

Non serve installare nulla e non servono permessi di amministratore.

1. Copia la cartella {name} sul tuo PC (per esempio in Documenti).
   Da una cartella di rete funziona, ma l'avvio è molto più lento.
2. Doppio clic su "Avvia KPI metadata.bat".
   Si apre una finestra nera: lasciala aperta. Dopo qualche secondo il browser mostra l'app.
   L'indirizzo http://127.0.0.1:8501 è servito dal tuo PC: non passa dalla rete
   (se hai già l'app aperta, la seconda usa 8502).
3. Per chiudere l'app chiudi la finestra nera.

Il primo avvio può richiedere fino a un minuto (controllo dell'antivirus): la pagina
mostra "Avvio in corso". Se resta vuota, attendi e premi F5.
Se qualcosa non va, la finestra nera mostra il messaggio d'errore.

File utili
- app\\config\\settings.toml        impostazioni (fogli, tabelle, date, nome file)
- app\\config\\schema_metadata.csv  i metadati del form
- app\\config\\lola_mapping.csv     il mapping LoLa -> DB
- %LOCALAPPDATA%\\kpimeta\\backups  i backup del DB fatti prima di ogni scrittura
- kpimeta.bat                    la riga di comando (kpimeta.bat doctor percorso\\DB.xlsx)

Versione creata il {date} con Python {python}.
"""


def fail(message: str) -> None:
    print(f"ERRORE: {message}", file=sys.stderr)
    raise SystemExit(2)


def base_python() -> Path:
    """Installation of the running Python (the base one when run from a venv)."""
    base = Path(sys.base_prefix)
    if sys.version_info < (3, 10):
        fail("serve Python 3.10 o superiore")
    if "WindowsApps" in str(base) or "WindowsApps" in sys.executable:
        fail("il Python del Microsoft Store non si puo' copiare: usa un Python installato da python.org")
    if (base / "conda-meta").exists():
        fail("Python di Anaconda/conda: usa un Python installato da python.org")
    return base


def stdlib_dir(base: Path) -> Path:
    return base / "Lib" if WINDOWS else base / "lib" / f"python{sys.version_info[0]}.{sys.version_info[1]}"


def python_exe(root: Path) -> Path:
    return root / "python.exe" if WINDOWS else root / "bin" / "python3"


def copy_python(base: Path, target: Path) -> Path:
    """Copy the Python installation without its packages, tests, docs and headers."""
    stdlib = stdlib_dir(base)
    skip = {base / name for name in ("Doc", "Tools", "Scripts", "include", "libs", "share")}
    skip |= {stdlib / "site-packages", stdlib / "test", base / "lib" / "pkgconfig"}

    def ignore(directory: str, names: list[str]) -> set[str]:
        here = Path(directory)
        return {n for n in names if here / n in skip or (here == stdlib and n.startswith("config-"))}

    shutil.copytree(base, target, ignore=ignore, symlinks=True)
    (stdlib_dir(target)).joinpath("site-packages").mkdir(exist_ok=True)
    exe = python_exe(target)
    if not exe.exists():
        fail(f"eseguibile Python non trovato nella copia: {exe}")
    return exe


def clean_env() -> dict[str, str]:
    env = {k: v for k, v in os.environ.items()
           if k not in ("PYTHONHOME", "PYTHONPATH", "PIP_REQUIRE_VIRTUALENV", "VIRTUAL_ENV")}
    env["PYTHONNOUSERSITE"] = "1"
    return env


def run(command: list[str | Path]) -> None:
    print("  $", " ".join(str(c) for c in command))
    subprocess.run([str(c) for c in command], check=True, env=clean_env())


def install_requirements(exe: Path, with_optional: bool) -> None:
    lines = (ROOT / "requirements.txt").read_text(encoding="utf-8").splitlines()
    keep = [line for line in lines
            if with_optional or not line.strip().lower().startswith(OPTIONAL_REQUIREMENTS)]
    with tempfile.TemporaryDirectory() as tmp:
        requirements = Path(tmp) / "requirements.txt"
        requirements.write_text("\n".join(keep) + "\n", encoding="utf-8")
        run([exe, "-m", "ensurepip", "--default-pip"])
        run([exe, "-m", "pip", "install", "--disable-pip-version-check", "--no-warn-script-location",
             "-r", requirements])


def remove_package_tests(root: Path) -> int:
    """Drop the test suites shipped inside the libraries (pandas, numpy, pyarrow...): never imported
    by the app, about 10% of the size."""
    site_packages = stdlib_dir(root) / "site-packages"
    removed = 0
    for folder in sorted(site_packages.glob("*/**/tests"), key=lambda p: len(p.parts)):
        if folder.is_dir() and folder.exists():
            removed += folder_size(folder)
            shutil.rmtree(folder)
    return removed


def copy_app(target: Path) -> None:
    target.mkdir(parents=True)
    for item in APP_ITEMS:
        source = ROOT / item
        if source.is_dir():
            shutil.copytree(source, target / item, ignore=shutil.ignore_patterns("__pycache__", "*.pyc"))
        else:
            shutil.copy2(source, target / item)
    config = (ROOT / ".streamlit" / "config.toml").read_text(encoding="utf-8")
    if "[server]" in config:
        fail(".streamlit/config.toml ha gia' una sezione [server]: unisci a mano le impostazioni del build")
    (target / ".streamlit").mkdir()
    (target / ".streamlit" / "config.toml").write_text(config.rstrip() + "\n" + SERVER_CONFIG, encoding="utf-8")


def write_launchers(bundle: Path, name: str) -> None:
    import datetime as dt

    if WINDOWS:
        (bundle / "Avvia KPI metadata.bat").write_text(LAUNCHER_BAT, encoding="ascii", newline="\r\n")
        (bundle / "kpimeta.bat").write_text(CLI_BAT, encoding="ascii", newline="\r\n")
    else:
        launcher = bundle / "avvia.sh"
        launcher.write_text(LAUNCHER_SH, encoding="utf-8")
        launcher.chmod(0o755)
    version = ".".join(str(v) for v in sys.version_info[:3])
    (bundle / "LEGGIMI.txt").write_text(README.format(name=name, date=dt.date.today().strftime("%d/%m/%Y"),
                                                      python=version),
                                        encoding="utf-8-sig", newline="\r\n" if WINDOWS else "\n")


def folder_size(path: Path) -> int:
    return sum(f.stat().st_size for f in path.rglob("*") if f.is_file() and not f.is_symlink())


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Crea la versione portatile dell'app (cartella + zip)")
    parser.add_argument("--output", type=Path, default=ROOT / "dist", help="cartella di destinazione")
    parser.add_argument("--name", default="KPI_metadata", help="nome della cartella e dello zip")
    parser.add_argument("--with-scipy", action="store_true",
                        help="includi scipy (serve solo per convertire il vecchio mapping .mat)")
    parser.add_argument("--no-zip", action="store_true", help="non creare lo zip")
    args = parser.parse_args(argv)

    base = base_python()
    bundle = args.output / args.name
    if bundle.exists():
        shutil.rmtree(bundle)
    bundle.mkdir(parents=True)
    if not (ROOT / "config" / "lola_mapping.csv").exists():
        print("ATTENZIONE: config/lola_mapping.csv non esiste: i colleghi non potranno leggere gli output LoLa "
              "finche' non lo aggiungono in app/config (python -m kpimeta convert-mapping ...)")

    print(f"1/4 copia di Python {sys.version.split()[0]} da {base}")
    exe = copy_python(base, bundle / "python")
    print("2/4 installazione delle librerie")
    install_requirements(exe, args.with_scipy)
    print(f"   rimossi {remove_package_tests(bundle / 'python') / 1e6:.0f} MB di test delle librerie")
    print("3/4 copia dell'app")
    copy_app(bundle / "app")
    run([exe, "-m", "compileall", "-q", bundle / "app"])
    write_launchers(bundle, args.name)
    print(f"   cartella pronta: {bundle} ({folder_size(bundle) / 1e6:.0f} MB)")
    if not args.no_zip:
        print("4/4 creazione dello zip")
        archive = shutil.make_archive(str(args.output / args.name), "zip", root_dir=args.output, base_dir=args.name)
        print(f"   {archive} ({Path(archive).stat().st_size / 1e6:.0f} MB)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
