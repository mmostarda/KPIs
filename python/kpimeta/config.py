"""Settings (config/settings.toml), CSV helpers and per-user preferences."""

from __future__ import annotations

import csv
import io
import json
import os
import sys
from dataclasses import dataclass, field
from pathlib import Path

if sys.version_info >= (3, 11):
    import tomllib
else:  # pragma: no cover - Python 3.10
    import tomli as tomllib

PACKAGE_ROOT = Path(__file__).resolve().parent.parent
DEFAULT_CONFIG_DIR = PACKAGE_ROOT / "config"


class ConfigError(Exception):
    """Invalid configuration file (settings, schema or mapping)."""


@dataclass
class Settings:
    config_dir: Path
    schema_path: Path
    lola_mapping_path: Path
    metadata_sheet: str = "Metadata"
    metadata_table: str = "TableMet"
    kpi_sheet: str = "KPIs"
    kpi_table: str = "TableKPI"
    id_column: str = "ID"
    templates_sheet: str = "Metadata_Templates"
    template_id_column: str = "ID_Template"
    template_name_column: str = "IDName_Template"
    date_storage: str = "text"  # "text" | "excel"
    date_text_format: str = "%d/%m/%Y"
    strict_row_order: bool = True
    strict_choices: bool = False
    refuse_if_open: bool = True
    backup_dir: Path | None = None
    kpi_excluded_columns: list[str] = field(default_factory=lambda: ["ID_Template", "IDName_Template"])
    filename_separator: str = "_"
    filename_min_tokens: int = 7
    filename_date_format: str = "%Y%m%d"
    template_tokens: int = 2
    filename_tokens: dict[str, int] = field(default_factory=dict)
    lola_filename_header: str = "File Name"


def default_config_dir() -> Path:
    env = os.environ.get("KPIMETA_CONFIG")
    return Path(env) if env else DEFAULT_CONFIG_DIR


def load_settings(config_dir: str | Path | None = None) -> Settings:
    config_dir = Path(config_dir) if config_dir else default_config_dir()
    path = config_dir / "settings.toml"
    try:
        data = tomllib.loads(path.read_text(encoding="utf-8"))
    except FileNotFoundError as exc:
        raise ConfigError(f"File di impostazioni non trovato: {path}") from exc
    except tomllib.TOMLDecodeError as exc:
        raise ConfigError(f"{path.name} non valido: {exc}") from exc

    paths = data.get("paths", {})
    db = data.get("db", {})
    kpi = data.get("kpi", {})
    fname = data.get("filename", {})
    lola = data.get("lola", {})

    def resolve(value: str, default: str) -> Path:
        candidate = Path(value or default)
        return candidate if candidate.is_absolute() else config_dir / candidate

    settings = Settings(
        config_dir=config_dir,
        schema_path=resolve(paths.get("schema", ""), "schema_metadata.csv"),
        lola_mapping_path=resolve(paths.get("lola_mapping", ""), "lola_mapping.csv"),
    )
    for key in ("metadata_sheet", "metadata_table", "kpi_sheet", "kpi_table", "id_column",
                "templates_sheet", "template_id_column", "template_name_column",
                "date_storage", "date_text_format"):
        if key in db:
            setattr(settings, key, str(db[key]))
    for key in ("strict_row_order", "strict_choices", "refuse_if_open"):
        if key in db:
            setattr(settings, key, bool(db[key]))
    if db.get("backup_dir"):
        settings.backup_dir = Path(os.path.expandvars(str(db["backup_dir"]))).expanduser()
    if "excluded_columns" in kpi:
        settings.kpi_excluded_columns = [str(c) for c in kpi["excluded_columns"]]
    settings.filename_separator = str(fname.get("separator", settings.filename_separator))
    settings.filename_min_tokens = int(fname.get("min_tokens", settings.filename_min_tokens))
    settings.filename_date_format = str(fname.get("date_format", settings.filename_date_format))
    settings.template_tokens = int(fname.get("template_tokens", settings.template_tokens))
    settings.filename_tokens = {str(k): int(v) for k, v in fname.get("tokens", {}).items()}
    settings.lola_filename_header = str(lola.get("filename_header", settings.lola_filename_header))

    if settings.date_storage not in ("text", "excel"):
        raise ConfigError("db.date_storage deve essere \"text\" oppure \"excel\"")
    for column, position in settings.filename_tokens.items():
        if position < 1:
            raise ConfigError(f"filename.tokens.{column}: la posizione parte da 1")
    return settings


def read_csv_records(path: str | Path, required: list[str],
                     optional: tuple[str, ...] = ()) -> list[dict[str, str]]:
    """Read a ';' or ',' separated CSV (UTF-8, or Windows-1252 as saved by Excel).

    Header names are matched case-insensitively: `required` and `optional`
    columns are returned under their canonical name ("" when an optional
    column is absent). Rows with a wrong number of fields raise ConfigError
    with the line number.
    """
    path = Path(path)
    raw = path.read_bytes()
    for encoding in ("utf-8-sig", "cp1252"):
        try:
            text = raw.decode(encoding)
            break
        except UnicodeDecodeError:
            continue
    first_line = text.splitlines()[0] if text.strip() else ""
    delimiter = ";" if first_line.count(";") >= first_line.count(",") else ","
    reader = csv.reader(io.StringIO(text), delimiter=delimiter)
    try:
        header = [h.strip() for h in next(reader)]
    except StopIteration as exc:
        raise ConfigError(f"{path.name} è vuoto") from exc
    lookup = {h.lower(): h for h in header}
    missing = [name for name in required if name.lower() not in lookup]
    if missing:
        raise ConfigError(f"{path.name}: mancano le colonne {', '.join(missing)} (trovate: {', '.join(header)})")
    records = []
    for values in reader:
        if not any(v.strip() for v in values):
            continue
        if len(values) != len(header):
            raise ConfigError(
                f"{path.name}, riga {reader.line_num}: {len(values)} campi invece di {len(header)} "
                f"(controlla separatori '{delimiter}' e virgolette)")
        row = dict(zip(header, (v.strip() for v in values)))
        for name in (*required, *optional):
            row[name] = row.get(lookup.get(name.lower(), ""), "")
        records.append(row)
    return records


def user_data_dir() -> Path:
    """%LOCALAPPDATA%\\kpimeta on Windows, ~/.local/share/kpimeta elsewhere."""
    base = os.environ.get("LOCALAPPDATA") or os.environ.get("XDG_DATA_HOME")
    root = Path(base) if base else Path.home() / ".local" / "share"
    return root / "kpimeta"


def backup_root(settings: Settings) -> Path:
    return settings.backup_dir or (user_data_dir() / "backups")


def load_prefs() -> dict:
    path = user_data_dir() / "prefs.json"
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def save_prefs(**values: str) -> None:
    prefs = load_prefs()
    prefs.update({k: v for k, v in values.items() if v is not None})
    path = user_data_dir() / "prefs.json"
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(prefs, indent=2, ensure_ascii=False), encoding="utf-8")
    except OSError:
        pass  # preferences are a convenience only
