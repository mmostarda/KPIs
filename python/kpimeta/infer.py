"""Metadata inferred from the log file name (port of step 8b and 6c of
writeMetadata_ALL_DB_FromGUI).

Current rule (MATLAB): X_X_Date_X_DrivingMode_Maneuvre_Driver_X.mf4, tokens
split on "_": token 3 = Date (YYYYMMDD), 5 = Driving_Mode, 6 = ManeuvreName,
7 = Driver; the first 2 tokens are the template name. Positions and separator
come from settings.toml.
"""

from __future__ import annotations

import datetime as dt
from pathlib import PureWindowsPath

from .config import Settings
from .schema import Schema
from .values import to_text


def stem(filename: str) -> str:
    """Base name without folder and last extension (MATLAB fileparts)."""
    name = PureWindowsPath(str(filename).strip()).name  # handles both / and \
    return name.rsplit(".", 1)[0] if "." in name else name


def tokens(filename: str, settings: Settings) -> list[str]:
    return stem(filename).split(settings.filename_separator)


def infer_from_filename(filename: object, settings: Settings,
                        schema: Schema) -> tuple[dict[str, object], str | None]:
    """({column: value}, problem). An empty dict means nothing could be inferred."""
    text = to_text(filename)
    if not text:
        return {}, "nome file vuoto"
    parts = tokens(text, settings)
    if len(parts) < settings.filename_min_tokens:
        return {}, (f"'{stem(text)}' ha {len(parts)} token separati da "
                    f"'{settings.filename_separator}', ne servono almeno {settings.filename_min_tokens}")
    values: dict[str, object] = {}
    problems = []
    for column, position in settings.filename_tokens.items():
        if position > len(parts):
            problems.append(f"token {position} assente per {column}")
            continue
        token = parts[position - 1].strip()
        if schema.type_of(column) == "date":
            try:
                values[column] = dt.datetime.strptime(token, settings.filename_date_format).date()
            except ValueError:
                problems.append(f"'{token}' non è una data {settings.filename_date_format} per {column}")
            continue
        values[column] = token
    return values, ("; ".join(problems) or None)


def template_name(filename: object, settings: Settings) -> str | None:
    text = to_text(filename)
    if not text:
        return None
    parts = tokens(text, settings)
    if len(parts) >= settings.template_tokens:
        return settings.filename_separator.join(parts[:settings.template_tokens])
    return stem(text)
