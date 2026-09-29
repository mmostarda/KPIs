"""LoLa -> DB column mapping (replaces Mapping_LoLa_DB_mapping.mat).

Sources accepted: the CSV used by kpimeta, the original Excel file read by
load_mapping_from_excel.m (columns KPI_DB_Name / KPI_LoLa_Name) and the .mat
file produced by it (struct `mapping` with LoLa_Name / DB_Name cell arrays).
"""

from __future__ import annotations

import csv
from dataclasses import dataclass
from pathlib import Path

from .config import ConfigError, read_csv_records
from .values import is_missing, to_text
from .xlsx import open_workbook

LOLA_COL, DB_COL, TYPE_COL = "KPI_LoLa_Name", "KPI_DB_Name", "Type"
TYPES = ("auto", "text", "number", "date")

MATLAB_EXPORT_HINT = (
    "In MATLAB:\n"
    "  load('Mapping_LoLa_DB_mapping.mat');\n"
    "  T = table(mapping.LoLa_Name(:), mapping.DB_Name(:), "
    "'VariableNames', {'KPI_LoLa_Name','KPI_DB_Name'});\n"
    "  writetable(T, 'lola_mapping.csv');"
)


class MappingError(ConfigError):
    pass


@dataclass(frozen=True)
class MappingEntry:
    lola_name: str
    db_name: str
    type: str = "auto"


@dataclass
class LolaMapping:
    entries: list[MappingEntry]
    source: str = ""
    skipped_rows: int = 0

    def db_names(self) -> list[str]:
        return [e.db_name for e in self.entries]


def _is_invalid(value: object) -> bool:
    # load_mapping_from_excel.m skipped empty, NA, NaN and <undefined>
    return is_missing(value)


def build_mapping(pairs: list[tuple[object, object, object]], source: str) -> LolaMapping:
    entries: list[MappingEntry] = []
    skipped = 0
    for lola, db, typ in pairs:
        if _is_invalid(lola) or _is_invalid(db):
            skipped += 1
            continue
        typ_text = (to_text(typ) or "auto").lower()
        if typ_text not in TYPES:
            raise MappingError(f"{source}: tipo '{typ_text}' non valido per '{lola}' ({', '.join(TYPES)})")
        entries.append(MappingEntry(to_text(lola), to_text(db), typ_text))
    if not entries:
        raise MappingError(f"{source}: nessuna coppia valida LoLa -> DB")

    problems = []
    by_lola: dict[str, str] = {}
    by_db: dict[str, str] = {}
    for e in entries:
        key = e.lola_name.strip().lower()
        if key in by_lola:
            problems.append(f"nome LoLa '{e.lola_name}' ripetuto")
        by_lola[key] = e.db_name
        if e.db_name in by_db:
            problems.append(f"nome DB '{e.db_name}' usato per '{by_db[e.db_name]}' e '{e.lola_name}'")
        by_db[e.db_name] = e.lola_name
    if problems:
        raise MappingError(f"{source}: " + "; ".join(problems))
    return LolaMapping(entries, source=source, skipped_rows=skipped)


def _from_csv(path: Path) -> LolaMapping:
    records = read_csv_records(path, required=[LOLA_COL, DB_COL], optional=(TYPE_COL,))
    return build_mapping([(r[LOLA_COL], r[DB_COL], r[TYPE_COL]) for r in records], path.name)


def _from_excel(path: Path) -> LolaMapping:
    wb = open_workbook(path, MappingError, read_only=True, data_only=True)
    try:
        rows = list(wb.worksheets[0].iter_rows(values_only=True))
    finally:
        wb.close()
    if not rows:
        raise MappingError(f"{path.name} è vuoto")
    header = [str(h).strip().lower() if h is not None else "" for h in rows[0]]
    try:
        i_lola = header.index(LOLA_COL.lower())
        i_db = header.index(DB_COL.lower())
    except ValueError as exc:
        raise MappingError(f"{path.name}: servono le colonne {LOLA_COL} e {DB_COL}") from exc
    i_type = header.index(TYPE_COL.lower()) if TYPE_COL.lower() in header else None
    pairs = []
    for row in rows[1:]:
        row = list(row) + [None] * (len(header) - len(row))
        pairs.append((row[i_lola], row[i_db], row[i_type] if i_type is not None else None))
    return build_mapping(pairs, path.name)


def _from_mat(path: Path) -> LolaMapping:
    try:
        from scipy.io import loadmat
    except ImportError as exc:
        raise MappingError("Per leggere il .mat serve scipy (pip install scipy).\n" + MATLAB_EXPORT_HINT) from exc
    try:
        data = loadmat(path, squeeze_me=True, struct_as_record=False, chars_as_strings=True)
    except NotImplementedError as exc:  # MATLAB v7.3 (HDF5) files
        raise MappingError(f"{path.name} è in formato MATLAB v7.3, non leggibile da scipy.\n"
                           + MATLAB_EXPORT_HINT) from exc
    if "mapping" not in data:
        raise MappingError(f"{path.name}: variabile 'mapping' non trovata")
    mapping = data["mapping"]
    try:
        lola, db = mapping.LoLa_Name, mapping.DB_Name
    except AttributeError as exc:
        raise MappingError(f"{path.name}: 'mapping' deve avere i campi LoLa_Name e DB_Name") from exc

    def as_list(value) -> list:
        if isinstance(value, str):
            return [value]
        return [v if isinstance(v, str) else (None if getattr(v, "size", 1) == 0 else str(v))
                for v in list(value.ravel())]

    lola_list, db_list = as_list(lola), as_list(db)
    if len(lola_list) != len(db_list):
        raise MappingError(f"{path.name}: LoLa_Name e DB_Name hanno lunghezze diverse")
    return build_mapping([(a, b, None) for a, b in zip(lola_list, db_list)], path.name)


def load_mapping(path: str | Path) -> LolaMapping:
    path = Path(path)
    if not path.exists():
        raise MappingError(f"Mapping LoLa non trovato: {path}\n"
                           "Crealo con: python -m kpimeta convert-mapping <file .mat/.xlsx> "
                           "-o config/lola_mapping.csv")
    suffix = path.suffix.lower()
    if suffix == ".csv":
        return _from_csv(path)
    if suffix in (".xlsx", ".xlsm"):
        return _from_excel(path)
    if suffix == ".mat":
        return _from_mat(path)
    raise MappingError(f"Formato mapping non supportato: {path.name} (usa .csv, .xlsx o .mat)")


def save_mapping_csv(mapping: LolaMapping, path: str | Path) -> Path:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.writer(handle, delimiter=";")
        writer.writerow([LOLA_COL, DB_COL, TYPE_COL])
        for e in mapping.entries:
            writer.writerow([e.lola_name, e.db_name, "" if e.type == "auto" else e.type])
    return path
