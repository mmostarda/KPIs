"""The Excel DB read by Power BI: reading, checks ("doctor") and safe writing.

The DB file keeps its name, sheets, tables (TableMet / TableKPI) and columns.
Every write follows the same procedure:
  1. refuse if the file is open in Excel (lock file ~$... or file locked);
  2. copy the DB into the backup folder;
  3. load it with openpyxl, apply ALL the changes in memory, save to a temporary
     file in the same folder, re-open it and verify tables, IDs and internal parts;
  4. replace the original only if nobody modified it meanwhile (atomic os.replace,
     the same technique Excel uses when saving).
If any step fails the original file is untouched. Compared to the MATLAB app
(one hidden Excel instance per row, no transaction) a batch of any size is one
single save, without Excel installed.
"""

from __future__ import annotations

import copy
import datetime as dt
import os
import posixpath
import re
import shutil
import time
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable

import pandas as pd
from openpyxl import load_workbook
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.utils.cell import column_index_from_string, get_column_letter, range_boundaries
from openpyxl.worksheet.table import TableColumn

from .batch import Batch
from .config import Settings, backup_root
from .issues import ERROR, INFO, WARNING, Issue, errors
from .schema import Schema
from .validate import KpiDecision, check_decisions, duplicates_vs_db, kpi_resolution, validate_batch
from .values import (MIN_YEAR, MAX_YEAR, ConversionError, is_missing,
                     normalize_name, plain, to_date, to_int, to_number, to_text)
from .xlsx import invalid_file_message, open_workbook


class DBError(Exception):
    pass


class DBLockedError(DBError):
    pass


class DBChangedError(DBError):
    pass


# ---------------------------------------------------------------------------
# Package inspection (what openpyxl could lose when saving)
# ---------------------------------------------------------------------------

_NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_NS_PKG = "http://schemas.openxmlformats.org/package/2006/relationships"

# categories checked before/after saving: if present before and missing after -> abort
_PART_CATEGORIES = (
    ("grafici", r"^xl/charts/"),
    ("immagini", r"^xl/media/"),
    ("forme o disegni", r"^xl/drawings/drawing\d+\.xml$"),
    ("tabelle pivot", r"^xl/pivot(Tables|Cache)/"),
    ("modello dati Power Pivot", r"^xl/model/"),
    ("slicer o timeline", r"^xl/(slicers|slicerCaches|timelines|timelineCaches)/"),
    ("connessioni dati", r"^xl/(connections\.xml$|queryTables/)"),
    ("collegamenti esterni", r"^xl/externalLinks/"),
    ("macro VBA", r"^xl/vbaProject\.bin$"),
    ("commenti", r"^xl/comments\d*\.xml$|^xl/comments/"),
    ("commenti moderni (thread)", r"^xl/threadedComments/"),
    ("controlli o oggetti incorporati", r"^xl/(ctrlProps|activeX|embeddings)/"),
    ("proprietà personalizzate (es. etichette di riservatezza)", r"^docProps/custom\.xml$"),
    ("tabelle Excel", r"^xl/tables/"),
)
_REWRITTEN = ("grafici", "immagini")  # kept by openpyxl but re-serialized
_X14_FEATURES = ((b"x14:dataValidation", "convalide dati con riferimenti ad altri fogli"),
                 (b"x14:sparklineGroup", "sparkline"),
                 (b"x14:conditionalFormatting", "formattazioni condizionali avanzate"))
_FORMULA_CELL_RE = re.compile(rb'<c\s[^>]*?\br="([A-Z]+[0-9]+)"[^>]*>\s*<f[\s>/]')


@dataclass
class TableDef:
    """Definition of an Excel table read straight from xl/tables/*.xml."""
    name: str
    sheet: str
    ref: str
    columns: list[str]
    column_ids: list[int]
    header_rows: int = 1
    totals_rows: int = 0
    calculated: list[str] = field(default_factory=list)


@dataclass
class PackageScan:
    parts: dict[str, int]  # category -> number of parts
    customxml: int = 0
    power_query: bool = False
    x14: dict[str, list[str]] = field(default_factory=dict)  # feature -> sheets
    formulas: dict[str, list[str]] = field(default_factory=dict)  # sheet -> cell coordinates
    tables: dict[str, TableDef] = field(default_factory=dict)  # lower-case name -> definition

    def tables_on(self, sheet: str) -> list[TableDef]:
        return [t for t in self.tables.values() if t.sheet == sheet]


def _sheet_parts(archive: zipfile.ZipFile) -> dict[str, str]:
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    targets = {r.get("Id"): r.get("Target", "") for r in rels.findall(f"{{{_NS_PKG}}}Relationship")}
    parts = {}
    for sheet in workbook.iter(f"{{{_NS_MAIN}}}sheet"):
        target = targets.get(sheet.get(f"{{{_NS_REL}}}id"), "")
        part = target.lstrip("/") if target.startswith("/") else posixpath.normpath("xl/" + target)
        parts[sheet.get("name")] = part
    return parts


def _resolve_part(base_part: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(base_part), target))


def _read_table_defs(archive: zipfile.ZipFile, sheet: str, part: str, names: set[str]) -> list[TableDef]:
    rels_part = posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels")
    if rels_part not in names:
        return []
    out = []
    for rel in ET.fromstring(archive.read(rels_part)).findall(f"{{{_NS_PKG}}}Relationship"):
        if not rel.get("Type", "").endswith("/table"):
            continue
        table_part = _resolve_part(part, rel.get("Target", ""))
        if table_part not in names:
            continue
        root = ET.fromstring(archive.read(table_part))
        columns, ids, calculated = [], [], []
        container = root.find(f"{{{_NS_MAIN}}}tableColumns")
        for column in (container if container is not None else []):
            columns.append(column.get("name", ""))
            ids.append(int(column.get("id", "0")))
            if column.find(f"{{{_NS_MAIN}}}calculatedColumnFormula") is not None:
                calculated.append(column.get("name", ""))
        out.append(TableDef(root.get("displayName") or root.get("name") or "", sheet, root.get("ref", ""),
                            columns, ids, int(root.get("headerRowCount", "1")),
                            int(root.get("totalsRowCount", "0")), calculated))
    return out


def scan_package(path: str | Path) -> PackageScan:
    """One pass over the xlsx package: parts, formulas, table definitions."""
    try:
        archive = zipfile.ZipFile(path)
    except zipfile.BadZipFile as exc:
        raise DBError(invalid_file_message(path)) from exc
    with archive:
        names = archive.namelist()
        name_set = set(names)
        if "xl/workbook.xml" not in name_set:
            raise DBError(invalid_file_message(path))
        parts = {label: sum(1 for n in names if re.search(pattern, n)) for label, pattern in _PART_CATEGORIES}
        custom = [n for n in names if n.startswith("customXml/") and n.endswith(".xml")
                  and "/_rels/" not in n]
        power_query = any(b"DataMashup" in archive.read(n) for n in custom)
        scan = PackageScan({k: v for k, v in parts.items() if v}, len(custom), power_query)
        for sheet, part in _sheet_parts(archive).items():
            if part not in name_set:
                continue
            xml = archive.read(part)
            for marker, label in _X14_FEATURES:
                if marker in xml:
                    scan.x14.setdefault(label, []).append(sheet)
            cells = [c.decode() for c in _FORMULA_CELL_RE.findall(xml)]
            if cells:
                scan.formulas[sheet] = cells
            for table in _read_table_defs(archive, sheet, part, name_set):
                scan.tables[table.name.lower()] = table
    return scan


def lost_features(before: PackageScan, after: PackageScan) -> tuple[list[str], list[str]]:
    """(blocking losses, non-blocking losses) between two scans."""
    blocking = [label for label, count in before.parts.items()
                if count and after.parts.get(label, 0) < (count if label == "tabelle Excel" else 1)]
    if before.power_query and not after.power_query:
        blocking.append("query Power Query")
    soft = [f"{label} ({', '.join(sheets)})" for label, sheets in before.x14.items() if label not in after.x14]
    if before.customxml and not after.customxml and not before.power_query:
        soft.append("metadati customXml (es. proprietà SharePoint)")
    return blocking, soft


# ---------------------------------------------------------------------------
# Lock and backup
# ---------------------------------------------------------------------------

def lock_status(path: str | Path) -> str | None:
    """Why the file looks open in Excel, or None."""
    path = Path(path)
    name = path.name
    try:
        owners = list(path.parent.glob("~$*"))
    except OSError:
        owners = []
    for owner in owners:
        rest = owner.name[2:]
        if rest and (rest == name or (name.endswith(rest) and len(name) - len(rest) <= 2)):
            return f"il file è aperto in Excel (trovato il file di lock {owner.name})"
    if os.name == "nt" and path.exists():
        try:
            handle = os.open(str(path), os.O_RDWR | getattr(os, "O_BINARY", 0))
            os.close(handle)
        except PermissionError:
            return "il file è aperto da un altro programma (scrittura negata)"
        except OSError:
            pass
    return None


def make_backup(path: Path, settings: Settings) -> Path:
    folder = backup_root(settings) / path.stem
    folder.mkdir(parents=True, exist_ok=True)
    stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    target = folder / f"{path.stem}_{stamp}{path.suffix}"
    shutil.copy2(path, target)
    return target


# ---------------------------------------------------------------------------
# Reading tables
# ---------------------------------------------------------------------------

@dataclass
class TableInfo:
    name: str
    sheet: str
    headers: list[str]
    df: pd.DataFrame
    min_col: int
    max_col: int
    header_row: int
    ref_last_row: int  # last row inside the table range (totals row excluded)
    last_used_row: int  # last non-empty data row (header_row if the table is empty)
    totals_rows: int = 0
    calculated_columns: list[str] = field(default_factory=list)
    header_mismatch: list[str] = field(default_factory=list)

    @property
    def first_data_row(self) -> int:
        return self.header_row + 1

    @property
    def ref(self) -> str:
        return (f"{get_column_letter(self.min_col)}{self.header_row}:"
                f"{get_column_letter(self.max_col)}{self.ref_last_row + self.totals_rows}")

    def column_of(self, header: str) -> int:
        return self.min_col + self.headers.index(header)


def _cell_text(value: Any) -> str:
    return to_text(plain(value)) or ""


def _find_table(wb, name: str, preferred_sheet: str | None = None):
    if preferred_sheet in wb.sheetnames and name in wb[preferred_sheet].tables:
        ws = wb[preferred_sheet]
        return ws, ws.tables[name]
    for ws in wb.worksheets:
        for table in ws.tables.values():
            if table.displayName.lower() == name.lower():
                return ws, table
    raise DBError(f"tabella Excel '{name}' non trovata nel file (fogli: {', '.join(wb.sheetnames)})")


def _range_values(ws, min_row: int, max_row: int, min_col: int, max_col: int) -> list[list[Any]]:
    if max_row < min_row:
        return []
    return [[plain(v) for v in row] for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col,
                                                            max_col=max_col, values_only=True)]


def _info_from_rows(ws, name: str, ref: str, names: list[str], totals: int, calculated: list[str]) -> TableInfo:
    """Read a table range from a normal or read-only worksheet."""
    min_col, min_row, max_col, max_row = range_boundaries(ref)
    header = _range_values(ws, min_row, min_row, min_col, max_col)
    cells = [_cell_text(v) for v in (header[0] if header else [])]
    cells += [""] * (max_col - min_col + 1 - len(cells))
    headers = names if len(names) == max_col - min_col + 1 else cells
    mismatch = [f"{n}/{c}" for n, c in zip(names, cells) if n != c]
    rows = _range_values(ws, min_row + 1, max_row - totals, min_col, max_col)
    used = len(rows)
    while used and all(is_missing(v) for v in rows[used - 1]):
        used -= 1
    frame = pd.DataFrame(rows[:used], columns=headers, dtype=object)
    return TableInfo(name, ws.title, headers, frame, min_col, max_col, min_row, max_row - totals,
                     min_row + used, totals, calculated, mismatch)


def _table_info(ws, table) -> TableInfo:
    """Table of a workbook opened for writing (openpyxl objects)."""
    if (1 if table.headerRowCount is None else int(table.headerRowCount)) != 1:
        raise DBError(f"la tabella '{table.displayName}' non ha una riga di intestazione")
    calculated = [tc.name for tc in table.tableColumns if getattr(tc, "calculatedColumnFormula", None)]
    return _info_from_rows(ws, table.displayName, table.ref, [tc.name for tc in table.tableColumns],
                           int(table.totalsRowCount or 0), calculated)


def _table_info_ro(wb, table: TableDef) -> TableInfo:
    """Table of a workbook opened read-only (definition parsed from the package)."""
    if table.header_rows != 1:
        raise DBError(f"la tabella '{table.name}' non ha una riga di intestazione")
    return _info_from_rows(wb[table.sheet], table.name, table.ref, table.columns, table.totals_rows,
                           table.calculated)


def _sheet_frame(wb, scan: PackageScan, sheet: str,
                 prefer_table: str | None = None) -> tuple[list[str], pd.DataFrame, str | None]:
    """A support/template sheet: its single Excel table, or the range with headers on row 1."""
    tables = scan.tables_on(sheet)
    table = tables[0] if len(tables) == 1 else next(
        (t for t in tables if prefer_table and t.name.lower() == prefer_table.lower()), None)
    if table is not None:
        info = _table_info_ro(wb, table)
        return info.headers, info.df, table.name
    rows = [[plain(v) for v in r] for r in wb[sheet].iter_rows(values_only=True)]
    if not rows:
        return [], pd.DataFrame(), None
    header = [_cell_text(v) for v in rows[0]]
    keep = [i for i, h in enumerate(header) if h]
    data = [[r[i] if i < len(r) else None for i in keep] for r in rows[1:]]
    data = [r for r in data if not all(is_missing(v) for v in r)]
    return [header[i] for i in keep], pd.DataFrame(data, columns=[header[i] for i in keep], dtype=object), None


def _find_header(headers: list[str], name: str) -> str | None:
    if name in headers:
        return name
    matches = [h for h in headers if normalize_name(h) == normalize_name(name)]
    return matches[0] if len(matches) == 1 else None


@dataclass
class SupportList:
    sheet: str
    column: str  # column name in the schema
    header: str | None  # actual header found
    values: list[str]
    headers: list[str]
    table: str | None
    problem: str | None = None


def _read_support(wb, scan: PackageScan, sheet: str, column: str) -> SupportList:
    if sheet not in wb.sheetnames:
        return SupportList(sheet, column, None, [], [], None, f"foglio '{sheet}' non trovato")
    headers, frame, table = _sheet_frame(wb, scan, sheet, prefer_table=sheet)
    header = _find_header(headers, column)
    if header is None:
        return SupportList(sheet, column, None, [], headers, table,
                           f"colonna '{column}' non trovata nel foglio '{sheet}' (colonne: {', '.join(headers)})")
    values = list(dict.fromkeys(v for v in (to_text(x) for x in frame[header]) if v))
    return SupportList(sheet, column, header, values, headers, table)


# ---------------------------------------------------------------------------
# Snapshot + doctor
# ---------------------------------------------------------------------------

@dataclass
class DBSnapshot:
    path: Path
    mtime_ns: int
    size: int
    metadata: TableInfo | None
    kpis: TableInfo | None
    templates: pd.DataFrame | None
    support: dict[tuple[str, str], SupportList]
    scan: PackageScan
    lock: str | None
    next_id: int
    kpi_names: list[str]
    issues: list[Issue]

    @property
    def blockers(self) -> list[Issue]:
        return errors(self.issues)

    @property
    def writable(self) -> bool:
        return not self.blockers

    def options(self, schema: Schema) -> dict[str, list[str]]:
        """Allowed values per schema column (fixed lists + support sheets)."""
        out = {}
        for f in schema.fields:
            if f.choices:
                out[f.column] = list(f.choices)
            elif f.uses_support_sheet:
                support = self.support.get((f.choices_sheet, f.choices_column))
                out[f.column] = list(support.values) if support else []
        return out

    def metadata_row(self, record_id: Any, id_column: str = "ID") -> dict[str, Any] | None:
        if self.metadata is None or id_column not in self.metadata.df.columns:
            return None
        key = to_text(record_id)
        for _, row in self.metadata.df.iterrows():
            if to_text(row[id_column]) == key:
                return {c: row[c] for c in self.metadata.df.columns}
        return None


def _numeric_ids(values) -> list[float]:
    out = []
    for v in values:
        try:
            number = to_number(v)
        except ConversionError:
            continue
        if number is not None:
            out.append(number)
    return out


def next_id(meta: TableInfo | None, kpi: TableInfo | None, id_column: str) -> int:
    ids = []
    for info in (meta, kpi):
        if info is not None and id_column in info.df.columns:
            ids += _numeric_ids(info.df[id_column])
    return int(max(ids)) + 1 if ids else 1


def kpi_columns_of(kpi: TableInfo, schema: Schema) -> list[str]:
    """KPI columns = TableKPI columns that are not metadata (MATLAB setdiff)."""
    return [h for h in kpi.headers if schema.match(h) is None]


def consistency_issues(meta: TableInfo, kpi: TableInfo, settings: Settings) -> list[Issue]:
    """MATLAB validateDBConsistency: one KPI row per metadata row, same IDs."""
    id_column = settings.id_column
    issues = []
    for info in (meta, kpi):
        if id_column not in info.headers:
            issues.append(Issue(ERROR, f"colonna '{id_column}' mancante nella tabella {info.name}"))
    if issues:
        return issues
    meta_ids = [to_text(v) for v in meta.df[id_column]]
    kpi_ids = [to_text(v) for v in kpi.df[id_column]]
    for info, ids in ((meta, meta_ids), (kpi, kpi_ids)):
        empty = [str(i + info.first_data_row) for i, v in enumerate(ids) if v is None]
        if empty:
            issues.append(Issue(ERROR, f"{info.name}: righe Excel senza ID: {', '.join(empty[:20])}"))
        seen, dup = set(), []
        for v in ids:
            if v is not None and v in seen and v not in dup:
                dup.append(v)
            seen.add(v)
        if dup:
            issues.append(Issue(ERROR, f"{info.name}: ID duplicati: {', '.join(dup[:20])}"))
    if len(meta_ids) != len(kpi_ids):
        issues.append(Issue(ERROR, f"{meta.name} ha {len(meta_ids)} righe, {kpi.name} ne ha {len(kpi_ids)}"))
    only_meta = [v for v in meta_ids if v is not None and v not in set(kpi_ids)]
    only_kpi = [v for v in kpi_ids if v is not None and v not in set(meta_ids)]
    if only_meta:
        issues.append(Issue(ERROR, f"ID senza riga KPI: {', '.join(only_meta[:20])}"))
    if only_kpi:
        issues.append(Issue(ERROR, f"ID KPI senza riga di metadati: {', '.join(only_kpi[:20])}"))
    if not only_meta and not only_kpi and len(meta_ids) == len(kpi_ids) and meta_ids != kpi_ids:
        wrong = [str(i + 1) for i, (a, b) in enumerate(zip(meta_ids, kpi_ids)) if a != b]
        issues.append(Issue(ERROR if settings.strict_row_order else WARNING,
                            f"stessi ID ma in ordine diverso (righe {', '.join(wrong[:20])})"))
    return issues


def _date_kind(value: Any) -> str:
    value = plain(value)
    if is_missing(value):
        return "vuote"
    if isinstance(value, (dt.date, dt.datetime)):
        return "data Excel" if MIN_YEAR <= value.year <= MAX_YEAR else "anomale (anno fuori 1900-2100)"
    if isinstance(value, str):
        try:
            dt.datetime.strptime(value.strip(), "%d/%m/%Y")
            return "testo gg/mm/aaaa"
        except ValueError:
            try:
                to_date(value)
                return "testo in altro formato"
            except ConversionError:
                return "testo non data"
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        try:
            to_date(value)
            return "numero (seriale Excel)"
        except ConversionError:
            return "anomale (anno fuori 1900-2100)"
    return "altro"


def date_report(snapshot: DBSnapshot, schema: Schema) -> pd.DataFrame:
    """How the date columns are stored today (MATLAB wrote text, serials and datenum)."""
    rows = []
    for info in (snapshot.metadata, snapshot.kpis):
        if info is None:
            continue
        for header in info.headers:
            f = schema.match(header)
            if f is None or f.type != "date":
                continue
            counts: dict[str, Any] = {"tabella": info.name, "colonna": header}
            for value in info.df[header]:
                kind = _date_kind(value)
                counts[kind] = counts.get(kind, 0) + 1
            rows.append(counts)
    frame = pd.DataFrame(rows).fillna(0)
    for column in frame.columns[2:]:
        frame[column] = frame[column].astype(int)
    return frame


def read_db(path: str | Path, schema: Schema, settings: Settings) -> DBSnapshot:
    """Read the DB (cached values, as Power BI sees them) and run all the checks."""
    path = Path(path)
    if not path.exists():
        raise DBError(f"file non trovato: {path}")
    if path.suffix.lower() not in (".xlsx", ".xlsm"):
        raise DBError(f"{path.name}: formato non supportato (serve .xlsx o .xlsm)")
    stat = path.stat()
    scan = scan_package(path)
    issues: list[Issue] = []
    lock = lock_status(path)
    if lock:
        issues.append(Issue(WARNING, f"{lock}: si può consultare ma non scrivere"))

    wb = open_workbook(path, DBError, read_only=True, data_only=True)  # streaming: cached values, as Power BI
    try:
        return _read_open_db(wb, path, stat, scan, lock, issues, schema, settings)
    finally:
        wb.close()


def _read_open_db(wb, path: Path, stat, scan: PackageScan, lock: str | None, issues: list[Issue],
                  schema: Schema, settings: Settings) -> DBSnapshot:
    meta = kpi = None
    for attr, table, sheet in (("meta", settings.metadata_table, settings.metadata_sheet),
                               ("kpi", settings.kpi_table, settings.kpi_sheet)):
        definition = scan.tables.get(table.lower())
        if definition is None:
            issues.append(Issue(ERROR, f"tabella Excel '{table}' non trovata nel file "
                                       f"(fogli: {', '.join(wb.sheetnames)})"))
            continue
        try:
            info = _table_info_ro(wb, definition)
        except DBError as exc:
            issues.append(Issue(ERROR, str(exc)))
            continue
        if info.sheet != sheet:
            issues.append(Issue(INFO, f"tabella {table} trovata nel foglio '{info.sheet}' invece di '{sheet}'"))
        if info.totals_rows:
            issues.append(Issue(ERROR, f"{table}: la riga dei totali non è supportata in scrittura"))
        if info.calculated_columns:
            issues.append(Issue(ERROR, f"{table}: colonne calcolate non supportate in scrittura: "
                                       f"{', '.join(info.calculated_columns)}"))
        if info.header_mismatch:
            issues.append(Issue(WARNING, f"{table}: intestazioni diverse dai nomi della tabella: "
                                         f"{', '.join(info.header_mismatch)}"))
        in_table = _formulas_inside(scan, info)
        if in_table:
            issues.append(Issue(ERROR, f"{table}: celle con formula ({', '.join(in_table[:10])}); "
                                       "salvando con openpyxl perderebbero il valore calcolato letto da Power BI"))
        if attr == "meta":
            meta = info
        else:
            kpi = info

    other_formulas = {s: c for s, c in scan.formulas.items()
                      if s not in {settings.metadata_sheet, settings.kpi_sheet}}
    if other_formulas:
        detail = ", ".join(f"{s} ({len(c)})" for s, c in other_formulas.items())
        issues.append(Issue(WARNING, f"formule in altri fogli: {detail}. Dopo una scrittura restano le "
                                     "formule ma non i valori calcolati: se Power BI legge questi fogli, "
                                     "apri e salva il file in Excel una volta"))
    for label in scan.parts:
        if label in _REWRITTEN:
            issues.append(Issue(WARNING, f"il file contiene {label}: openpyxl li conserva ma li riscrive, "
                                         "controllane l'aspetto dopo la prima scrittura"))
    if scan.parts.get("modello dati Power Pivot") or scan.parts.get("slicer o timeline") \
            or scan.parts.get("connessioni dati") or scan.power_query:
        issues.append(Issue(ERROR, "il file contiene modello dati, query, connessioni o slicer: "
                                   "openpyxl non li conserva, la scrittura è disabilitata"))
    for label, sheets in scan.x14.items():
        issues.append(Issue(WARNING, f"{label} nei fogli {', '.join(sheets)}: potrebbero andare persi"))

    if meta is not None and kpi is not None:
        issues += consistency_issues(meta, kpi, settings)
        for column in schema.kpi_table_columns():
            if column not in kpi.headers:
                issues.append(Issue(ERROR, f"{kpi.name}: manca la colonna obbligatoria '{column}'"))
    if meta is not None:
        missing = [c for c in schema.columns if _find_header(meta.headers, c) is None]
        if missing:
            issues.append(Issue(WARNING, f"{meta.name}: colonne dello schema assenti (non verranno "
                                         f"scritte): {', '.join(missing)}"))
        fuzzy = [f"{c}→{_find_header(meta.headers, c)}" for c in schema.columns
                 if c not in meta.headers and _find_header(meta.headers, c)]
        if fuzzy:
            issues.append(Issue(INFO, f"{meta.name}: colonne abbinate ignorando maiuscole/spazi: {', '.join(fuzzy)}"))
        extra = [h for h in meta.headers if schema.match(h) is None]
        if extra:
            issues.append(Issue(INFO, f"{meta.name}: colonne non presenti nello schema: {', '.join(extra)}"))

    support = {}
    for sheet, column in schema.support_sheets():
        support[(sheet, column)] = _read_support(wb, scan, sheet, column)
        if support[(sheet, column)].problem:
            issues.append(Issue(WARNING, f"lista di supporto: {support[(sheet, column)].problem}"))
    templates = None
    if settings.templates_sheet in wb.sheetnames:
        _, templates, _ = _sheet_frame(wb, scan, settings.templates_sheet)
        if _find_header(list(templates.columns), settings.template_name_column) is None:
            issues.append(Issue(WARNING, f"{settings.templates_sheet}: colonna "
                                         f"'{settings.template_name_column}' non trovata"))
    else:
        issues.append(Issue(WARNING, f"foglio '{settings.templates_sheet}' non trovato: template non disponibili"))

    snapshot = DBSnapshot(
        path=path, mtime_ns=stat.st_mtime_ns, size=stat.st_size, metadata=meta, kpis=kpi,
        templates=templates, support=support, scan=scan, lock=lock,
        next_id=next_id(meta, kpi, settings.id_column),
        kpi_names=kpi_columns_of(kpi, schema) if kpi is not None else [],
        issues=issues,
    )
    report = date_report(snapshot, schema)
    if not report.empty:
        anomalies = [f"{r['tabella']}.{r['colonna']} ({int(r.get('anomale (anno fuori 1900-2100)', 0))})"
                     for _, r in report.iterrows() if r.get("anomale (anno fuori 1900-2100)", 0)]
        if anomalies:
            issues.append(Issue(WARNING, "date anomale (probabile bug datenum dell'app MATLAB): "
                                         + ", ".join(anomalies)))
    return snapshot


def _inside(coord: str, min_col: int, max_col: int, min_row: int, max_row: int) -> bool:
    match = re.match(r"([A-Z]+)(\d+)", coord)
    if not match:
        return False
    col, row = column_index_from_string(match.group(1)), int(match.group(2))
    return min_col <= col <= max_col and min_row <= row <= max_row


def _formulas_inside(scan: PackageScan, info: TableInfo) -> list[str]:
    return [c for c in scan.formulas.get(info.sheet, [])
            if _inside(c, info.min_col, info.max_col, info.first_data_row, info.ref_last_row)]


# ---------------------------------------------------------------------------
# Writing
# ---------------------------------------------------------------------------

@dataclass
class WriteResult:
    action: str
    path: Path
    backup: Path | None = None
    ids: list[int] = field(default_factory=list)
    new_kpi_columns: list[str] = field(default_factory=list)
    skipped_columns: list[str] = field(default_factory=list)
    messages: list[str] = field(default_factory=list)
    elapsed_s: float = 0.0


@dataclass
class _Expect:
    tables: dict[str, tuple[str, list[str]]] = field(default_factory=dict)  # name -> (ref, headers)
    cells: list[tuple[str, int, int, Any]] = field(default_factory=list)  # sheet, row, col, value
    header_cells: set[str] = field(default_factory=set)  # tables whose header cells must match

    def table(self, table, info: TableInfo) -> None:
        self.tables[table.displayName] = (table.ref, list(info.headers))
        if not info.header_mismatch:
            self.header_cells.add(table.displayName)


def _clean_text(text: str | None) -> str | None:
    return ILLEGAL_CHARACTERS_RE.sub("", text) if text else text


def excel_value(value: Any, field_type: str | None, settings: Settings) -> tuple[Any, str | None]:
    """Python value -> (cell value, number format) following settings.date_storage."""
    value = plain(value)
    if is_missing(value):
        return None, None
    if field_type == "date" or (field_type is None and isinstance(value, (dt.date, dt.datetime))):
        try:
            day, _ = to_date(value)
        except ConversionError:
            return _clean_text(to_text(value)), None
        if settings.date_storage == "excel":
            return dt.datetime(day.year, day.month, day.day), "dd/mm/yyyy"
        return day.strftime(settings.date_text_format), None
    try:
        if field_type in ("id", "int"):
            return to_int(value), None
        if field_type == "number":
            return to_number(value), None
    except ConversionError:
        return _clean_text(to_text(value)), None
    if field_type in ("text", "choice", "multichoice") or isinstance(value, str):
        return _clean_text(to_text(value)), None
    return value, None


def _write_cell(cell, value: Any, number_format: str | None) -> None:
    cell.value = value
    if isinstance(value, str) and value.startswith("="):
        cell.data_type = "s"  # text, never a formula
    if number_format:
        cell.number_format = number_format


def _intersects(a: tuple[int, int, int, int], b: tuple[int, int, int, int]) -> bool:
    return not (a[2] < b[0] or b[2] < a[0] or a[3] < b[1] or b[3] < a[1])


def _ensure_free(ws, table, rows: range, cols: range) -> None:
    if not rows or not cols:
        return
    area = (cols.start, rows.start, cols.stop - 1, rows.stop - 1)
    for other in ws.tables.values():
        if other is not table and _intersects(area, range_boundaries(other.ref)):
            raise DBError(f"non c'è spazio: la tabella '{other.displayName}' è adiacente a "
                          f"'{table.displayName}' nel foglio {ws.title}")
    for merged in ws.merged_cells.ranges:
        if _intersects(area, (merged.min_col, merged.min_row, merged.max_col, merged.max_row)):
            raise DBError(f"celle unite ({merged.coord}) nell'area dove va estesa la tabella "
                          f"'{table.displayName}'")
    for r in rows:
        for c in cols:
            if not is_missing(ws.cell(r, c).value):
                raise DBError(f"la cella {get_column_letter(c)}{r} del foglio {ws.title} non è vuota: "
                              f"serve spazio per estendere la tabella '{table.displayName}'")


def _check_writable(info: TableInfo, scan: PackageScan) -> None:
    if info.totals_rows:
        raise DBError(f"{info.name}: riga dei totali non supportata")
    if info.calculated_columns:
        raise DBError(f"{info.name}: colonne calcolate non supportate ({', '.join(info.calculated_columns)})")
    formulas = _formulas_inside(scan, info)
    if formulas:
        raise DBError(f"{info.name}: celle con formula ({', '.join(formulas[:10])}): salvando si perderebbe "
                      "il valore calcolato letto da Power BI")


def _set_ref(table, info_min_col: int, header_row: int, max_col: int, last_row: int) -> str:
    table.ref = f"{get_column_letter(info_min_col)}{header_row}:{get_column_letter(max_col)}{last_row}"
    if table.autoFilter is not None:
        table.autoFilter.ref = table.ref
    return table.ref


def _add_column(ws, table, info: TableInfo, name: str) -> TableInfo:
    col = info.max_col + 1
    _ensure_free(ws, table, range(info.header_row, info.ref_last_row + 1), range(col, col + 1))
    for r in range(info.header_row, info.ref_last_row + 1):
        source = ws.cell(r, col - 1)
        if source.has_style:
            ws.cell(r, col)._style = copy.copy(source._style)
    header = ws.cell(info.header_row, col)
    header.value = name
    header.data_type = "s"
    previous = ws.column_dimensions[get_column_letter(col - 1)]
    if previous.width:
        ws.column_dimensions[get_column_letter(col)].width = previous.width
    next_id_value = max((tc.id for tc in table.tableColumns), default=0) + 1
    table.tableColumns.append(TableColumn(id=next_id_value, name=name))
    _set_ref(table, info.min_col, info.header_row, col, info.ref_last_row)
    frame = info.df.copy()
    frame[name] = pd.Series([None] * len(frame), dtype=object)
    return TableInfo(info.name, info.sheet, [*info.headers, name], frame, info.min_col, col, info.header_row,
                     info.ref_last_row, info.last_used_row, info.totals_rows, info.calculated_columns,
                     info.header_mismatch)


def _append_rows(ws, table, info: TableInfo, rows: list[dict[str, Any]],
                 types: dict[str, str | None], settings: Settings) -> tuple[int, int]:
    """Write rows after the last used row, extending the table. Returns (first, last) Excel rows."""
    if not rows:
        return info.last_used_row + 1, info.last_used_row
    start = info.last_used_row + 1
    end = start + len(rows) - 1
    if end > info.ref_last_row:
        _ensure_free(ws, table, range(max(start, info.ref_last_row + 1), end + 1),
                     range(info.min_col, info.max_col + 1))
    style_row = info.last_used_row if info.last_used_row > info.header_row else None
    for k, values in enumerate(rows):
        r = start + k
        for j, header in enumerate(info.headers):
            c = info.min_col + j
            cell = ws.cell(r, c)
            if style_row is not None:
                source = ws.cell(style_row, c)
                if source.has_style:
                    cell._style = copy.copy(source._style)
            value, number_format = excel_value(values.get(header), types.get(header), settings)
            _write_cell(cell, value, number_format)
    if end > info.ref_last_row:
        _set_ref(table, info.min_col, info.header_row, info.max_col, end)
    return start, end


def _numeric_cells(path: Path, sheet: str, column: int) -> dict[int, str]:
    """Values of the numeric cells of one column, read with a single regex pass on the sheet XML."""
    letter = get_column_letter(column).encode()
    pattern = re.compile(rb'<c\s[^>]*?\br="' + letter + rb'(\d+)"[^>]*>(?:<f[^>]*>[^<]*</f>|<f[^>]*/>)?<v>([^<]*)</v>')
    with zipfile.ZipFile(path) as archive:
        part = _sheet_parts(archive)[sheet]
        xml = archive.read(part)
    return {int(row): value.decode() for row, value in pattern.findall(xml)}


def _verify(tmp: Path, expect: _Expect, before: PackageScan) -> list[str]:
    """Check the saved file without loading it again: package parts, table
    definitions, header cells and the IDs just written."""
    after = scan_package(tmp)
    blocking, soft = lost_features(before, after)
    if blocking:
        raise DBError("il salvataggio avrebbe perso: " + ", ".join(blocking))
    for name, (ref, headers) in expect.tables.items():
        table = after.tables.get(name.lower())
        if table is None or table.ref != ref or table.columns != headers:
            found = table.ref if table else "tabella assente"
            raise DBError(f"verifica fallita sulla tabella {name}: {found} invece di {ref}")
        min_col, _, max_col, _ = range_boundaries(ref)
        if len(set(table.column_ids)) != len(headers) or len(headers) != max_col - min_col + 1:
            raise DBError(f"verifica fallita sulle colonne della tabella {name}")
    if expect.header_cells:
        wb = load_workbook(tmp, read_only=True)
        try:
            for name in expect.header_cells:
                table = after.tables[name.lower()]
                min_col, header_row, max_col, _ = range_boundaries(table.ref)
                row = _range_values(wb[table.sheet], header_row, header_row, min_col, max_col)
                if [_cell_text(v) for v in (row[0] if row else [])] != table.columns:
                    raise DBError(f"verifica fallita sulle intestazioni della tabella {name}")
        finally:
            wb.close()
    columns: dict[tuple[str, int], dict[int, str]] = {}
    for sheet, r, c, value in expect.cells:
        if (sheet, c) not in columns:
            columns[(sheet, c)] = _numeric_cells(tmp, sheet, c)
        found = columns[(sheet, c)].get(r)
        if found is None or to_text(to_number(found)) != to_text(value):
            raise DBError(f"verifica fallita: {sheet}!{get_column_letter(c)}{r} = {found!r}, atteso {value!r}")
    return [f"dopo il salvataggio non sono più presenti: {s}" for s in soft]


def _transaction(path: str | Path, settings: Settings, action: str,
                 mutate: Callable[[Any, PackageScan], tuple[WriteResult, _Expect]]) -> WriteResult:
    started = time.perf_counter()
    path = Path(path)
    if path.suffix.lower() not in (".xlsx", ".xlsm"):
        raise DBError(f"{path.name}: formato non supportato (serve .xlsx o .xlsm)")
    lock = lock_status(path)
    if lock and settings.refuse_if_open:
        raise DBLockedError(f"{lock}. Chiudi il file e riprova.")
    before_stat = path.stat()
    before = scan_package(path)
    backup = make_backup(path, settings)
    wb = open_workbook(path, DBError, keep_vba=path.suffix.lower() == ".xlsm", rich_text=True)
    result, expect = mutate(wb, before)
    tmp = path.with_name(f".kpimeta-tmp-{os.getpid()}-{int(time.time() * 1000)}{path.suffix}")
    try:
        wb.save(tmp)
        result.messages += _verify(tmp, expect, before)
        now = path.stat()
        if (now.st_mtime_ns, now.st_size) != (before_stat.st_mtime_ns, before_stat.st_size):
            raise DBChangedError("il DB è stato modificato da qualcun altro durante la scrittura: "
                                 "nessuna modifica applicata, riprova")
        os.replace(tmp, path)
    except PermissionError as exc:
        raise DBLockedError(f"impossibile sostituire {path.name}: file aperto o in sola lettura ({exc})") from exc
    finally:
        if tmp.exists():
            try:
                tmp.unlink()
            except OSError:
                pass
    result.action = action
    result.path = path
    result.backup = backup
    result.elapsed_s = time.perf_counter() - started
    return result


def _resolve_columns(columns: list[str], headers: list[str]) -> tuple[dict[str, str], list[str]]:
    mapping, missing = {}, []
    for column in columns:
        header = _find_header(headers, column)
        if header is None:
            missing.append(column)
        else:
            mapping[column] = header
    return mapping, missing


def _kpi_types(info: TableInfo, schema: Schema) -> dict[str, str | None]:
    return {h: (schema.match(h).type if schema.match(h) else "number") for h in info.headers}


def _meta_types(info: TableInfo, schema: Schema) -> dict[str, str | None]:
    return {h: (schema.match(h).type if schema.match(h) else None) for h in info.headers}


def build_rows(batch: Batch, schema: Schema, settings: Settings, decisions: dict[str, KpiDecision],
               db_kpi_columns: list[str], start_id: int) -> tuple[list[dict], list[dict], list[str]]:
    """Metadata rows, KPI rows (same ID) and the KPI columns to create."""
    id_column = schema.id_field().column
    copied = [c for c in schema.kpi_table_columns() if c != id_column]
    known, unknown, _ = kpi_resolution(batch, db_kpi_columns, settings)
    targets = {name: name for name in known}
    new_columns = []
    for name in unknown:
        decision = decisions.get(name)
        if decision is None or decision.action == "discard":
            continue
        if decision.action == "add":
            targets[name] = name
            new_columns.append(name)
        elif decision.action == "map" and decision.target:
            targets[name] = decision.target
    meta_rows, kpi_rows = [], []
    for r in range(batch.n_rows):
        record_id = start_id + r
        meta = {c: batch.df.at[r, c] for c in batch.meta_columns}
        meta[id_column] = record_id
        kpi = {id_column: record_id}
        for column in copied:
            kpi[column] = batch.df.at[r, column] if column in batch.df.columns else None
        for name, target in targets.items():
            kpi[target] = batch.df.at[r, name]
        meta_rows.append(meta)
        kpi_rows.append(kpi)
    return meta_rows, kpi_rows, new_columns


def preview_write(snapshot: DBSnapshot, batch: Batch, schema: Schema, settings: Settings,
                  decisions: dict[str, KpiDecision]) -> tuple[pd.DataFrame, pd.DataFrame, list[str], list[str]]:
    """What would be appended to TableMet and TableKPI (values as they will be stored)."""
    meta_rows, kpi_rows, new_columns = build_rows(batch, schema, settings, decisions,
                                                  snapshot.kpi_names, snapshot.next_id)
    meta_headers = snapshot.metadata.headers
    kpi_headers = snapshot.kpis.headers + [c for c in new_columns if c not in snapshot.kpis.headers]
    mapping, missing = _resolve_columns(batch.meta_columns + [settings.id_column], meta_headers)
    meta_types = _meta_types(snapshot.metadata, schema)
    kpi_types = {h: (schema.match(h).type if schema.match(h) else "number") for h in kpi_headers}

    def render(rows, headers, types, remap):
        out = []
        for row in rows:
            values = {remap.get(k, k): v for k, v in row.items()}
            out.append({h: excel_value(values.get(h), types.get(h), settings)[0] for h in headers})
        return pd.DataFrame(out, columns=headers, dtype=object)

    kpi_map = {c: (_find_header(kpi_headers, c) or c) for c in kpi_rows[0]} if kpi_rows else {}
    skipped = [c for c in missing if any(not is_missing(batch.df.at[r, c]) for r in range(batch.n_rows))]
    return (render(meta_rows, meta_headers, meta_types, mapping),
            render(kpi_rows, kpi_headers, kpi_types, kpi_map), new_columns, skipped)


def write_batch(path: str | Path, batch: Batch, schema: Schema, settings: Settings,
                decisions: dict[str, KpiDecision] | None = None) -> WriteResult:
    """Append the batch: one row in TableMet and one in TableKPI per batch row, same ID."""
    decisions = decisions or {}
    id_column = schema.id_field().column
    if batch.n_rows == 0:
        raise DBError("nessuna riga da scrivere")
    invalid = errors(validate_batch(batch, schema, settings))  # types, limits, fixed lists
    if invalid:
        raise DBError("scrittura annullata, valori non validi:\n- "
                      + "\n- ".join(i.label() for i in invalid[:20]))

    def mutate(wb, scan):
        ws_m, t_m = _find_table(wb, settings.metadata_table, settings.metadata_sheet)
        ws_k, t_k = _find_table(wb, settings.kpi_table, settings.kpi_sheet)
        meta, kpi = _table_info(ws_m, t_m), _table_info(ws_k, t_k)
        _check_writable(meta, scan)
        _check_writable(kpi, scan)
        problems = errors(consistency_issues(meta, kpi, settings))
        for column in schema.kpi_table_columns():
            if column not in kpi.headers:
                problems.append(Issue(ERROR, f"{kpi.name}: manca la colonna '{column}'"))
        problems += duplicates_vs_db(batch, schema, meta.df)
        db_kpis = kpi_columns_of(kpi, schema)
        problems += errors(check_decisions(batch, db_kpis, decisions, settings))
        if problems:
            raise DBError("scrittura annullata:\n- " + "\n- ".join(p.label() for p in problems[:20]))

        start_id = next_id(meta, kpi, id_column)
        meta_rows, kpi_rows, new_columns = build_rows(batch, schema, settings, decisions, db_kpis, start_id)
        for name in new_columns:
            if name not in kpi.headers:
                kpi = _add_column(ws_k, t_k, kpi, name)
        mapping, missing = _resolve_columns(batch.meta_columns + [id_column], meta.headers)
        kpi_map, _ = _resolve_columns(list(kpi_rows[0]) if kpi_rows else [], kpi.headers)
        rows_m = [{mapping[k]: v for k, v in row.items() if k in mapping} for row in meta_rows]
        rows_k = [{kpi_map[k]: v for k, v in row.items() if k in kpi_map} for row in kpi_rows]
        first_m, last_m = _append_rows(ws_m, t_m, meta, rows_m, _meta_types(meta, schema), settings)
        first_k, last_k = _append_rows(ws_k, t_k, kpi, rows_k, _kpi_types(kpi, schema), settings)

        ids = [start_id + r for r in range(batch.n_rows)]
        expect = _Expect()
        expect.table(t_m, meta)
        expect.table(t_k, kpi)
        id_col_m = meta.column_of(mapping[id_column])
        id_col_k = kpi.column_of(kpi_map[id_column])
        for k, record_id in enumerate(ids):
            expect.cells.append((ws_m.title, first_m + k, id_col_m, record_id))
            expect.cells.append((ws_k.title, first_k + k, id_col_k, record_id))
        skipped = [c for c in missing if any(not is_missing(batch.df.at[r, c]) for r in range(batch.n_rows))]
        result = WriteResult("append", Path(path), ids=ids, new_kpi_columns=new_columns, skipped_columns=skipped)
        if skipped:
            result.messages.append("valori non scritti perché la colonna non esiste in "
                                   f"{meta.name}: {', '.join(skipped)}")
        return result, expect

    return _transaction(path, settings, "append", mutate)


def overwrite_metadata(path: str | Path, record_id: Any, values: dict[str, Any], schema: Schema,
                       settings: Settings) -> WriteResult:
    """Overwrite the metadata row with this ID; IDName/Date are also updated in the KPI row."""
    id_column = schema.id_field().column

    def mutate(wb, scan):
        ws_m, t_m = _find_table(wb, settings.metadata_table, settings.metadata_sheet)
        ws_k, t_k = _find_table(wb, settings.kpi_table, settings.kpi_sheet)
        meta, kpi = _table_info(ws_m, t_m), _table_info(ws_k, t_k)
        _check_writable(meta, scan)
        key = to_text(record_id)
        rows = [i for i, v in enumerate(meta.df[id_column]) if to_text(v) == key] \
            if id_column in meta.df.columns else []
        if len(rows) != 1:
            raise DBError(f"ID {key}: {len(rows)} righe trovate in {meta.name} (ne serve esattamente una)")
        excel_row = meta.first_data_row + rows[0]
        mapping, missing = _resolve_columns([c for c in values if c != id_column], meta.headers)
        types = _meta_types(meta, schema)
        changed = []
        expect = _Expect()
        expect.table(t_m, meta)
        for column, header in mapping.items():
            cell = ws_m.cell(excel_row, meta.column_of(header))
            new_value, number_format = excel_value(values[column], types.get(header), settings)
            if to_text(cell.value) != to_text(new_value):
                changed.append(header)
                _write_cell(cell, new_value, number_format)
        messages = []
        synced = []
        kpi_rows = [i for i, v in enumerate(kpi.df[id_column]) if to_text(v) == key] \
            if id_column in kpi.df.columns else []
        if kpi_rows:
            _check_writable(kpi, scan)
            for column in schema.kpi_table_columns():
                if column == id_column or column not in values or column not in kpi.headers:
                    continue
                cell = ws_k.cell(kpi.first_data_row + kpi_rows[0], kpi.column_of(column))
                new_value, number_format = excel_value(values[column], schema.type_of(column), settings)
                if to_text(cell.value) != to_text(new_value):
                    _write_cell(cell, new_value, number_format)
                    synced.append(column)
        else:
            messages.append(f"nessuna riga in {kpi.name} con ID {key}")
        if synced:
            messages.append(f"aggiornati anche in {kpi.name}: {', '.join(synced)}")
        expect.cells.append((ws_m.title, excel_row, meta.column_of(id_column), record_id))
        summary = f"colonne modificate: {', '.join(changed) or 'nessuna'}"
        result = WriteResult("overwrite", Path(path), ids=[int(to_number(record_id))],
                             skipped_columns=missing, messages=[summary] + messages)
        return result, expect

    return _transaction(path, settings, "overwrite", mutate)


def add_support_value(path: str | Path, sheet: str, values: dict[str, Any], settings: Settings) -> WriteResult:
    """Append a row to a support sheet (Cluster, Brand, ...). A numeric first column
    that is not given is filled with max + 1 (as in the MATLAB add-row dialog)."""

    def mutate(wb, scan):
        if sheet not in wb.sheetnames:
            raise DBError(f"foglio '{sheet}' non trovato")
        ws = wb[sheet]
        tables = list(ws.tables.values())
        table = tables[0] if len(tables) == 1 else next(
            (t for t in tables if t.displayName.lower() == sheet.lower()), None)
        expect = _Expect()
        if table is not None:
            info = _table_info(ws, table)
            _check_writable(info, scan)
            row = {(_find_header(info.headers, k) or k): v for k, v in values.items()}
            first = info.headers[0]
            if is_missing(row.get(first)):
                numbers = _numeric_ids(info.df[first])
                if numbers or info.df.empty:
                    row[first] = int(max(numbers)) + 1 if numbers else 1
            start, _ = _append_rows(ws, table, info, [row], {}, settings)
            expect.table(table, info)
            target_row = start
        else:
            headers = [_cell_text(c.value) for c in ws[1]]
            target_row = ws.max_row + 1
            for c, header in enumerate(headers, start=1):
                value = values.get(header)
                if c == 1 and is_missing(value):
                    numbers = _numeric_ids(ws.cell(r, 1).value for r in range(2, ws.max_row + 1))
                    value = int(max(numbers)) + 1 if numbers else None
                _write_cell(ws.cell(target_row, c), *excel_value(value, None, settings))
        return WriteResult("support", Path(path), messages=[f"aggiunta riga {target_row} al foglio {sheet}"]), expect

    return _transaction(path, settings, "support", mutate)
