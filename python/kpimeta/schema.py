"""Metadata schema: the single source of truth for the metadata columns.

Derived from GeneralMetadataMap / ICMetadataMap and from the App Designer
components of IC_KPIs_App (type, limits, support sheet of each "DB" button).
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from .config import ConfigError, read_csv_records

FIELD_TYPES = ("id", "text", "int", "number", "date", "choice", "multichoice")
LOLA_POLICIES = ("", "keep", "keep_if_valid")
COLUMNS = ("column", "label", "tab", "section", "type", "choices_sheet", "choices_column",
           "choices", "min", "max", "default", "required", "lola_policy", "in_kpi_table", "notes")


@dataclass(frozen=True)
class Field:
    column: str
    label: str
    tab: str
    section: str
    type: str
    choices_sheet: str = ""
    choices_column: str = ""
    choices: tuple[str, ...] = ()
    min: float | None = None
    max: float | None = None
    default: str = ""
    required: bool = False
    lola_policy: str = ""  # keep = always from LoLa; keep_if_valid = LoLa wins if valid
    in_kpi_table: bool = False  # copied into every row of the KPI table (ID, IDName, Date)
    notes: str = ""

    @property
    def uses_support_sheet(self) -> bool:
        return self.type == "choice" and bool(self.choices_sheet)


class Schema:
    def __init__(self, fields: list[Field], source: str = ""):
        self.fields = list(fields)
        self.source = source
        self._by_column = {f.column: f for f in self.fields}
        self._by_lower = {f.column.lower(): f for f in self.fields}

    @property
    def columns(self) -> list[str]:
        return [f.column for f in self.fields]

    def __contains__(self, column: str) -> bool:
        return column in self._by_column

    def get(self, column: str) -> Field | None:
        return self._by_column.get(column)

    def match(self, column: str) -> Field | None:
        """Exact match first, then case-insensitive."""
        return self._by_column.get(column) or self._by_lower.get(str(column).strip().lower())

    def type_of(self, column: str) -> str | None:
        field = self.get(column)
        return field.type if field else None

    def tabs(self) -> list[str]:
        return list(dict.fromkeys(f.tab for f in self.fields))

    def sections(self, tab: str) -> list[str]:
        return list(dict.fromkeys(f.section for f in self.fields if f.tab == tab))

    def fields_in(self, tab: str, section: str) -> list[Field]:
        return [f for f in self.fields if f.tab == tab and f.section == section]

    def id_field(self) -> Field:
        ids = [f for f in self.fields if f.type == "id"]
        if len(ids) != 1:
            raise ConfigError("lo schema deve avere esattamente un campo di tipo 'id'")
        return ids[0]

    def kpi_table_columns(self) -> list[str]:
        return [f.column for f in self.fields if f.in_kpi_table]

    def support_sheets(self) -> list[tuple[str, str]]:
        return list(dict.fromkeys((f.choices_sheet, f.choices_column)
                                  for f in self.fields if f.uses_support_sheet))


def _number(text: str, where: str) -> float | None:
    if not text:
        return None
    try:
        return float(text)
    except ValueError as exc:
        raise ConfigError(f"{where}: '{text}' non è un numero") from exc


def _flag(text: str, where: str) -> bool:
    value = text.strip().lower()
    if value in ("", "no", "false", "0"):
        return False
    if value in ("yes", "si", "sì", "true", "1", "x"):
        return True
    raise ConfigError(f"{where}: usa yes/no (trovato '{text}')")


def load_schema(path: str | Path) -> Schema:
    path = Path(path)
    records = read_csv_records(path, required=list(COLUMNS[:5]), optional=COLUMNS[5:])
    fields = []
    seen: set[str] = set()
    for number, rec in enumerate(records, start=2):
        where = f"{path.name}, campo '{rec['column']}' (riga dati {number})"
        column = rec["column"]
        if not column:
            raise ConfigError(f"{path.name}, riga dati {number}: colonna 'column' vuota")
        if column.lower() in seen:
            raise ConfigError(f"{where}: colonna duplicata")
        seen.add(column.lower())
        ftype = rec["type"].lower()
        if ftype not in FIELD_TYPES:
            raise ConfigError(f"{where}: tipo '{rec['type']}' non valido ({', '.join(FIELD_TYPES)})")
        choices = tuple(c.strip() for c in rec["choices"].split("|") if c.strip())
        if ftype == "choice" and not choices and not (rec["choices_sheet"] and rec["choices_column"]):
            raise ConfigError(f"{where}: un campo 'choice' richiede 'choices' oppure foglio e colonna")
        policy = rec["lola_policy"].lower()
        if policy not in LOLA_POLICIES:
            raise ConfigError(f"{where}: lola_policy '{rec['lola_policy']}' non valida")
        fields.append(Field(
            column=column,
            label=rec["label"] or column,
            tab=rec["tab"] or "General",
            section=rec["section"] or "Altro",
            type=ftype,
            choices_sheet=rec["choices_sheet"],
            choices_column=rec["choices_column"],
            choices=choices,
            min=_number(rec["min"], where),
            max=_number(rec["max"], where),
            default=rec["default"],
            required=_flag(rec["required"], where),
            lola_policy=policy,
            in_kpi_table=_flag(rec["in_kpi_table"], where),
            notes=rec["notes"],
        ))
    schema = Schema(fields, source=str(path))
    schema.id_field()
    return schema
