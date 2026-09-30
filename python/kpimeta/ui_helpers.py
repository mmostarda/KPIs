"""UI logic that does not depend on Streamlit (unit-testable): form values,
review grid <-> batch conversion, exports."""

from __future__ import annotations

import datetime as dt
import io
from dataclasses import dataclass
from typing import Any

import pandas as pd

from .batch import INFERENCE_METHODS, Batch, export_batch_excel, set_value
from .issues import ERROR, INFO, WARNING, Issue
from .schema import Field, Schema
from .values import ConversionError, canonical, is_missing, split_multi, to_date, to_int, to_number, to_text

PROBLEMS_COL = "⚠ Problemi"
SOURCE_COL = "_source_file"
SEVERITY_LABEL = {ERROR: "Errore", WARNING: "Avviso", INFO: "Info"}


# --- form -------------------------------------------------------------------

def widget_key(column: str) -> str:
    return f"f::{column}"


def widget_value(field: Field, value: Any) -> Any:
    """Python value -> value accepted by the Streamlit widget of this field."""
    try:
        if field.type in ("text", "choice"):
            return to_text(value) or ""
        if field.type == "multichoice":
            return split_multi(value)
        if field.type in ("int", "id"):
            return to_int(value)
        if field.type == "number":
            return to_number(value)
        if field.type == "date":
            return to_date(value)[0]
    except ConversionError:
        pass
    return "" if field.type in ("text", "choice") else ([] if field.type == "multichoice" else None)


def field_value(field: Field, value: Any) -> Any:
    """Widget value -> Python value (None when empty)."""
    if field.type in ("text", "choice"):
        return (value or "").strip() or None
    if field.type == "multichoice":
        return ";".join(value) if value else None
    return value


def form_defaults(schema: Schema, with_defaults: bool) -> dict[str, Any]:
    """Initial widget values; schema defaults (ABS/ESC/TCS = ON) only when asked."""
    out = {}
    for f in schema.fields:
        raw = f.default if with_defaults and f.default else None
        out[widget_key(f.column)] = widget_value(f, raw)
    return out


def values_to_widgets(schema: Schema, values: dict[str, Any]) -> dict[str, Any]:
    """Values coming from a template / DB row / Excel row -> widget keys."""
    out = {}
    for column, value in values.items():
        f = schema.match(str(column))
        if f is not None and f.type != "id" and not is_missing(value):
            out[widget_key(f.column)] = widget_value(f, value)
    return out


# --- review grid ---------------------------------------------------------------

def _kind(batch: Batch, schema: Schema, column: str) -> str | None:
    if column in batch.kpi_columns:
        return "number"
    return schema.type_of(column)


def display_frame(batch: Batch, schema: Schema, problems: dict[int, str], rows: list[int] | None = None,
                  id_column: str = "ID") -> pd.DataFrame:
    """Batch -> DataFrame for st.data_editor, with dtypes that do not depend on the
    values (so that editing a cell never changes the editor identity)."""
    rows = list(range(batch.n_rows)) if rows is None else rows
    data: dict[str, pd.Series] = {}
    index = pd.Index([r + 1 for r in rows], name="riga")  # 1-based, like the problem list
    data[PROBLEMS_COL] = pd.Series([problems.get(r, "") for r in rows], index=index, dtype="string")
    columns = batch.review_columns() + ([SOURCE_COL] if SOURCE_COL in batch.df.columns else [])
    for column in columns:
        data[column] = _display_series([batch.df.at[r, column] for r in rows], _kind(batch, schema, column), index)
    return pd.DataFrame(data, index=index)


def _display_series(values: list[Any], kind: str | None, index: pd.Index) -> pd.Series:
    if kind in ("int", "id"):
        return pd.Series([to_int_safe(v) for v in values], index=index, dtype="Int64")
    if kind == "number":
        return pd.Series([to_float_safe(v) for v in values], index=index, dtype="float64")
    if kind == "date":
        return pd.Series(pd.to_datetime([v if isinstance(v, dt.date) else None for v in values]), index=index)
    return pd.Series([to_text(v) for v in values], index=index, dtype="string")


def to_int_safe(value: Any) -> int | None:
    try:
        return to_int(value)
    except ConversionError:
        return None


def to_float_safe(value: Any) -> float | None:
    try:
        number = to_number(value)
    except ConversionError:
        return None
    return float("nan") if number is None else number


def apply_edits(batch: Batch, schema: Schema, view_rows: list[int], edited_rows: dict,
                id_column: str = "ID") -> Batch:
    """Apply st.data_editor state ({position: {column: value}}) to a copy of the batch."""
    if not edited_rows:
        return batch
    result = batch.copy()
    for position, changes in edited_rows.items():
        row = view_rows[int(position)]
        for column, value in changes.items():
            if column in (PROBLEMS_COL, SOURCE_COL, id_column) or column not in result.df.columns:
                continue
            _set_edit(result, schema, row, column, value)
    return result


def _set_edit(batch: Batch, schema: Schema, row: int, column: str, value: Any) -> None:
    kind = _kind(batch, schema, column)
    if isinstance(value, str) and kind == "date" and "T" in value:
        value = value.split("T")[0]
    set_value(batch, row, column, value, kind, "edit")


# --- values inferred automatically (file name, per-row template, ...) ----------------------

FILE_COL, ROWS_COL, FROM_COL, MIXED_COL = "File", "Righe", "Dedotti da", "Valori diversi"
INFERRED_INFO_COLUMNS = (FILE_COL, ROWS_COL, FROM_COL, MIXED_COL)


@dataclass
class InferredGroup:
    """The rows of one file (its repetitions): inference works on the file name."""
    name: str
    rows: list[int]
    methods: list[str]  # inference methods that filled something in these rows


def inferred_columns(batch: Batch) -> list[str]:
    """Columns an inference method can fill, in review order."""
    targets = {c for columns in batch.inference_columns.values() for c in columns}
    return [c for c in batch.review_columns() if c in targets]


def inferred_groups(batch: Batch) -> list[InferredGroup]:
    """One group per file name (every file, also those whose name was not recognized)."""
    if not batch.inference_columns:
        return []
    by_row: dict[int, set[str]] = {}
    for (row, _), method in batch.inferred.items():
        by_row.setdefault(row, set()).add(method)
    has_file = "FileName" in batch.df.columns
    groups: dict[str, InferredGroup] = {}
    for r in range(batch.n_rows):
        name = (to_text(batch.df.at[r, "FileName"]) if has_file else None) or f"riga {r + 1}"
        group = groups.setdefault(name, InferredGroup(name, [], []))
        group.rows.append(r)
        group.methods += [m for m in INFERENCE_METHODS if m in by_row.get(r, ()) and m not in group.methods]
    return list(groups.values())


def _rows_text(rows: list[int]) -> str:
    numbers = [r + 1 for r in rows]
    if len(numbers) > 2 and numbers == list(range(numbers[0], numbers[-1] + 1)):
        return f"{numbers[0]}–{numbers[-1]}"
    return ", ".join(str(n) for n in numbers)


def inferred_frame(batch: Batch, schema: Schema, columns: list[str], groups: list[InferredGroup]) -> pd.DataFrame:
    """Table of the inferred values, one row per file (the value of its first row)."""
    index = pd.RangeIndex(len(groups))
    mixed = [[c for c in columns if len({to_text(batch.df.at[r, c]) for r in g.rows}) > 1] for g in groups]
    data = {
        FILE_COL: pd.Series([g.name for g in groups], index=index, dtype="string"),
        ROWS_COL: pd.Series([_rows_text(g.rows) for g in groups], index=index, dtype="string"),
        FROM_COL: pd.Series([", ".join(INFERENCE_METHODS[m] for m in g.methods) or "nessun valore dedotto"
                             for g in groups], index=index, dtype="string"),
        MIXED_COL: pd.Series([", ".join(m) for m in mixed], index=index, dtype="string"),
    }
    for column in columns:
        data[column] = _display_series([batch.df.at[g.rows[0], column] for g in groups],
                                       _kind(batch, schema, column), index)
    return pd.DataFrame(data, index=index)


def apply_group_edits(batch: Batch, schema: Schema, groups: list[InferredGroup], edited_rows: dict) -> Batch:
    """Edits of the inferred-values table: each value goes to every row of its file."""
    if not edited_rows:
        return batch
    result = batch.copy()
    for position, changes in edited_rows.items():
        group = groups[int(position)]
        for column, value in changes.items():
            if column in INFERRED_INFO_COLUMNS or column not in result.df.columns:
                continue
            for row in group.rows:
                _set_edit(result, schema, row, column, value)
    return result


def problems_text(issues: list[Issue]) -> dict[int, str]:
    by_row: dict[int, list[str]] = {}
    for issue in issues:
        if issue.row is not None and issue.severity in (ERROR, WARNING):
            prefix = "✖" if issue.severity == ERROR else "!"
            by_row.setdefault(issue.row, []).append(f"{prefix} {issue.column + ': ' if issue.column else ''}"
                                                    f"{issue.message}")
    return {row: " | ".join(messages) for row, messages in by_row.items()}


def missing_support_values(issues: list[Issue], schema: Schema,
                           options: dict[str, list[str]]) -> dict[tuple[str, str], list[int]]:
    """(column, value) -> rows, for the choice values missing from their support sheet in the DB."""
    out: dict[tuple[str, str], list[int]] = {}
    for issue in issues:
        f = schema.get(issue.column) if issue.column else None
        value = to_text(issue.value)
        if f is None or f.type != "choice" or not f.uses_support_sheet or not value or issue.row is None:
            continue
        if value not in options.get(f.column, []):
            rows = out.setdefault((f.column, value), [])
            if issue.row not in rows:
                rows.append(issue.row)
    return out


def issues_frame(issues: list[Issue]) -> pd.DataFrame:
    order = {ERROR: 0, WARNING: 1, INFO: 2}
    ordered = sorted(issues, key=lambda i: (order.get(i.severity, 3), i.row if i.row is not None else -1))
    return pd.DataFrame({
        "Gravità": [SEVERITY_LABEL.get(i.severity, i.severity) for i in ordered],
        "Riga": pd.Series([i.row + 1 if i.row is not None else None for i in ordered], dtype="Int64"),
        "Colonna": [i.column or "" for i in ordered],
        "Messaggio": [i.message for i in ordered],
    })


# --- exports / comparisons -----------------------------------------------------

def batch_excel_bytes(batch: Batch) -> bytes:
    buffer = io.BytesIO()
    export_batch_excel(batch, buffer)
    return buffer.getvalue()


def one_row_excel_bytes(values: dict[str, Any]) -> bytes:
    """Form values as a one-row Excel with DB headers (like 'Export template to excel')."""
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.append(list(values))
    row = []
    for value in values.values():
        if isinstance(value, dt.date) and not isinstance(value, dt.datetime):
            value = dt.datetime(value.year, value.month, value.day)
        row.append(value)
    ws.append(row)
    for cell in ws[2]:
        if isinstance(cell.value, dt.datetime):
            cell.number_format = "dd/mm/yyyy"
    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


def diff_frame(schema: Schema, original: dict[str, Any], new: dict[str, Any]) -> pd.DataFrame:
    """Columns whose value changes when overwriting a DB row."""
    rows = []
    for column, value in new.items():
        f = schema.get(column)
        if f is None or f.type == "id":
            continue
        before = original.get(column)
        if canonical(before, f.type) != canonical(value, f.type):
            rows.append({"Colonna": column, "Valore attuale nel DB": to_text(before) or "",
                         "Nuovo valore": to_text(value) or ""})
    return pd.DataFrame(rows, columns=["Colonna", "Valore attuale nel DB", "Nuovo valore"])
