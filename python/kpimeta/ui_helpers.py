"""UI logic that does not depend on Streamlit (unit-testable): form values,
review grid <-> batch conversion, exports."""

from __future__ import annotations

import datetime as dt
import io
from typing import Any

import pandas as pd

from .batch import Batch, export_batch_excel, set_value
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
        values = [batch.df.at[r, column] for r in rows]
        kind = _kind(batch, schema, column)
        if kind in ("int", "id"):
            series = pd.Series([to_int_safe(v) for v in values], index=index, dtype="Int64")
        elif kind == "number":
            series = pd.Series([to_float_safe(v) for v in values], index=index, dtype="float64")
        elif kind == "date":
            series = pd.Series(pd.to_datetime([v if isinstance(v, dt.date) else None for v in values]),
                               index=index)
        else:
            series = pd.Series([to_text(v) for v in values], index=index, dtype="string")
        data[column] = series
    return pd.DataFrame(data, index=index)


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
            kind = _kind(result, schema, column)
            if isinstance(value, str) and kind == "date" and "T" in value:
                value = value.split("T")[0]
            set_value(result, row, column, value, kind, "edit")
    return result


def problems_text(issues: list[Issue]) -> dict[int, str]:
    by_row: dict[int, list[str]] = {}
    for issue in issues:
        if issue.row is not None and issue.severity in (ERROR, WARNING):
            prefix = "✖" if issue.severity == ERROR else "!"
            by_row.setdefault(issue.row, []).append(f"{prefix} {issue.column + ': ' if issue.column else ''}"
                                                    f"{issue.message}")
    return {row: " | ".join(messages) for row, messages in by_row.items()}


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
