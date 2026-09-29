"""Checks before writing: types/limits/lists from the schema (no GUI side effects,
unlike MATLAB validateImportedMetadata), duplicates, unknown KPI columns."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any

import pandas as pd

from .batch import Batch
from .config import Settings
from .issues import ERROR, INFO, WARNING, Issue
from .schema import Schema
from .values import ConversionError, canonical, coerce, normalize_name, split_multi, to_text


def canonicalize_choices(batch: Batch, schema: Schema, options: dict[str, list[str]]) -> int:
    """'on' -> 'ON', 'brembo' -> 'Brembo' when the match in the list is unique."""
    fixed = 0
    for column in batch.meta_columns:
        f = schema.get(column)
        if f is None or f.type not in ("choice", "multichoice"):
            continue
        allowed = list(f.choices) or options.get(column) or []
        by_lower: dict[str, list[str]] = {}
        for option in allowed:
            by_lower.setdefault(option.lower(), []).append(option)
        for r in range(batch.n_rows):
            value = batch.df.at[r, column]
            if value is None:
                continue
            if f.type == "choice":
                match = by_lower.get(str(value).lower(), [])
                if value not in allowed and len(match) == 1:
                    batch.df.at[r, column] = match[0]
                    fixed += 1
            else:
                tokens = split_multi(value)
                new = [by_lower[t.lower()][0] if t not in allowed and len(by_lower.get(t.lower(), [])) == 1
                       else t for t in tokens]
                if new != tokens:
                    batch.df.at[r, column] = ";".join(new)
                    fixed += 1
    return fixed


def validate_batch(batch: Batch, schema: Schema, settings: Settings,
                   options: dict[str, list[str]] | None = None) -> list[Issue]:
    """Issues for every cell. `options` = values of the support sheets per column."""
    options = options or {}
    issues: list[Issue] = []
    for (row, column), (raw, message) in sorted(batch.conv_issues.items()):
        if row < batch.n_rows and batch.df.at[row, column] is None:
            issues.append(Issue(ERROR, f"valore '{raw}' non valido: {message}", row, column, raw))
    for column in batch.meta_columns:
        f = schema.get(column)
        if f is None or f.type == "id":
            continue
        allowed = list(f.choices) or options.get(column) or []
        for row in range(batch.n_rows):
            value = batch.df.at[row, column]
            if value is None:
                if f.required and (row, column) not in batch.conv_issues:
                    issues.append(Issue(ERROR, "campo obbligatorio", row, column))
                continue
            try:
                value, _ = coerce(f.type, value)
            except ConversionError as exc:
                issues.append(Issue(ERROR, str(exc), row, column, value))
                continue
            if f.type in ("number", "int"):
                if f.min is not None and value < f.min:
                    issues.append(Issue(ERROR, f"{value:g} è minore del minimo {f.min:g}", row, column, value))
                if f.max is not None and value > f.max:
                    issues.append(Issue(ERROR, f"{value:g} è maggiore del massimo {f.max:g}", row, column, value))
            elif f.type == "choice" and allowed and value not in allowed:
                strict = bool(f.choices) or settings.strict_choices
                where = f"foglio {f.choices_sheet}" if f.uses_support_sheet else "lista"
                issues.append(Issue(ERROR if strict else WARNING,
                                    f"'{value}' non è nella {where}", row, column, value))
            elif f.type == "multichoice" and allowed:
                unknown = [t for t in split_multi(value) if t not in allowed]
                if unknown:
                    issues.append(Issue(WARNING, f"valori non previsti: {', '.join(unknown)} "
                                                 f"(ammessi: {', '.join(allowed)})", row, column, value))
    for column in batch.kpi_columns:
        for row in range(batch.n_rows):
            value = batch.df.at[row, column]
            if value is None:
                continue
            try:
                coerce("number", value)
            except ConversionError as exc:
                issues.append(Issue(ERROR, str(exc), row, column, value))
    issues.extend(batch.notes)
    return issues


def _row_keys(frame: pd.DataFrame, columns: list[str], schema: Schema) -> list[tuple[str, ...]]:
    """One comparable key per row; values are converted column by column with a cache
    (the same values repeat a lot, e.g. Brand or Date)."""
    cache: dict[tuple, str] = {}
    keyed_columns = []
    for column in columns:
        field_type = schema.type_of(column)
        keys = []
        for value in frame[column].tolist():
            try:
                cache_key = (field_type, type(value), value)
                key = cache.get(cache_key)
            except TypeError:  # unhashable value
                cache_key, key = None, None
            if key is None:
                key = canonical(value, field_type)
                if cache_key is not None:
                    cache[cache_key] = key
            keys.append(key)
        keyed_columns.append(keys)
    if not keyed_columns:
        return [() for _ in range(len(frame))]
    return list(zip(*keyed_columns))


def internal_duplicates(batch: Batch, schema: Schema) -> list[Issue]:
    """Identical metadata rows inside the batch (MATLAB checkInternalDuplicates)."""
    id_column = schema.id_field().column
    columns = [c for c in batch.meta_columns if c != id_column]
    groups: dict[tuple[str, ...], list[int]] = {}
    for row, key in zip(range(batch.n_rows), _row_keys(batch.df, columns, schema)):
        groups.setdefault(key, []).append(row)
    issues = []
    for rows in groups.values():
        if len(rows) > 1:
            numbers = ", ".join(str(r + 1) for r in rows)
            issues.append(Issue(ERROR, f"metadati identici nelle righe {numbers}", rows[1]))
    return issues


def duplicates_vs_db(batch: Batch, schema: Schema, db_metadata: pd.DataFrame) -> list[Issue]:
    """Rows already present in TableMet: same values in all common columns except ID
    (MATLAB checkMetadataDuplicatesAgainstDB)."""
    id_column = schema.id_field().column
    columns = [c for c in batch.meta_columns if c != id_column and c in db_metadata.columns]
    if not columns or db_metadata.empty:
        return []
    db_index: dict[tuple[str, ...], list[str]] = {}
    ids = db_metadata[id_column] if id_column in db_metadata.columns else [None] * len(db_metadata)
    for key, db_id in zip(_row_keys(db_metadata.reset_index(drop=True), columns, schema), ids):
        db_index.setdefault(key, []).append(to_text(db_id) or "?")
    issues = []
    for row, key in zip(range(batch.n_rows), _row_keys(batch.df, columns, schema)):
        if key in db_index:
            issues.append(Issue(ERROR, f"riga già presente nel DB (ID {', '.join(db_index[key])})", row))
    return issues


@dataclass
class KpiDecision:
    action: str  # keep | add | map | discard
    target: str | None = None


def kpi_resolution(batch: Batch, db_kpi_columns: list[str],
                   settings: Settings) -> tuple[list[str], list[str], dict[str, str]]:
    """(known KPI, unknown KPI, suggested existing column for each unknown one)."""
    excluded = {c.lower() for c in settings.kpi_excluded_columns}
    imported = [c for c in batch.kpi_columns if c.lower() not in excluded]
    known = [c for c in imported if c in db_kpi_columns]
    unknown = [c for c in imported if c not in db_kpi_columns]
    suggestions = {}
    by_norm: dict[str, list[str]] = {}
    for existing in db_kpi_columns:
        by_norm.setdefault(normalize_name(existing), []).append(existing)
    for name in unknown:
        match = by_norm.get(normalize_name(name), [])
        if len(match) == 1:
            suggestions[name] = match[0]
    return known, unknown, suggestions


def check_decisions(batch: Batch, db_kpi_columns: list[str], decisions: dict[str, KpiDecision],
                    settings: Settings) -> list[Issue]:
    known, unknown, _ = kpi_resolution(batch, db_kpi_columns, settings)
    issues = []
    targets: dict[str, str] = {k: k for k in known}
    for name in unknown:
        decision = decisions.get(name)
        if decision is None or decision.action not in ("add", "map", "discard"):
            issues.append(Issue(ERROR, f"KPI '{name}' non presente in TableKPI: scegli aggiungi, "
                                       "mappa su una colonna esistente o scarta", column=name))
            continue
        if decision.action == "discard":
            continue
        target = name if decision.action == "add" else (decision.target or "")
        if decision.action == "map" and target not in db_kpi_columns:
            issues.append(Issue(ERROR, f"KPI '{name}': colonna di destinazione '{target}' inesistente",
                                column=name))
            continue
        if target in targets:
            issues.append(Issue(ERROR, f"'{name}' e '{targets[target]}' finirebbero nella stessa colonna "
                                       f"'{target}'", column=name))
        targets[target] = name
    return issues


def summarize(issues: list[Issue]) -> dict[str, int]:
    counts = {ERROR: 0, WARNING: 0, INFO: 0}
    for issue in issues:
        counts[issue.severity] = counts.get(issue.severity, 0) + 1
    return counts


def issues_by_row(issues: list[Issue]) -> dict[int, list[Issue]]:
    rows: dict[int, list[Issue]] = {}
    for issue in issues:
        if issue.row is not None:
            rows.setdefault(issue.row, []).append(issue)
    return rows


def describe_value(value: Any) -> str:
    return to_text(value) or ""
