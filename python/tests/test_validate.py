import pandas as pd

from kpimeta.batch import batch_manual
from kpimeta.issues import ERROR, WARNING
from kpimeta.validate import (KpiDecision, canonicalize_choices, check_decisions, duplicates_vs_db,
                              internal_duplicates, kpi_resolution, validate_batch)


def _one(schema, **values):
    return batch_manual(schema, values)


def _issues(batch, schema, settings, options=None):
    return {(i.column, i.severity) for i in validate_batch(batch, schema, settings, options)}


def test_limits_and_fixed_lists(schema, settings):
    batch = _one(schema, StartSpeed=350, EndSpeed=-1)
    batch.df.at[0, "ABS"] = "MAYBE"
    issues = _issues(batch, schema, settings)
    assert ("StartSpeed", ERROR) in issues and ("EndSpeed", ERROR) in issues and ("ABS", ERROR) in issues


def test_support_lists_warn_or_block(schema, settings):
    batch = _one(schema, Brand="Unknown brand")
    options = {"Brand": ["Brand X"]}
    assert ("Brand", WARNING) in _issues(batch, schema, settings, options)
    settings.strict_choices = True
    assert ("Brand", ERROR) in _issues(batch, schema, settings, options)


def test_case_is_fixed_when_unambiguous(schema):
    batch = _one(schema, ABS="on", Brand="brand x", CoOpType="brake;SUSP")
    fixed = canonicalize_choices(batch, schema, {"Brand": ["Brand X"]})
    assert fixed == 3
    assert (batch.df.at[0, "ABS"], batch.df.at[0, "Brand"], batch.df.at[0, "CoOpType"]) == ("ON", "Brand X", "Brake;Susp")


def test_required_flag(schema, settings):
    from dataclasses import replace

    from kpimeta.schema import Schema

    strict = Schema([replace(f, required=True) if f.column == "IDName" else f for f in schema.fields])
    batch = batch_manual(strict, {})
    assert ("IDName", ERROR) in _issues(batch, strict, settings)


def test_duplicates(schema):
    batch = batch_manual(schema, {"IDName": "x", "Date": "05/03/2026"})
    batch.df = pd.concat([batch.df, batch.df], ignore_index=True)
    assert len(internal_duplicates(batch, schema)) == 1
    db = pd.DataFrame([{"ID": 12, "IDName": "x", "Date": "05/03/2026"}], dtype=object)
    for c in schema.columns:
        if c not in db.columns:
            db[c] = None
    found = duplicates_vs_db(batch, schema, db)
    assert len(found) == 2 and "ID 12" in found[0].message


def test_kpi_resolution(schema, settings):
    batch = batch_manual(schema, {}, ["KPI_A", "kpi b", "NEW", "ID_Template"])
    known, unknown, suggestions = kpi_resolution(batch, ["KPI_A", "KPI_B"], settings)
    assert known == ["KPI_A"] and unknown == ["kpi b", "NEW"]
    assert suggestions == {"kpi b": "KPI_B"}
    missing = check_decisions(batch, ["KPI_A", "KPI_B"], {}, settings)
    assert len(missing) == 2
    clash = check_decisions(batch, ["KPI_A", "KPI_B"], {"kpi b": KpiDecision("map", "KPI_A"),
                                                         "NEW": KpiDecision("discard")}, settings)
    assert len(clash) == 1 and "stessa colonna" in clash[0].message
    ok = check_decisions(batch, ["KPI_A", "KPI_B"], {"kpi b": KpiDecision("map", "KPI_B"),
                                                      "NEW": KpiDecision("add")}, settings)
    assert ok == []
