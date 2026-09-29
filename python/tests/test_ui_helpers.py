import datetime as dt

import pandas as pd

from kpimeta.batch import batch_from_frame, batch_manual
from kpimeta.ui_helpers import (PROBLEMS_COL, apply_edits, diff_frame, display_frame, field_value, form_defaults,
                                one_row_excel_bytes, values_to_widgets, widget_key, widget_value)


def _batch(schema, settings):
    frame = pd.DataFrame([{"IDName": "a", "Date": "05/03/2026", "KPI_A": 1.5, "StartSpeed": 10},
                          {"IDName": None, "Date": None, "KPI_A": "n.a.", "StartSpeed": None}], dtype=object)
    return batch_from_frame(frame, schema, settings, "excel")


def test_display_dtypes_do_not_depend_on_values(schema, settings):
    batch = _batch(schema, settings)
    full = display_frame(batch, schema, {0: "x"})
    empty = display_frame(batch_manual(schema, {}, ["KPI_A"]), schema, {})
    common = [c for c in full.columns if c in empty.columns]
    assert all(full[c].dtype == empty[c].dtype for c in common)  # editor identity stays stable
    assert list(full.index) == [1, 2] and full.at[1, PROBLEMS_COL] == "x"
    assert full.at[1, "Date"] == pd.Timestamp(2026, 3, 5)


def test_apply_edits_maps_view_positions(schema, settings):
    batch = _batch(schema, settings)
    edited = apply_edits(batch, schema, [1], {"0": {"KPI_A": 2.5, "Date": "2026-04-01T00:00:00.000",
                                                    "IDName": "b", "ID": 99, PROBLEMS_COL: "zz"}})
    assert edited.df.at[1, "KPI_A"] == 2.5 and (1, "KPI_A") not in edited.conv_issues
    assert edited.df.at[1, "Date"] == dt.date(2026, 4, 1) and edited.df.at[1, "IDName"] == "b"
    assert edited.df.at[1, "ID"] is None  # ID is never editable
    assert batch.df.at[1, "IDName"] is None  # original untouched


def test_form_round_trip(schema):
    defaults = form_defaults(schema, with_defaults=True)
    assert defaults[widget_key("ABS")] == "ON" and defaults[widget_key("Brand")] == ""
    widgets = values_to_widgets(schema, {"Date": "05/03/2026", "CoOpType": "Brake;Tire", "StartSpeed": "12",
                                         "Repetition_ID": 3.0, "unknown": 1})
    assert widgets == {widget_key("Date"): dt.date(2026, 3, 5), widget_key("CoOpType"): ["Brake", "Tire"],
                       widget_key("StartSpeed"): 12.0, widget_key("Repetition_ID"): 3}
    assert field_value(schema.get("CoOpType"), ["Brake", "Tire"]) == "Brake;Tire"
    assert field_value(schema.get("IDName"), "  ") is None
    assert widget_value(schema.get("Date"), "not a date") is None


def test_diff_and_export(schema):
    diff = diff_frame(schema, {"ID": 1, "IDName": "a", "Date": "05/03/2026"},
                      {"IDName": "b", "Date": dt.date(2026, 3, 5)})
    assert list(diff["Colonna"]) == ["IDName"]  # same date in another format is not a change
    assert one_row_excel_bytes({"IDName": "x", "Date": dt.date(2026, 3, 5)})[:2] == b"PK"
