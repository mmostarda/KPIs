import datetime as dt

import pandas as pd

from kpimeta.batch import batch_from_frame, batch_manual, compose
from kpimeta.issues import ERROR, WARNING, Issue
from kpimeta.lola import read_lola_files
from kpimeta.mapping import load_mapping
from kpimeta.ui_helpers import (FROM_COL, MIXED_COL, PROBLEMS_COL, ROWS_COL, apply_edits, apply_group_edits,
                                diff_frame, display_frame, field_value, form_defaults, inferred_columns,
                                inferred_frame, inferred_groups, missing_support_values,
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


def test_missing_support_values(schema):
    issues = [Issue(WARNING, "'Brand Z' non è nel foglio Brand", 0, "Brand", "Brand Z"),
              Issue(WARNING, "'Brand Z' non è nel foglio Brand", 2, "Brand", "Brand Z"),
              Issue(ERROR, "'X' non è nella lista", 1, "ABS", "X"),  # fixed list, no support sheet
              Issue(ERROR, "valore non valido", 1, "KPI_A", "n.a.")]
    assert missing_support_values(issues, schema, {"Brand": ["Brand X"]}) == {("Brand", "Brand Z"): [0, 2]}
    assert missing_support_values(issues, schema, {"Brand": ["Brand X", "Brand Z"]}) == {}


def test_inferred_values_are_edited_per_file(demo, schema, settings):
    frame, _ = read_lola_files(demo["lola"], load_mapping(demo["mapping"]), settings)
    batch = compose(batch_from_frame(frame, schema, settings, "lola"), schema, settings, {}, infer_filename=True)
    columns, groups = inferred_columns(batch), inferred_groups(batch)
    assert columns == ["Date", "ManeuvreName", "Driving_Mode", "Driver"]
    assert [g.rows for g in groups] == [[0, 1, 2], [3, 4], [5], [6], [7]]  # one group per file
    assert groups[4].name == "shortname.mf4" and groups[4].methods == []  # not recognised: filled by hand
    table = inferred_frame(batch, schema, columns, groups)
    assert table.at[0, "Driver"] == "DriverA" and table.at[0, ROWS_COL] == "1–3"
    assert table.at[0, FROM_COL] == "nome file" and table.at[4, FROM_COL] == "nessun valore dedotto"

    edited = apply_group_edits(batch, schema, groups, {"0": {"Driver": "Mario", "File": "ignored"},
                                                       "4": {"Date": "2026-05-02T00:00:00.000"}})
    assert [edited.df.at[r, "Driver"] for r in range(4)] == ["Mario", "Mario", "Mario", "DriverB"]
    assert edited.df.at[7, "Date"] == dt.date(2026, 5, 2)
    assert edited.origin[(0, "Driver")] == "edit" and batch.df.at[0, "Driver"] == "DriverA"
    assert [g.rows for g in inferred_groups(edited)] == [g.rows for g in groups]  # same table after edits

    one_row = apply_edits(batch, schema, list(range(batch.n_rows)), {"1": {"Driver": "Other"}})
    assert inferred_frame(one_row, schema, columns, inferred_groups(one_row)).at[0, MIXED_COL] == "Driver"
