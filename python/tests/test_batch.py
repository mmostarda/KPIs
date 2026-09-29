import datetime as dt

import pandas as pd
import pytest

from kpimeta.batch import (BatchError, apply_common, batch_from_excel, batch_from_frame, batch_manual, compose,
                           export_batch_excel)
from kpimeta.excel_db import read_db
from kpimeta.lola import read_lola_files
from kpimeta.mapping import load_mapping


@pytest.fixture()
def lola_batch(demo, schema, settings):
    frame, _ = read_lola_files(demo["lola"], load_mapping(demo["mapping"]), settings)
    snapshot = read_db(demo["db"], schema, settings)
    return batch_from_frame(frame, schema, settings, "lola", snapshot.metadata.headers), snapshot


def test_split_metadata_and_kpis(lola_batch):
    batch, _ = lola_batch
    assert batch.kpi_columns == ["KPI_A", "KPI_B", "KPI_C", "KPI_D"]
    assert batch.meta_columns[:4] == ["ID", "Repetition_ID", "FileName", "StartSpeed"]
    assert batch.df.at[0, "KPI_C"] is None and (0, "KPI_C") in batch.conv_issues


def test_precedence(lola_batch, schema, settings):
    base, snapshot = lola_batch
    common = {"Repetition_ID": 99, "FileName": "x", "StartSpeed": 50, "Brand": "Brand Y", "Driver": "Form",
              "IDName": "", "Track": None}
    batch = compose(base, schema, settings, common, templates=snapshot.templates,
                    infer_template=True, infer_filename=True)
    df = batch.df
    assert list(df["Repetition_ID"])[:3] == [1, 2, 3]  # protected: always from LoLa
    assert df.at[0, "FileName"].startswith("PRJ1_VEH1")
    assert df.at[0, "StartSpeed"] == 100.2 and df.at[2, "StartSpeed"] == 50  # form only where LoLa is empty
    assert df.at[0, "Brand"] == "Brand X"  # template beats form
    assert df.at[6, "Brand"] == "Brand Y"  # template not found -> form
    assert df.at[0, "Driver"] == "DriverA"  # file name beats form
    assert df.at[7, "Driver"] == "Form"  # file name not recognised -> form
    assert df.at[0, "Date"] == dt.date(2026, 3, 15)
    assert df.at[0, "B_F_Cal_Date"] == dt.date(2026, 1, 10)
    assert df.at[0, "Track"] == "Track 1"  # empty form values never overwrite
    assert any("PRJ9_VEH9" in n.message for n in batch.notes)


def test_fill_mode_only_fills_empty_cells(lola_batch, schema):
    batch, _ = lola_batch
    batch = batch.copy()
    batch.df.at[0, "Track"] = "Mine"
    apply_common(batch, schema, {"Track": "Form"}, mode="fill")
    assert batch.df.at[0, "Track"] == "Mine" and batch.df.at[1, "Track"] == "Form"


def test_excel_import_roundtrip(lola_batch, schema, settings, tmp_path):
    batch, snapshot = lola_batch
    path = export_batch_excel(batch, tmp_path / "Metadata_ALL_DB.xlsx")
    again = batch_from_excel(path, schema, settings, snapshot.metadata.headers)
    assert again.kpi_columns == batch.kpi_columns
    assert again.n_rows == batch.n_rows
    assert again.conv_issues.keys() == batch.conv_issues.keys()  # 'n.a.' kept so it can be fixed
    pd.testing.assert_series_equal(again.df["KPI_A"], batch.df["KPI_A"])


def test_excel_import_rules(tmp_path, schema, settings):
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.append(["id", "IDName", "ID_Template", "Date", "My KPI", "StartSpeed"])
    ws.append([7, "a", 1, "05/03/2026", "1.25", 10])
    ws.append([None, None, None, None, None, None])
    ws.append([8, "b", 1, dt.datetime(2026, 3, 6), 3, "fast"])
    path = tmp_path / "import.xlsx"
    wb.save(path)
    batch = batch_from_excel(path, schema, settings)
    assert batch.n_rows == 2  # empty row skipped
    assert batch.kpi_columns == ["My KPI"]  # ID_Template never a KPI, 'id' matched to ID
    assert list(batch.df["ID"]) == [None, None]  # IDs are assigned when writing
    assert list(batch.df["My KPI"]) == [1.25, 3.0]
    assert batch.df.at[1, "Date"] == dt.date(2026, 3, 6)
    assert (1, "StartSpeed") in batch.conv_issues


def test_duplicate_headers_are_rejected(tmp_path, schema, settings):
    from openpyxl import Workbook

    wb = Workbook()
    wb.active.append(["IDName", "IDName"])
    wb.active.append(["a", "b"])
    wb.save(tmp_path / "dup.xlsx")
    with pytest.raises(BatchError, match="ripetute"):
        batch_from_excel(tmp_path / "dup.xlsx", schema, settings)


def test_manual_batch(schema):
    batch = batch_manual(schema, {"IDName": "one", "Date": dt.date(2026, 1, 2), "ABS": "OFF"}, ["KPI_A"])
    assert batch.n_rows == 1
    assert batch.df.at[0, "ABS"] == "OFF" and batch.kpi_columns == ["KPI_A"]
