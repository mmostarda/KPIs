import numpy as np
import pytest
from openpyxl import Workbook

from kpimeta.mapping import MappingError, load_mapping, save_mapping_csv


def test_mat_file_like_load_mapping_from_excel(tmp_path):
    from scipy.io import savemat

    path = tmp_path / "Mapping_LoLa_DB_mapping.mat"
    savemat(path, {"mapping": {"LoLa_Name": np.array(["File Name", "Kpi A", "Kpi B"], dtype=object),
                               "DB_Name": np.array(["FileName", "KPI_A", "KPI_B"], dtype=object)}})
    mapping = load_mapping(path)
    assert [(e.lola_name, e.db_name) for e in mapping.entries] == [
        ("File Name", "FileName"), ("Kpi A", "KPI_A"), ("Kpi B", "KPI_B")]
    csv_path = save_mapping_csv(mapping, tmp_path / "lola_mapping.csv")
    assert load_mapping(csv_path).entries == mapping.entries


def test_original_excel_mapping_skips_na_rows(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.append(["KPI_DB_Name", "KPI_LoLa_Name"])
    ws.append(["FileName", "File Name"])
    ws.append(["KPI_A", "NA"])
    ws.append([None, "Kpi X"])
    ws.append(["KPI_B", "Kpi B"])
    path = tmp_path / "Mapping_LoLa_DB.xlsx"
    wb.save(path)
    mapping = load_mapping(path)
    assert mapping.db_names() == ["FileName", "KPI_B"]
    assert mapping.skipped_rows == 2


def test_duplicates_are_rejected(tmp_path):
    path = tmp_path / "m.csv"
    path.write_text("KPI_LoLa_Name;KPI_DB_Name\nA;X\nB;X\n", encoding="utf-8")
    with pytest.raises(MappingError, match="'X'"):
        load_mapping(path)


def test_type_column_is_validated(tmp_path):
    path = tmp_path / "m.csv"
    path.write_text("KPI_LoLa_Name;KPI_DB_Name;Type\nA;X;datetime\n", encoding="utf-8")
    with pytest.raises(MappingError, match="datetime"):
        load_mapping(path)


def test_missing_file_explains_how_to_create_it(tmp_path):
    with pytest.raises(MappingError, match="convert-mapping"):
        load_mapping(tmp_path / "nope.csv")
