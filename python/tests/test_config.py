import pytest

from kpimeta.config import ConfigError, read_csv_records
from kpimeta.schema import load_schema

# Order of GeneralMetadataMap + ICMetadataMap in IC_KPIs_App.mlapp (initializeDbMappings)
MATLAB_COLUMNS = [
    "ID", "IDName", "Date", "Cluster", "Sub_Cluster", "RefFolder", "Results_Category", "Brand", "Model",
    "Specs", "VIN", "RefName", "ManeuvreName", "StartSpeed", "EndSpeed", "Repetition_ID", "Track", "Surface",
    "Condition", "Road_mu_coeff", "Driving_Mode", "ABS", "ESC", "TCS", "T_Front_Name", "T_Rear_Name",
    "T_Front_type", "T_Rear_type", "FileName", "MassInfo", "Driver",
    "B_F_System", "B_F_BL", "B_F_Cal_Date", "B_F_Cal_ID", "B_F_Type", "B_R_System", "B_R_BL", "B_R_Type",
    "B_R_Cal_Date", "B_R_Cal_ID", "Su_System", "Su_Type", "Su_Mode", "Su_BL_Cal_Date", "Su_Cal_ID", "Su_BL",
    "SmT_Type", "SmT_BL_Cal_Date", "SmT_BL", "SmT_Cal_ID", "St_Type", "St_Cal_Date", "St_BL", "St_Cal_ID",
    "St_Mode", "CoOpType", "CoOpBL_Cal_Date", "CoOpBL", "CoOp_Cal_ID", "B_F_BL_Phase", "B_R_BL_Phase",
    "T_Front_Pressure", "T_Rear_Pressure", "T_Front_Thread_Depth", "T_Rear_Thread_Depth", "SmT_SW_Version",
    "SmT_SW_TireID", "SmT_SW_Road", "SmT_SW_Wear", "SmT_SW_Mass",
]


def test_schema_matches_matlab_maps(schema):
    assert schema.columns == MATLAB_COLUMNS
    assert schema.kpi_table_columns() == ["ID", "IDName", "Date"]
    assert schema.get("ABS").choices == ("ON", "OFF")
    assert schema.get("StartSpeed").max == 300 and schema.get("EndSpeed").max == 200
    assert (schema.get("B_F_Type").choices_sheet, schema.get("B_F_Type").choices_column) == ("Brake", "BrakeType")
    assert schema.get("Repetition_ID").lola_policy == "keep"
    assert schema.get("StartSpeed").lola_policy == "keep_if_valid"


def test_settings_defaults_match_matlab(settings):
    assert (settings.metadata_table, settings.kpi_table) == ("TableMet", "TableKPI")
    assert settings.filename_tokens == {"Date": 3, "Driving_Mode": 5, "ManeuvreName": 6, "Driver": 7}
    assert settings.template_tokens == 2
    assert settings.date_storage == "text"


def test_csv_reader_reports_broken_rows(tmp_path):
    path = tmp_path / "bad.csv"
    path.write_text("a;b\n1;2\n3;4;5\n", encoding="utf-8")
    with pytest.raises(ConfigError, match="riga 3"):
        read_csv_records(path, ["a", "b"])


def test_csv_reader_accepts_excel_italian_files(tmp_path):
    path = tmp_path / "it.csv"
    path.write_bytes("A;B\nperché;città\n".encode("cp1252"))
    assert read_csv_records(path, ["a", "b"])[0]["a"] == "perché"
    comma = tmp_path / "comma.csv"
    comma.write_text('a,b\n"x, y",2\n', encoding="utf-8")
    assert read_csv_records(comma, ["a", "b"])[0]["a"] == "x, y"


def test_schema_errors_are_explicit(tmp_path, settings):
    text = settings.schema_path.read_text(encoding="utf-8").replace(";ABS;General;Vehicle settings;choice;",
                                                                     ";ABS;General;Vehicle settings;bogus;")
    path = tmp_path / "schema.csv"
    path.write_text(text, encoding="utf-8")
    with pytest.raises(ConfigError, match="bogus"):
        load_schema(path)
