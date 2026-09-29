import datetime as dt

import pytest

from kpimeta.values import (ConversionError, canonical, coerce, is_missing, split_multi, to_date, to_int,
                            to_number, to_text)


@pytest.mark.parametrize("raw", ["05/03/2026", "5/3/2026", "2026-03-05", "20260305", "2026-03-05 00:00:00",
                                 46086, 46086.0, dt.datetime(2026, 3, 5), dt.date(2026, 3, 5)])
def test_dates_accepted(raw):
    assert to_date(raw) == (dt.date(2026, 3, 5), None)


def test_ambiguous_date_is_day_first():
    # MATLAB datetime() without a format read 01/02/2026 as January 2nd
    assert to_date("01/02/2026")[0] == dt.date(2026, 2, 1)


def test_month_first_only_when_unavoidable():
    day, warning = to_date("03/15/2026")
    assert day == dt.date(2026, 3, 15) and "mm/gg" in warning


@pytest.mark.parametrize("raw", [None, "", "  ", "NaT", "NA", "n/a", float("nan"), "<undefined>"])
def test_missing(raw):
    assert is_missing(raw)
    assert to_date(raw) == (None, None)
    assert to_number(raw) is None
    assert to_text(raw) is None


def test_matlab_datenum_bug_is_detected_with_the_right_date():
    with pytest.raises(ConversionError, match="01/01/2026"):
        to_date(739983)  # datenum(2026,1,1) written into an Excel cell
    with pytest.raises(ConversionError, match="datenum"):
        to_date(dt.datetime(3926, 1, 1))  # the same number shown by Excel as a date


def test_invalid_dates():
    for raw in ["31/02/2026", "hello", "2026315"]:
        with pytest.raises(ConversionError):
            to_date(raw)


def test_numbers():
    assert to_number("1.5") == 1.5
    assert to_number(" 12 ") == 12
    assert to_number("1e3") == 1000
    assert to_number(True) == 1.0
    with pytest.raises(ConversionError, match="punto"):
        to_number("1,5")  # MATLAB str2double returned 15
    with pytest.raises(ConversionError):
        to_number("abc")
    assert to_int("7") == 7
    with pytest.raises(ConversionError):
        to_int(7.5)


def test_text():
    assert to_text(5.0) == "5"
    assert to_text(0.1 + 0.2) == "0.3"
    assert to_text(dt.date(2026, 3, 5)) == "05/03/2026"
    assert to_text("  x  ") == "x"


def test_multi_and_coerce():
    assert split_multi("Brake; Susp;;Brake") == ["Brake", "Susp"]
    assert coerce("multichoice", "Brake,Tire") == ("Brake;Tire", None)
    assert coerce("int", "3") == (3, None)
    assert coerce("date", "20260305") == (dt.date(2026, 3, 5), None)


def test_canonical_keys_match_across_representations():
    assert canonical("05/03/2026", "date") == canonical(46086, "date") == canonical(dt.datetime(2026, 3, 5), "date")
    assert canonical("1", "number") == canonical(1.0, "number")
    assert canonical("Susp;Brake", "multichoice") == canonical("Brake;Susp", "multichoice")
