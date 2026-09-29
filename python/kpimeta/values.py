"""Value conversions: the single place where missing values, numbers and dates
are interpreted.

In the MATLAB app the same job was spread over applyRowToGUI, writeStructToGUI,
assignValueToGUIComponent, sanitizeKpiValue and cell2str_safe/cell2date_safe,
each with slightly different rules (e.g. an empty number became 0, NaN or was
left unchanged; dates were parsed with or without an explicit format).
"""

from __future__ import annotations

import datetime as dt
import math
import numbers
import re
from typing import Any

# Union of the missing markers used in the MATLAB code
# (validateRowAgainstMapping, sanitizeKpiValue, cell2date_safe, is_invalid).
MISSING_TOKENS = frozenset(
    {"", "na", "n/a", "nan", "nat", "null", "none", "-", "--", "missing",
     "<missing>", "<undefined>", "#n/a"}
)

EXCEL_MAX_SERIAL = 2958465  # 31/12/9999
MATLAB_DATENUM_OFFSET = 693960  # datenum(d) - Excel serial(d)
MIN_YEAR, MAX_YEAR = 1900, 2100

MULTI_SEPARATOR = ";"

_NUMBER_RE = re.compile(r"^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$")
_COMMA_DECIMAL_RE = re.compile(r"^[+-]?\d+,\d+$")
_ISO_DATETIME_RE = re.compile(r"^(\d{4}-\d{2}-\d{2})[ T]\d{2}:\d{2}(:\d{2}(\.\d+)?)?$")
_DATE_FORMATS = ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d.%m.%Y", "%Y/%m/%d",
                 "%d %b %Y", "%b %d, %Y")


class ConversionError(ValueError):
    """A value cannot be interpreted with the requested type."""


def plain(value: Any) -> Any:
    """Rich text cells (openpyxl CellRichText) -> str; anything else unchanged."""
    if type(value).__name__ == "CellRichText":
        return str(value)
    return value


def is_missing(value: Any) -> bool:
    value = plain(value)
    if value is None:
        return True
    if isinstance(value, bool):
        return False
    if isinstance(value, numbers.Real):
        try:
            return math.isnan(float(value))
        except (TypeError, ValueError):
            return False
    if isinstance(value, str):
        return value.strip().lower() in MISSING_TOKENS
    # pandas.NA / pandas.NaT without importing pandas at module level
    if type(value).__name__ in {"NAType", "NaTType"}:
        return True
    return False


def to_text(value: Any) -> str | None:
    """Text representation; None for missing values. Integers never get '.0'."""
    value = plain(value)
    if is_missing(value):
        return None
    if isinstance(value, str):
        return value.strip()
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, numbers.Integral):
        return str(int(value))
    if isinstance(value, numbers.Real):
        number = float(value)
        return str(int(number)) if number.is_integer() else format(number, ".15g")
    if isinstance(value, dt.datetime):
        if value.time() == dt.time(0):
            return value.strftime("%d/%m/%Y")
        return value.strftime("%d/%m/%Y %H:%M:%S")
    if isinstance(value, dt.date):
        return value.strftime("%d/%m/%Y")
    text = str(value).strip()
    return text or None


def looks_numeric(text: str) -> bool:
    return bool(_NUMBER_RE.match(text.strip()))


def to_number(value: Any) -> float | None:
    """Float or None. Text must use '.' as decimal separator: '1,5' is rejected
    (MATLAB str2double turned it into 15 without any warning)."""
    value = plain(value)
    if is_missing(value):
        return None
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, numbers.Real):
        number = float(value)
        if math.isinf(number):
            raise ConversionError("valore infinito")
        return number
    if isinstance(value, (dt.date, dt.datetime, dt.time)):
        raise ConversionError("è una data, non un numero")
    text = str(value).strip().replace(" ", "").replace(" ", "")
    if _NUMBER_RE.match(text):
        return float(text)
    if _COMMA_DECIMAL_RE.match(text):
        raise ConversionError(f"'{text}': usa il punto come separatore decimale")
    raise ConversionError(f"'{str(value).strip()}' non è un numero")


def to_int(value: Any) -> int | None:
    number = to_number(value)
    if number is None:
        return None
    if not float(number).is_integer():
        raise ConversionError(f"{number:g} non è un numero intero")
    return int(number)


def excel_serial_to_date(serial: float) -> dt.date:
    from openpyxl.utils.datetime import from_excel

    if not 1 <= serial <= EXCEL_MAX_SERIAL:
        raise ConversionError(f"{serial:g} non è un numero di serie di data Excel")
    return from_excel(serial).date()


def _check_year(day: dt.date, original: Any) -> dt.date:
    if MIN_YEAR <= day.year <= MAX_YEAR:
        return day
    hint = ""
    number = original if isinstance(original, numbers.Real) else None
    if isinstance(original, (dt.date, dt.datetime)):
        # the cell was a date: rebuild its serial to test the MATLAB datenum bug
        base = dt.date(1899, 12, 30)
        number = (day - base).days
    if number is not None:
        try:
            fixed = excel_serial_to_date(float(number) - MATLAB_DATENUM_OFFSET)
            if MIN_YEAR <= fixed.year <= MAX_YEAR:
                hint = f" (probabile datenum MATLAB: la data corretta sarebbe {fixed:%d/%m/%Y})"
        except ConversionError:
            pass
    raise ConversionError(f"data fuori intervallo {MIN_YEAR}-{MAX_YEAR}: {day:%d/%m/%Y}{hint}")


def to_date(value: Any) -> tuple[dt.date | None, str | None]:
    """(date or None, warning or None). Raises ConversionError when not a date.

    Accepted: date/datetime cells, Excel serial numbers, text as dd/mm/yyyy
    (main format), yyyy-mm-dd, yyyymmdd and a few unambiguous variants.
    mm/dd/yyyy is used only when dd/mm/yyyy is impossible, with a warning.
    """
    value = plain(value)
    if is_missing(value):
        return None, None
    if isinstance(value, dt.datetime):
        return _check_year(value.date(), value), None
    if isinstance(value, dt.date):
        return _check_year(value, value), None
    if isinstance(value, bool):
        raise ConversionError("un valore vero/falso non è una data")
    if isinstance(value, numbers.Real):
        return _check_year(excel_serial_to_date(float(value)), value), None

    text = str(value).strip()
    match = _ISO_DATETIME_RE.match(text)
    if match:
        text = match.group(1)
    if text.isdigit():
        if len(text) == 8:
            try:
                return _check_year(dt.datetime.strptime(text, "%Y%m%d").date(), text), None
            except ValueError as exc:
                raise ConversionError(f"'{text}' non è una data AAAAMMGG valida") from exc
        return _check_year(excel_serial_to_date(float(text)), float(text)), None
    for fmt in _DATE_FORMATS:
        try:
            return _check_year(dt.datetime.strptime(text, fmt).date(), text), None
        except ValueError:
            continue
    try:
        day = dt.datetime.strptime(text, "%m/%d/%Y").date()
    except ValueError:
        raise ConversionError(f"'{text}' non è una data valida (atteso gg/mm/aaaa)") from None
    return _check_year(day, text), f"'{text}' interpretata come mm/gg/aaaa ({day:%d/%m/%Y})"


def split_multi(value: Any) -> list[str]:
    """'Brake;Susp' -> ['Brake', 'Susp'] (also accepts ',' as separator)."""
    text = to_text(value)
    if not text:
        return []
    tokens: list[str] = []
    for token in re.split(r"[;,]", text):
        token = token.strip()
        if token and token not in tokens:
            tokens.append(token)
    return tokens


def coerce(field_type: str | None, value: Any) -> tuple[Any, str | None]:
    """Convert a raw value to the Python type used for a schema field type.

    Returns (value, warning). Raises ConversionError for invalid values.
    """
    if field_type in ("text", "choice"):
        return to_text(value), None
    if field_type == "multichoice":
        tokens = split_multi(value)
        return (MULTI_SEPARATOR.join(tokens) if tokens else None), None
    if field_type in ("int", "id"):
        return to_int(value), None
    if field_type == "number":
        return to_number(value), None
    if field_type == "date":
        return to_date(value)
    return native(value), None


def native(value: Any) -> Any:
    """Value of a column with unknown type: keep numbers/dates, clean text."""
    value = plain(value)
    if is_missing(value):
        return None
    if isinstance(value, str):
        return value.strip()
    if isinstance(value, dt.datetime) and value.time() == dt.time(0):
        return value.date()
    if isinstance(value, numbers.Integral) and not isinstance(value, bool):
        return int(value)
    if isinstance(value, numbers.Real) and not isinstance(value, bool):
        return float(value)
    return value


def canonical(value: Any, field_type: str | None) -> str:
    """Comparable text used to detect duplicated rows (MATLAB normalizeColumn)."""
    try:
        if field_type == "date":
            day, _ = to_date(value)
            return day.isoformat() if day else ""
        if field_type in ("number", "int", "id"):
            number = to_number(value)
            return "" if number is None else format(number, ".12g")
        if field_type == "multichoice":
            return MULTI_SEPARATOR.join(sorted(split_multi(value)))
    except ConversionError:
        pass
    value = plain(value)
    if isinstance(value, (dt.date, dt.datetime)):
        day = value.date() if isinstance(value, dt.datetime) else value
        return day.isoformat()
    return to_text(value) or ""


def normalize_name(name: Any) -> str:
    """'Sub Cluster' / 'sub_cluster' -> 'subcluster' (header matching fallback)."""
    return re.sub(r"[^0-9a-z]", "", str(name).lower())
