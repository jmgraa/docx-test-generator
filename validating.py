import os
from datetime import datetime


def validate_input(file_path, date, copy_number, output_dir, margin_h, margin_v):
    return [
        _validate_date(date),
        _validate_copy_number(copy_number),
        _validate_output_dir(output_dir),
        _is_docx_file(file_path),
        _validate_margin(margin_h),
        _validate_margin(margin_v),
    ]


def _validate_date(date_string):
    date_string = date_string.strip()

    try:
        datetime.strptime(date_string, "%Y-%m-%d")
        return True
    except ValueError:
        return "Niepoprawna data lub jej format (RRRR-MM-DD)"


def _validate_copy_number(copy_number):
    copy_number = copy_number.strip()

    if not copy_number.isnumeric() or int(copy_number) <= 0:
        return "Niewłaściwa liczba kopii"

    return True


def _validate_output_dir(output_dir):
    output_dir = output_dir.strip()

    if os.path.isdir(output_dir):
        return True

    return "Niewłaściwa ścieżka do folderu wyjściowego"


def _is_docx_file(path):
    path = path.strip()

    if os.path.isfile(path) and path.lower().endswith(".docx"):
        return True

    return "Niepoprawna ścieżka do pliku Word"


def _validate_margin(value):
    s = value.strip()

    try:
        f = float(s)
        if f < 0:
            return "Niewłaściwy margines (ujemna wartość)"
        return True
    except ValueError:
        return "Niewłaściwy margines (wpisz liczbę)"
