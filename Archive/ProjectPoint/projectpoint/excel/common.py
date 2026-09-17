from __future__ import annotations

from typing import Iterable

import pandas as pd

__all__ = ["read_excel_table"]


def _norm(value) -> str:
    if value is None:
        return ""
    return " ".join(str(value).strip().lower().replace("_", " ").split())


def read_excel_table(file_path, sheet_name=0, required_headers: Iterable[str] = (), max_scan_rows: int = 20):
    """Read an Excel table and auto-detect its header row.

    ProjectPoint templates intentionally contain a title/instruction row above the
    real table header. This helper makes generated templates directly importable.
    """
    required = {_norm(value) for value in required_headers if _norm(value)}
    if not required:
        return pd.read_excel(file_path, sheet_name=sheet_name)

    preview = pd.read_excel(file_path, sheet_name=sheet_name, header=None, nrows=max_scan_rows)
    header_index = None
    for idx, row in preview.iterrows():
        values = {_norm(value) for value in row.tolist() if _norm(value)}
        if required.issubset(values):
            header_index = int(idx)
            break

    if header_index is None:
        # Fall back to normal pandas behavior so the caller can produce its own
        # meaningful missing-column validation.
        return pd.read_excel(file_path, sheet_name=sheet_name)
    return pd.read_excel(file_path, sheet_name=sheet_name, header=header_index)
