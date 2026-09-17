"""Common visual language for generated CDE Excel templates."""
from __future__ import annotations

from copy import copy
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

ORANGE = "F7921E"
DARK = "2F2F2F"
LIGHT = "FFF4E8"
BORDER = "D9D9D9"
MUTED = "666666"


def apply_excel_style(workbook) -> None:
    """Apply visual properties without changing workbook structure or values."""
    thin = Side(style="thin", color=BORDER)
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for ws in workbook.worksheets:
        ws.sheet_view.showGridLines = False
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is None:
                    continue
                font = copy(cell.font)
                font.name = "Segoe UI"
                font.sz = 10
                cell.font = font
                # Assign a concrete Border (StyleProxy cannot be registered by
                # openpyxl when copied directly from another cell).
                cell.border = border
                alignment = copy(cell.alignment)
                alignment.vertical = "top"
                alignment.wrap_text = True
                cell.alignment = alignment
        first_values = [cell.value for cell in ws[1] if cell.value is not None]
        # A compact sheet (including Project Point roles) starts with headers.
        # A title-only first row has one value and is followed by the table.
        header_row = 1 if len(first_values) >= 2 else None
        if header_row is None:
            for row_idx in range(2, min(ws.max_row, 8) + 1):
                if sum(cell.value is not None for cell in ws[row_idx]) >= 2:
                    header_row = row_idx
                    break
        if len(first_values) == 1:
            for cell in ws[1]:
                if cell.value is not None:
                    cell.font = Font(name="Segoe UI", sz=14, bold=True, color="FFFFFF")
                    cell.fill = PatternFill("solid", fgColor=DARK)
                    cell.alignment = Alignment(vertical="center", wrap_text=True)
            ws.row_dimensions[1].height = 26
        if header_row is not None:
            for cell in ws[header_row]:
                if cell.value is not None:
                    # Required/locked/alternating fills carry meaning; only
                    # unfilled headers receive the common orange fill.
                    fill = copy(cell.fill)
                    fg = getattr(fill.fgColor, "rgb", "") or ""
                    if not fg.upper().endswith("FFF2CC"):
                        fill = PatternFill("solid", fgColor=ORANGE)
                    cell.fill = fill
                    text_color = "2F2F2F" if fg.upper().endswith("FFF2CC") else "FFFFFF"
                    cell.font = Font(name="Segoe UI", sz=10, bold=True, color=text_color)
                    cell.border = border
                    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            ws.row_dimensions[header_row].height = 30
