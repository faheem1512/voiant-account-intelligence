"""
Post-processes an already-written .xlsx (pandas.DataFrame.to_excel output)
into something readable on open: bold header, sized columns, wrapped long
text, frozen header row -- instead of the all-columns-8-characters-wide
default pandas/openpyxl produces.

Call this AFTER df.to_excel(...) has written the file, on every sheet
you want styled. Pure formatting -- never touches cell values.
"""

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

HEADER_FILL = PatternFill(start_color="D9E2F3", end_color="D9E2F3", fill_type="solid")
HEADER_FONT = Font(bold=True)
WRAP_ALIGNMENT = Alignment(wrap_text=True, vertical="top")
PLAIN_ALIGNMENT = Alignment(vertical="top")

MAX_COLUMN_WIDTH = 60
MIN_COLUMN_WIDTH = 10


def style_sheet(ws, max_width=MAX_COLUMN_WIDTH, min_width=MIN_COLUMN_WIDTH, freeze_header=True):
    if ws.max_row == 0 or ws.max_column == 0:
        return

    # Header row
    for cell in ws[1]:
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(wrap_text=True, vertical="center")

    # Per-column: size to content (capped), wrap anything that would have
    # needed a wider column than the cap allows.
    for col_idx in range(1, ws.max_column + 1):
        letter = get_column_letter(col_idx)
        longest = 0
        needs_wrap = False
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            length = len(str(value)) if value is not None else 0
            longest = max(longest, length)

        header_len = len(str(ws.cell(row=1, column=col_idx).value or ""))
        longest = max(longest, header_len)

        width = min(max_width, max(min_width, longest + 2))
        if longest + 2 > max_width:
            needs_wrap = True
        ws.column_dimensions[letter].width = width

        if needs_wrap:
            for row_idx in range(2, ws.max_row + 1):
                ws.cell(row=row_idx, column=col_idx).alignment = WRAP_ALIGNMENT
        else:
            for row_idx in range(2, ws.max_row + 1):
                ws.cell(row=row_idx, column=col_idx).alignment = PLAIN_ALIGNMENT

    if freeze_header:
        ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def style_xlsx(path, sheet_names=None, max_width=MAX_COLUMN_WIDTH, min_width=MIN_COLUMN_WIDTH):
    """sheet_names: list of sheet names to style, or None for every sheet
    in the workbook (skips sheets with no data)."""
    wb = load_workbook(path)
    targets = sheet_names or wb.sheetnames
    for name in targets:
        if name in wb.sheetnames:
            style_sheet(wb[name], max_width=max_width, min_width=min_width)
    wb.save(path)


# Same A/B/C/D palette the in-app results table used to use before it was
# replaced by Excel exports -- keeping it consistent rather than inventing
# a new one.
PRIORITY_BAND_COLORS = {
    "A": "DCFCE7",  # green
    "B": "FEF9C3",  # yellow
    "C": "FFEDD5",  # orange
    "D": "FEE2E2",  # red
}

# priority_requeue.xlsx is PRE-research (the free scanner's output) -- it
# genuinely can't have a Priority Score/Band yet, since that formula needs
# Platform Confidence, which only the paid Account Research step
# determines. This uses the same green/yellow visual language on Scanner
# Strength instead, so the file still reads consistently at a glance
# without pretending to be the scored, post-research output.
SCANNER_STRENGTH_COLORS = {
    "Strong": PRIORITY_BAND_COLORS["A"],
    "Moderate": PRIORITY_BAND_COLORS["B"],
}


def color_code_column(path, column_name, value_to_hex, sheet_name=None):
    """Fills each data row's cells based on the value in column_name,
    looked up in value_to_hex (e.g. PRIORITY_BAND_COLORS). Call this AFTER
    style_xlsx() -- both just load/modify/save, order between them doesn't
    matter, but keep it consistent. Rows whose value isn't in the map are
    left unstyled, not an error."""
    wb = load_workbook(path)
    ws = wb[sheet_name] if sheet_name else wb.active

    header = [c.value for c in ws[1]]
    if column_name not in header:
        wb.save(path)
        return
    col_idx = header.index(column_name) + 1

    for row_idx in range(2, ws.max_row + 1):
        value = ws.cell(row=row_idx, column=col_idx).value
        hex_color = value_to_hex.get(str(value).strip()) if value is not None else None
        if hex_color:
            fill = PatternFill(start_color=hex_color, end_color=hex_color, fill_type="solid")
            for col in range(1, ws.max_column + 1):
                ws.cell(row=row_idx, column=col).fill = fill

    wb.save(path)


def color_code_rows_by_list(path, hex_per_row, sheet_name=None):
    """Colors each data row using hex_per_row[i] (aligned by position, row
    2 = hex_per_row[0], row 3 = hex_per_row[1], ...) rather than deriving
    the color from a column's own displayed value. Needed for
    priority_requeue.xlsx: Priority Band is genuinely blank pre-research,
    so color_code_column() has nothing to key off, but Scanner Strength
    (which the row doesn't display as its own visible column in every
    layout) still gives a real, meaningful color. None entries are left
    unstyled, not an error."""
    wb = load_workbook(path)
    ws = wb[sheet_name] if sheet_name else wb.active

    for offset, row_idx in enumerate(range(2, ws.max_row + 1)):
        if offset >= len(hex_per_row):
            break
        hex_color = hex_per_row[offset]
        if hex_color:
            fill = PatternFill(start_color=hex_color, end_color=hex_color, fill_type="solid")
            for col in range(1, ws.max_column + 1):
                ws.cell(row=row_idx, column=col).fill = fill

    wb.save(path)
