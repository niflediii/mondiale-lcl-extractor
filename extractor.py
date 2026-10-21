"""Extract Mondiale VGL LCL rate tables from a PDF and build the "LCL Rates" workbook."""
import re
from datetime import datetime
from io import BytesIO

import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

SHEET_NAME = "LCL Rates"

HEADERS = [
    "POL", "POD", "Service Details", "Valid From", "Valid To", "Notes",
    "Currency", "Min", "Per M3", "BAF Currency", "BAF Rate", "BAF %", "LSS", "CAF%", "PSS", "EIS",
    "Ratio kg/m3", "Currency", "Rate", "Fuel %", "Min", "Base",
]

DATE_RE = r"\d{2}-[A-Za-z]{3}-\d{2}"
DATE_FMT = "%d-%b-%y"
# Row text layout: <dest> <origin> CUR min perM3 CUR bafRate <cells> <service> validFrom validTo
# (the PDF labels its first column "Destination" and its second "Origin")
CELL_RE = re.compile(r"^(?:-|[\d,.]+%?)$")
BOILERPLATE_START = "Rates are subject to variation"


def _clean(value):
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def _number(value):
    """'45.00' -> 45.0, '-' / '' -> None, anything else stays text."""
    text = _clean(value)
    if text in ("", "-"):
        return None
    try:
        return float(text.replace(",", "").rstrip("%"))
    except ValueError:
        return text


def _date(value):
    text = _clean(value)
    if not text:
        return None
    try:
        return datetime.strptime(text, DATE_FMT)
    except ValueError:
        return text


def _parse_row(text):
    """Parse one rate-table text line into the 21 output values, or None."""
    tokens = text.split()
    # dates are always the last two tokens
    if len(tokens) < 12 or not (re.fullmatch(DATE_RE, tokens[-1]) and re.fullmatch(DATE_RE, tokens[-2])):
        return None
    vto, vfrom = tokens[-1], tokens[-2]
    body = tokens[:-2]
    # locate "CUR min perM3 CUR baf" : two currency codes 3 tokens apart
    for i in range(1, len(body) - 4):
        if (
            re.fullmatch(r"[A-Z]{3}", body[i])
            and re.fullmatch(r"[\d,.]+", body[i + 1])
            and re.fullmatch(r"[\d,.]+", body[i + 2])
            and re.fullmatch(r"[A-Z]{3}", body[i + 3])
            and re.fullmatch(r"[\d,.]+", body[i + 4])
        ):
            break
    else:
        return None
    dest_origin = body[:i]
    if len(dest_origin) < 2:
        return None
    origin, dest = dest_origin[-1], " ".join(dest_origin[:-1])
    cur, vmin, m3, baf_cur, baf = body[i], body[i + 1], body[i + 2], body[i + 3], body[i + 4]
    cells, j = [], i + 5
    # 6 cells: BAF %, LSS, CAF %, PSS, EIS, ratio
    while len(cells) < 6 and j < len(body) and CELL_RE.match(body[j]):
        cells.append(body[j])
        j += 1
    if len(cells) < 6:
        return None
    onc_cur = ""
    if j < len(body) and re.fullmatch(r"[A-Z]{3}", body[j]):
        onc_cur = body[j]
        j += 1
    tail = []  # rate, fuel %, min, base
    while len(tail) < 4 and j < len(body) and CELL_RE.match(body[j]):
        tail.append(body[j])
        j += 1
    if len(tail) < 4:
        return None
    service = " ".join(body[j:])
    return [
        origin, dest, service, _date(vfrom), _date(vto), "",
        cur, _number(vmin), _number(m3), baf_cur, _number(baf),
        _number(cells[0]), _number(cells[1]), _number(cells[2]),
        _number(cells[3]), _number(cells[4]), _number(cells[5]),
        onc_cur, _number(tail[0]), _number(tail[1]), _number(tail[2]), _number(tail[3]),
    ]


def _notes_text(page, notes_top, rows_below):
    """Text of the 'Notes' block that starts at notes_top, up to the boiler-plate footer."""
    bottom = page.height
    for line in page.extract_text_lines():
        if line["top"] > notes_top and line["text"].startswith(BOILERPLATE_START):
            bottom = min(bottom, line["top"] - 1)
            break
    for top in rows_below:
        if top > notes_top:
            bottom = min(bottom, top - 1)
    text = page.crop((0, notes_top - 1, page.width, bottom)).extract_text() or ""
    lines = [_clean(line) for line in text.splitlines() if _clean(line)]
    if lines and lines[0].startswith("Notes"):
        lines[0] = lines[0][len("Notes"):].strip()
    return "\n".join(lines)


def extract_rows(pdf_source):
    """Return rows (lists of 21 output values) in PDF order."""
    output = []
    pending = []  # rows still waiting for their "Notes" block

    with pdfplumber.open(pdf_source) as pdf:
        for page in pdf.pages:
            events = []  # (top, kind, payload)
            lines = page.extract_text_lines()
            for line in lines:
                text = _clean(line["text"])
                if text.startswith("Notes"):
                    events.append((line["top"], "notes", None))
                    continue
                row = _parse_row(text)
                if row:
                    events.append((line["top"], "row", row))
            row_tops = [top for top, kind, _ in events if kind == "row"]

            for top, kind, payload in events:
                if kind == "row":
                    output.append(payload)
                    pending.append(payload)
                else:
                    notes = _notes_text(page, top, row_tops)
                    for row in pending:
                        row[5] = notes
                    pending.clear()

    return output


def build_workbook(rows):
    wb = Workbook()
    ws = wb.active
    ws.title = SHEET_NAME

    # Row 1: group banners, row 2: column headers, data from row 3
    ws.append([None] * len(HEADERS))
    ws.append(HEADERS)
    for row in rows:
        ws.append(row)

    blue = PatternFill("solid", fgColor="1F3864")
    green = PatternFill("solid", fgColor="E2EFDA")
    groups = [
        ("Freight", 7, 9, blue, Font(bold=True, color="FFFFFF")),
        ("BAF", 10, 12, green, Font(bold=True, color="1F3864")),
        ("On Carriage", 17, 22, green, Font(bold=True, color="1F3864")),
    ]
    for label, first, last, fill, font in groups:
        ws.merge_cells(start_row=1, start_column=first, end_row=1, end_column=last)
        cell = ws.cell(1, first, label)
        cell.fill, cell.font = fill, font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        for c in range(first, last + 1):
            ws.cell(1, c).fill = fill

    for cell in ws[2]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = blue
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[2].height = 32

    widths = {"A": 24, "B": 24, "C": 38, "D": 12, "E": 12, "F": 60}
    for idx in range(1, len(HEADERS) + 1):
        letter = get_column_letter(idx)
        ws.column_dimensions[letter].width = widths.get(letter, 12)

    numeric_cols = (8, 9, 11, 12, 13, 14, 15, 16, 17, 19, 20, 21, 22)
    for r in range(3, ws.max_row + 1):
        ws.cell(r, 4).number_format = "dd-mmm-yy"
        ws.cell(r, 5).number_format = "dd-mmm-yy"
        ws.cell(r, 6).alignment = Alignment(wrap_text=False, vertical="top")
        for c in numeric_cols:
            ws.cell(r, c).number_format = "0.00"

    ws.freeze_panes = "C3"
    ws.auto_filter.ref = f"A2:{get_column_letter(len(HEADERS))}{ws.max_row}"
    return wb


def pdf_to_xlsx_bytes(pdf_source):
    rows = extract_rows(pdf_source)
    if not rows:
        raise ValueError("No LCL rate rows were found in this PDF.")
    buffer = BytesIO()
    build_workbook(rows).save(buffer)
    buffer.seek(0)
    return buffer, len(rows)


def output_filename(today=None):
    today = today or datetime.now()
    return f"Mondiale_LCL_{today:%Y-%m-%d}.xlsx"
