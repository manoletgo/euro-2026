"""Build the Euro 2026 itinerary as both .xlsx and .pdf from a single data source.

Run from the repo root:
    python scripts/build_itinerary.py

Outputs (overwritten on each run):
    output/euro_2026_itinerary.xlsx
    output/euro_2026_itinerary.pdf

When the source data in travel_info_data/info_notes.txt changes, update the
ROWS list below and rerun. Keep the schema, type categories, and colors in
sync with CLAUDE.md.
"""

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from reportlab.lib import colors
from reportlab.lib.pagesizes import A3, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle


HEADERS = [
    "Date",
    "Day",
    "Time (Depart)",
    "City / Route",
    "Type",
    "Location / Details",
    "Departure",
    "Arrival",
    "Duration",
    "Notes",
]

# Type prefix -> hex fill color. Used by both xlsx and pdf renderers.
TYPE_COLORS = {
    "Flight":          "DDEBF7",
    "Train":           "FFF2CC",
    "Lodging":         "E2EFDA",
    "Activity":        "FCE4D6",
    "Luggage storage": "EDEDED",
}

HEADER_FILL = "305496"

# Column widths (in mm for pdf, scaled for xlsx).
COL_WIDTHS_MM = [22, 12, 22, 32, 22, 60, 50, 50, 18, 70]

ROWS = [
    ["2026-05-17", "Sun", "", "Manila → Istanbul", "Flight",
     "Flight J0085",
     "Ninoy Aquino International Airport (Manila)",
     "Istanbul", "",
     "Departure from NAIA; international flight to Istanbul; in-flight overnight"],
    ["2026-05-18", "Mon", "", "Istanbul → Vienna", "Flight",
     "Flight J1883",
     "Istanbul", "Vienna — 09:15", "",
     "Arrival in Vienna"],
    ["2026-05-18", "Mon", "", "Vienna", "Luggage storage",
     "Bounce Luggage Storage – Vienna Prince Mobile (near Südtiroler Platz Station), Wiednergürtel 48/IV, Vienna 1040",
     "", "", "",
     "Drop luggage on arrival (~9:15 am); pick up later once Airbnb check-in is available"],
    ["2026-05-18", "Mon", "", "Vienna", "Lodging (Airbnb)",
     "Margaretengürtel 6, 1050 Wien, Austria",
     "Check-in: 2026-05-18", "Check-out: 2026-05-22", "4 nights",
     "Google Maps: https://maps.app.goo.gl/bAQYMCvSrB3vRvXn7"],
    ["2026-05-18", "Mon", "Evening", "Vienna", "Activity",
     "Walk around Stephansplatz; visit St. Stephen's Cathedral",
     "", "", "",
     "Evening activity after Airbnb check-in"],
    ["2026-05-20", "Wed", "07:34", "Vienna → Budapest", "Train",
     "RegioJet RJ 1065",
     "Vienna Central Train Station — 07:34",
     "Budapest Déli — 10:14", "2h 40m",
     "Trainline: https://app.trainline.com/QrswMpOb12b"],
    ["2026-05-20", "Wed", "17:45", "Budapest → Vienna", "Train",
     "RegioJet RJ 1068 (same-day return)",
     "Budapest Déli — 17:45",
     "Vienna Central Train Station — 20:27", "2h 42m",
     "Trainline: https://app.trainline.com/MeSYEbQb12b"],
    ["2026-05-22", "Fri", "06:39", "Vienna → Prague", "Train",
     "RegioJet RJ 1030",
     "Vienna Central Train Station — 06:39",
     "Praha hl.n. — 10:56", "4h 17m",
     "Trainline: https://app.trainline.com/K0MINBRb12b"],
    ["2026-05-22", "Fri", "", "Prague", "Lodging (Airbnb)",
     "Španělská 759/4, 120 00 Vinohrady, Czechia",
     "Check-in: 2026-05-22", "Check-out: 2026-05-26", "4 nights",
     "Google Maps: https://maps.app.goo.gl/UcKWdy4NRJWvL3eZA"],
    ["2026-05-23", "Sat", "08:31", "Prague → Dresden", "Train",
     "Deutsche Bahn 178 (day-trip outbound)",
     "Praha hl.n. — 08:31", "Dresden Hbf — 10:50", "2h 19m",
     "Trainline: https://app.trainline.com/KXTU6ZSb12b"],
    ["2026-05-23", "Sat", "19:10", "Dresden → Prague", "Train",
     "Deutsche Bahn 385 (day-trip return)",
     "Dresden Hbf — 19:10", "Praha hl.n. — 21:25", "2h 15m",
     "Trainline: https://app.trainline.com/Uo4dOhUb12b"],
    ["2026-05-26", "Tue", "12:10", "Prague → Stockholm", "Flight",
     "Norwegian D82621 (Norwegian Air Sweden AOC AB); Ticket type: LOWFARE+; Seat 12C; Baggage: 1 checked + 1 underseat + 1 overhead",
     "Prague Václav Havel (PRG), Terminal 2 — 12:10",
     "Stockholm Arlanda (ARN), Terminal 5 — 14:05", "1h 55m",
     "Booking ref: Y9GHBZ; Ticket: 328 2404310661 (Norwegian Air Manolet.pdf)"],
    ["2026-05-26", "Tue", "", "Stockholm", "Lodging (Hotel)",
     "Bob W Stockholm Södermalm",
     "Check-in: 2026-05-26", "Check-out: 2026-05-30", "4 nights",
     "Google Maps: https://maps.app.goo.gl/QzqTafk4EjpTVc577"],
    ["2026-05-30", "Sat", "", "Stockholm → Istanbul", "Flight",
     "Flight J1796",
     "Stockholm Arlanda Airport", "Istanbul", "",
     "Transfer to Stockholm Arlanda; departure from Schengen area; in-flight overnight"],
    ["2026-05-31", "Sun", "", "Istanbul → Manila", "Flight",
     "Flight J0084",
     "Istanbul", "Manila", "",
     "Return flight to Manila"],
]

REPO_ROOT = Path(__file__).resolve().parent.parent
OUTPUT_DIR = REPO_ROOT / "output"
XLSX_PATH = OUTPUT_DIR / "euro_2026_itinerary.xlsx"
PDF_PATH = OUTPUT_DIR / "euro_2026_itinerary.pdf"


def fill_for_type(type_val: str) -> str | None:
    for prefix, hex_color in TYPE_COLORS.items():
        if type_val.startswith(prefix):
            return hex_color
    return None


def build_xlsx(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Itinerary"

    header_font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor=HEADER_FILL)
    body_font = Font(name="Arial", size=10)
    thin = Side(border_style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    wrap = Alignment(wrap_text=True, vertical="top")

    for col_idx, header in enumerate(HEADERS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border

    for r_idx, row in enumerate(ROWS, start=2):
        type_val = row[4]
        hex_color = fill_for_type(type_val)
        fill = PatternFill("solid", fgColor=hex_color) if hex_color else None
        for c_idx, val in enumerate(row, start=1):
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            cell.font = body_font
            if fill:
                cell.fill = fill
            cell.alignment = wrap
            cell.border = border

    # Roughly map mm widths to Excel column units (~1 mm ≈ 0.55 unit).
    for i, mm_width in enumerate(COL_WIDTHS_MM, start=1):
        ws.column_dimensions[get_column_letter(i)].width = max(8, mm_width * 0.85)

    ws.row_dimensions[1].height = 22
    ws.freeze_panes = "A2"

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def build_pdf(path: Path) -> None:
    styles = getSampleStyleSheet()
    cell_style = ParagraphStyle(
        "cell", parent=styles["BodyText"],
        fontName="Helvetica", fontSize=8, leading=10, wordWrap="CJK",
    )
    header_style = ParagraphStyle(
        "header", parent=styles["BodyText"],
        fontName="Helvetica-Bold", fontSize=9, leading=11,
        textColor=colors.white, alignment=1,
    )
    title_style = ParagraphStyle(
        "title", parent=styles["Title"],
        fontName="Helvetica-Bold", fontSize=16, spaceAfter=8,
    )

    def cellp(text: str) -> Paragraph:
        return Paragraph(str(text).replace("\n", "<br/>") if text else "", cell_style)

    table_data = [[Paragraph(h, header_style) for h in HEADERS]]
    table_data.extend([[cellp(c) for c in row] for row in ROWS])

    col_widths = [w * mm for w in COL_WIDTHS_MM]
    table = Table(table_data, colWidths=col_widths, repeatRows=1)

    style = TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(f"#{HEADER_FILL}")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("ALIGN", (0, 0), (-1, 0), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#BFBFBF")),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ])

    for i, row in enumerate(ROWS, start=1):
        hex_color = fill_for_type(row[4])
        if hex_color:
            style.add("BACKGROUND", (0, i), (-1, i), colors.HexColor(f"#{hex_color}"))

    table.setStyle(style)

    path.parent.mkdir(parents=True, exist_ok=True)
    doc = SimpleDocTemplate(
        str(path), pagesize=landscape(A3),
        leftMargin=10 * mm, rightMargin=10 * mm,
        topMargin=12 * mm, bottomMargin=10 * mm,
        title="Euro 2026 Itinerary", author="Manolet",
    )
    doc.build([Paragraph("Euro 2026 Itinerary", title_style), Spacer(1, 4), table])


def main() -> None:
    build_xlsx(XLSX_PATH)
    build_pdf(PDF_PATH)
    print(f"Saved: {XLSX_PATH}")
    print(f"Saved: {PDF_PATH}")
    print(f"Rows: {len(ROWS)}")


if __name__ == "__main__":
    main()
