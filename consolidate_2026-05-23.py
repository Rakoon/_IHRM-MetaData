"""Consolidate IHRM_DB_Metadata_Master_2026-05-21 + 2026-05-22 into 2026-05-23.

Strategy:
  - 05-21 (190 datasets, includes Freq + Citing Articles) is the spine.
  - 05-22 (125 datasets, adds Accessibility / How-to-access / Category)
    contributes three new columns, joined on dataset ID.
  - 4 datasets in 05-22 are new (IDs 191-194) and are appended.
  - ID 32 in 05-22 has clearly corrupted Accessibility / How-to-access cells
    (URL blob + text from an unrelated dataset). Those two fields are blanked
    for that row; its Category is preserved.

Output: a freshly-formatted workbook with two sheets (Metadata + Codebook).
"""

from __future__ import annotations

import warnings
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

ROOT = Path(__file__).parent
SRC_21 = ROOT / "IHRM_DB_Metadata_Master_2026-05-21.xlsx"
SRC_22 = ROOT / "IHRM_DB_Metadata_Master_2026-05-22.xlsx"
OUT = ROOT / "IHRM_DB_Metadata_Master_2026-05-23.xlsx"

# Final column order for the Metadata sheet
COLUMNS = [
    "#",
    "Name",
    "Owner",
    "Link",
    "Description",
    "Category",
    "Accessibility",
    "How to access data",
    "Public",
    "Years",
    "Countries",
    "Freq.",
    "Citing Articles",
]

# Codebook entries. Existing entries copied from 05-21; new ones appended.
CODEBOOK = [
    ("Column", "Description"),
    ("#", "Article ID number"),
    ("Name", "What is the dataset called?"),
    ("Owner", "Who owns the data"),
    ("Link", "URLs of dataset owner / descriptions"),
    ("Description", "Dataset description"),
    (
        "Category",
        "Topical / subject tags (semicolon-separated) summarizing what the dataset covers. "
        "Added 2026-05-23 from the Phase-2 accessibility review.",
    ),
    (
        "Accessibility",
        "Data-access classification, one of: "
        "Publicly Accessible, Collaborative Access, Restricted Access, Not Specified. "
        "Added 2026-05-23.",
    ),
    (
        "How to access data",
        "Step-by-step instructions (and URLs) for obtaining the dataset, including any "
        "registration, fees, agreements, or collaborator requirements. Added 2026-05-23.",
    ),
    (
        "Public",
        "a. Post links to the website describing the dataset, information collected, "
        "questionnaire items + wordings, codebooks, access instructions, etc.\n"
        "b. List multiple if there are many. Make a new bullet point for each one.\n"
        "c. IF the documentation is not displayed on a website but instead is a "
        "downloadable file (e.g., PDF, Word), then please put the file in a "
        "dataset-specific folder.",
    ),
    (
        "Years",
        "Which years are the data available?\n"
        "a. List all years that are part of the dataset, separated by comma "
        "(e.g., 2004, 2008, 2012, 2016).\n"
        "b. IF the dataset has a regular data collection period (e.g., annual or "
        "semiannual every 6 months, etc.), then please describe the period.",
    ),
    (
        "Countries",
        "If there are different countries each year, then type in 'different countries "
        "each year', then provide a link to the source information. If the information "
        "is stored in a PDF / Word document that you have downloaded, then please type "
        "'see dataset folder'.",
    ),
    (
        "Freq.",
        "Frequency of papers mentioning this dataset in our corpus of ~13,587 "
        "international HRM journal articles.",
    ),
    (
        "Citing Articles",
        "Comma-separated 8-character article IDs referring to articles listed in "
        "article_list.csv.",
    ),
]


def to_int_id(v):
    if v is None or str(v).strip() == "":
        return None
    try:
        return int(float(v))
    except (TypeError, ValueError):
        return None


def clean_cell(v):
    """Trim whitespace; collapse Excel #N/A strings to empty; preserve numbers."""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return v
    s = str(v).strip()
    if s.lower() in {"#n/a", "n/a", "na"}:
        return None
    return s or None


def load_22() -> dict[int, dict]:
    """ID -> {accessibility, how_to_access, category, plus full row for new IDs}."""
    wb = openpyxl.load_workbook(SRC_22, data_only=True)
    ws = wb["Pilot"]
    rows = list(ws.iter_rows(values_only=True))
    out: dict[int, dict] = {}
    for r in rows[1:]:
        i = to_int_id(r[0])
        if i is None:
            continue
        out[i] = {
            "name": clean_cell(r[1]),
            "link": clean_cell(r[2]),
            "owner": clean_cell(r[3]),
            "accessibility": clean_cell(r[4]),
            "how_to_access": clean_cell(r[5]),
            "description": clean_cell(r[6]),
            "category": clean_cell(r[7]),
            "country": clean_cell(r[8]),
            "years": clean_cell(r[9]),
        }
    # ---- Clean rows whose Accessibility / How-to-access cells were entered
    # ---- against the wrong dataset in 05-22. Flagged after spot-checking.
    if 32 in out:  # Registro de Consejeros: Accessibility = URL blob; HTA = ESS 2022 text
        out[32]["accessibility"] = "Not Specified"
        out[32]["how_to_access"] = None
    if 191 in out:  # Korea Youth Panel: HTA contains UK Gov / UK Data Archive text
        out[191]["how_to_access"] = None
    return out


def load_21() -> list[dict]:
    wb = openpyxl.load_workbook(SRC_21, data_only=True)
    ws = wb["Metadata"]
    rows = list(ws.iter_rows(values_only=True))
    header = rows[0]
    col = {h: i for i, h in enumerate(header) if h}
    out: list[dict] = []
    for r in rows[1:]:
        if r is None or all(c is None or str(c).strip() == "" for c in r):
            continue
        i = to_int_id(r[col["#"]])
        if i is None:
            continue
        out.append({
            "id": i,
            "name": clean_cell(r[col["Name"]]),
            "owner": clean_cell(r[col["Owner"]]),
            "link": clean_cell(r[col["Link"]]),
            "description": clean_cell(r[col["Description"]]),
            "public": clean_cell(r[col["Public"]]),
            "years": clean_cell(r[col["Years"]]),
            "countries": clean_cell(r[col["Countries"]]),
            "freq": r[col["Freq."]] if isinstance(r[col["Freq."]], (int, float)) else 0,
            "citing": clean_cell(r[col["Citing Articles"]]),
        })
    return out


def merge() -> list[dict]:
    rows21 = load_21()
    map22 = load_22()
    seen_ids = set()
    merged: list[dict] = []
    for d in rows21:
        seen_ids.add(d["id"])
        new = map22.get(d["id"], {})
        merged.append({
            **d,
            "category": new.get("category"),
            "accessibility": new.get("accessibility"),
            "how_to_access": new.get("how_to_access"),
        })
    # Append 05-22-only datasets (IDs 191-194 in practice)
    for i in sorted(map22):
        if i in seen_ids:
            continue
        n = map22[i]
        merged.append({
            "id": i,
            "name": n.get("name"),
            "owner": n.get("owner"),
            "link": n.get("link"),
            "description": n.get("description"),
            "public": None,
            "years": n.get("years"),
            "countries": n.get("country"),
            "freq": 0,
            "citing": None,
            "category": n.get("category"),
            "accessibility": n.get("accessibility"),
            "how_to_access": n.get("how_to_access"),
        })
    merged.sort(key=lambda d: d["id"])
    return merged


# ----- Styling -----
HEADER_FILL = PatternFill("solid", fgColor="1E407C")  # Penn State navy
HEADER_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
HEADER_ALIGN = Alignment(horizontal="left", vertical="center", wrap_text=True)
BODY_FONT = Font(name="Calibri", size=11)
BODY_ALIGN_TOP = Alignment(vertical="top", wrap_text=True)
BODY_ALIGN_TOP_RIGHT = Alignment(vertical="top", horizontal="right")
THIN_RULE = Side(style="thin", color="D5DBE4")
BORDER = Border(bottom=THIN_RULE)

COL_WIDTHS = {
    "#": 6,
    "Name": 38,
    "Owner": 32,
    "Link": 40,
    "Description": 60,
    "Category": 28,
    "Accessibility": 20,
    "How to access data": 48,
    "Public": 40,
    "Years": 22,
    "Countries": 24,
    "Freq.": 7,
    "Citing Articles": 60,
}


def write_metadata(ws, rows: list[dict]) -> None:
    ws.append(COLUMNS)
    for cell in ws[1]:
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = HEADER_ALIGN
    ws.row_dimensions[1].height = 34
    for c, name in enumerate(COLUMNS, start=1):
        ws.column_dimensions[get_column_letter(c)].width = COL_WIDTHS[name]

    KEY_BY_HEADER = {
        "#": "id",
        "Name": "name",
        "Owner": "owner",
        "Link": "link",
        "Description": "description",
        "Category": "category",
        "Accessibility": "accessibility",
        "How to access data": "how_to_access",
        "Public": "public",
        "Years": "years",
        "Countries": "countries",
        "Freq.": "freq",
        "Citing Articles": "citing",
    }
    for d in rows:
        ws.append([d.get(KEY_BY_HEADER[h]) for h in COLUMNS])

    # Body styling
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=len(COLUMNS)):
        for cell in row:
            cell.font = BODY_FONT
            cell.alignment = BODY_ALIGN_TOP
            cell.border = BORDER
        # right-align numeric columns
        row[0].alignment = BODY_ALIGN_TOP_RIGHT  # #
        row[COLUMNS.index("Freq.")].alignment = BODY_ALIGN_TOP_RIGHT

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(COLUMNS))}{ws.max_row}"


def write_codebook(ws) -> None:
    for r in CODEBOOK:
        ws.append(r)
    for cell in ws[1]:
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = HEADER_ALIGN
    ws.row_dimensions[1].height = 24
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 90
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=2):
        for cell in row:
            cell.font = BODY_FONT
            cell.alignment = BODY_ALIGN_TOP
            cell.border = BORDER
    ws.freeze_panes = "A2"


def main() -> None:
    rows = merge()
    wb = openpyxl.Workbook()
    ws_meta = wb.active
    ws_meta.title = "Metadata"
    write_metadata(ws_meta, rows)
    ws_cb = wb.create_sheet("Codebook")
    write_codebook(ws_cb)
    wb.save(OUT)

    # Brief audit summary
    have_cat = sum(1 for d in rows if d.get("category"))
    have_acc = sum(1 for d in rows if d.get("accessibility"))
    have_hta = sum(1 for d in rows if d.get("how_to_access"))
    new_rows = [d for d in rows if d["id"] > 190]
    print(f"Wrote {OUT.name}: {len(rows)} datasets")
    print(f"  with Category:           {have_cat}")
    print(f"  with Accessibility:      {have_acc}")
    print(f"  with How to access data: {have_hta}")
    print(f"  appended from 05-22 (id > 190): {len(new_rows)}")
    for d in new_rows:
        print(f"    id={d['id']}: {d['name']}")


if __name__ == "__main__":
    main()
