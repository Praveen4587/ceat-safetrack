"""Department workbook export helpers for SafeTrack."""
import io, re, shutil
from pathlib import Path
from datetime import datetime
from copy import copy
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent
TEMPLATE_DIR = ROOT / "templates" / "department_exports"

DEPARTMENTS = [
    {"key":"planning","label":"Planning","file":"16. Planning - UC-UA & Near Miss Reporting- FY 26-27.xlsx","aliases":["PLANNING","PRODUCTION PLANNING"]},
    {"key":"division_a","label":"Division A","file":"1.Div A. UC-UA-Reporting- FY 26-27.xlsx","aliases":["DIVISION - A","DIVISION A","DIV A"]},
    {"key":"general","label":"General","file":"2. General - UC-UA & Near Miss Reporting- FY 26-27.xlsx","aliases":["GENERAL"]},
    {"key":"tb","label":"TB - Bead & Sq Bias","file":"4. TB - Bead & Sq Bias UC-UA-Reporting- FY 26-27.xlsx","aliases":["TB","BEAD","SQUARE BIAS"]},
    {"key":"bias_curing","label":"Bias Curing","file":"5.Bias curing - UC-UA & Near Miss Reporting- FY 26-27.xlsx","aliases":["BIAS CURING"]},
    {"key":"tbr","label":"TBR","file":"6.TBR - UC-UA & Near Miss Reporting- FY 26-27.xlsx","aliases":["TBR"]},
    {"key":"tw","label":"TW","file":"7.TW - UC-UA & Near Miss Reporting- FY 26-27.xlsx","aliases":["TW"]},
    {"key":"radial","label":"Radial","file":"8.Radial- UC-UA & Near Miss Reporting- FY 26-27.xlsx","aliases":["RADIAL"]},
    {"key":"engineering","label":"Engineering","file":"9.Engineering- UC-UA & Near Miss Reporting- FY 26-27.xlsx","aliases":["ENGINEERING"]},
    {"key":"lab","label":"LAB","file":"10. LAB- UC-UA & Near Miss Reporting- FY 26-27.xlsx","aliases":["LAB","LABORATORY"]},
    {"key":"technical","label":"Technical","file":"11. Technical - UC-UA & Near Miss Reporting- FY 26-27.xlsx","aliases":["TECHNICAL"]},
    {"key":"qa","label":"QA","file":"12. QA - UC-UA & Near Miss Reporting- FY 26-27.xlsx","aliases":["QA","QUALITY ASSURANCE"]},
    {"key":"hr","label":"HR","file":"13.HR - UC-UA & Near Miss Reporting- FY 26-27.xlsx","aliases":["HR","HUMAN RESOURCES"]},
    {"key":"stores","label":"Stores","file":"14. Stores - UC-UA & Near Miss Reporting- FY 26-27.xlsx","aliases":["STORES","STORE"]},
    {"key":"warehouse","label":"Warehouse (WH)","file":"15. WH - UC-UA & Near Miss Reporting- FY 26-27.xlsx","aliases":["WH","WAREHOUSE"]},
]

def norm(value):
    return re.sub(r"[^a-z0-9]+", "", str(value or "").lower())

def division_access(dept, division):
    text = str(division or "").upper()
    return any(alias in text for alias in dept["aliases"])

def division_matches(dept, division, process="", location=""):
    text = str(division or "").upper()
    if not division_access(dept, division):
        return False
    # The imported roster combines Bias Curing and TBR. Keep manager access to
    # both templates, then split individual reports by their process/location.
    if "BIAS CURING" in text and "TBR" in text:
        marker = (str(process) + " " + str(location)).upper()
        if dept["key"] == "tbr":
            return "TBR" in marker
        if dept["key"] == "bias_curing":
            return "TBR" not in marker
    return True

def _header_row(ws, near_miss=False):
    for row in range(1, min(ws.max_row, 20) + 1):
        vals = [norm(ws.cell(row, col).value) for col in range(1, min(ws.max_column, 25) + 1)]
        if near_miss and any("descriptionofthenearmiss" in v for v in vals):
            return row
        if not near_miss and any(v in ("srlno", "srno", "sno", "serialno") for v in vals):
            if any("dateofreporting" in v for v in vals):
                return row
    return None

def _last_data_row(ws, header_row):
    if header_row is None:
        return 1
    max_col = min(ws.max_column, 20)
    rows = []
    for (r, c), cell in ws._cells.items():
        if isinstance(r, int) and isinstance(c, int) and r > header_row and c <= max_col and cell.value is not None:
            rows.append(r)
    return max(rows, default=header_row)

def _copy_row_style(ws, source_row, target_row, max_col):
    if source_row < 1:
        return
    for col in range(1, max_col + 1):
        src = ws.cell(source_row, col)
        dst = ws.cell(target_row, col)
        if src.has_style:
            dst._style = copy(src._style)
        if src.number_format:
            dst.number_format = src.number_format
        if src.alignment:
            dst.alignment = copy(src.alignment)
        if src.protection:
            dst.protection = copy(src.protection)
    if source_row in ws.row_dimensions:
        ws.row_dimensions[target_row].height = ws.row_dimensions[source_row].height

def _write_mapped_row(ws, header_row, row_num, values, near_miss=False):
    headers = {}
    for col in range(1, min(ws.max_column, 30) + 1):
        value = norm(ws.cell(header_row, col).value)
        if value:
            headers[value] = col

    def find_col(*needles):
        for header, col in headers.items():
            if any(n in header for n in needles):
                return col
        return None

    mapping = {}
    if near_miss:
        mapping = {
            find_col("srlno", "serialno"): values.get("serial"),
            find_col("date"): values.get("date"),
            find_col("time"): values.get("time"),
            find_col("reporter"): values.get("reporter"),
            find_col("employe e pf", "epfno"): values.get("epf"),
            find_col("location"): values.get("location"),
            find_col("descriptionofthenearmiss"): values.get("description"),
            find_col("rootcause"): values.get("root_cause"),
            find_col("preventiveandcorrectiveaction"): values.get("action"),
            find_col("status"): values.get("status"),
        }
    else:
        mapping = {
            find_col("srlno", "srno", "sno", "serialno"): values.get("serial"),
            find_col("dateofreporting"): values.get("date"),
            find_col("reportedby"): values.get("reporter"),
            find_col("employeeepfno", "epfno"): values.get("epf"),
            find_col("location"): values.get("location"),
            find_col("process"): values.get("process"),
            find_col("describetheunsafecondition", "unsafeconditionact"): values.get("description"),
            find_col("ucua", "type", "uc"): values.get("kind"),
            find_col("personresponsibleforrectification"): values.get("assignee"),
            find_col("planneddate", "plnneddate"): values.get("deadline"),
            find_col("actualdateofclosure", "actualdate"): values.get("closed"),
            find_col("statuscompletedpending", "status"): values.get("status"),
            find_col("remark"): values.get("action"),
        }
    for col, value in mapping.items():
        if col is not None and value is not None:
            cell = ws.cell(row_num, col)
            if isinstance(value, str):
                cell.value = value[:32000]
            else:
                cell.value = value

def _as_date(value):
    if not value:
        return None
    try:
        return datetime.strptime(str(value)[:10], "%Y-%m-%d")
    except (ValueError, TypeError):
        return str(value)[:10]

def _template_sheet(wb, near_miss=False):
    if near_miss:
        candidates = [ws for ws in wb.worksheets if "nearmiss" in norm(ws.title)]
    else:
        candidates = [ws for ws in wb.worksheets if ("uc" in norm(ws.title) or "fy" in norm(ws.title)) and "summary" not in norm(ws.title) and "summery" not in norm(ws.title)]
    for ws in candidates:
        header = _header_row(ws, near_miss)
        if header:
            return ws, header
    return None, None

def create_department_workbook(dept, reports):
    if not dept.get("file"):
        raise FileNotFoundError("No department workbook template is configured for this division.")
    path = TEMPLATE_DIR / dept["file"]
    if not path.exists():
        raise FileNotFoundError("The department Excel template is missing: " + dept["file"])
    wb = load_workbook(path)
    # All supplied workbooks are treated strictly as blank layout templates.
    # Remove legacy/example values from staff and summary tabs as well; keep
    # formulas and first-row labels so the workbook structure remains usable.
    report_sheets = {id(ws) for ws, _ in (_template_sheet(wb, False), _template_sheet(wb, True)) if ws is not None}
    for other_ws in wb.worksheets:
        if id(other_ws) in report_sheets:
            continue
        for cell in list(other_ws._cells.values()):
            if cell.row > 1 and cell.data_type != "f" and cell.value is not None:
                cell.value = None
    grouped = {"uc": [], "near": []}
    for r in reports:
        dt = _as_date(r.get("created"))
        created = str(r.get("created") or "")
        time_value = None
        try:
            time_value = datetime.strptime(created, "%Y-%m-%d %H:%M:%S").strftime("%H:%M")
        except ValueError:
            pass
        item = {
            "date": dt, "time": time_value, "reporter": r.get("reporter_name"),
            "epf": r.get("reporter_epf") or r.get("reporter_code"),
            "location": r.get("location"), "process": r.get("process"),
            "description": r.get("description"), "kind": r.get("kind"),
            "assignee": r.get("assigned_name"), "deadline": _as_date(r.get("deadline")),
            "closed": _as_date(r.get("closed")),
            "status": "Completed" if r.get("status") == "Closed" else "Pending",
            "action": r.get("latest_action") or "", "root_cause": "",
            "serial": None,
        }
        grouped["near" if r.get("kind") == "Near Miss" else "uc"].append(item)

    for key, is_near in (("uc", False), ("near", True)):
        ws, header_row = _template_sheet(wb, is_near)
        if ws is None:
            # Some legacy department workbooks do not include one of the report sheets.
            continue
        # The workbook is a FORMAT TEMPLATE ONLY. Remove every pre-existing
        # record from the selected report sheet before inserting app database rows.
        # Keep headings, merged cells, widths, formulas above the header and styling.
        source_style_row = header_row + 1
        # Some legacy templates use an extra sub-header row (Location / Process / UA).
        subheaders = [norm(ws.cell(source_style_row, col).value) for col in range(1, min(ws.max_column, 20) + 1)]
        if any(v in ("location", "process", "ua", "column1") for v in subheaders):
            source_style_row += 1
        first_data_row = source_style_row + 1 if source_style_row > header_row + 1 else header_row + 1
        # Preserve any secondary header row; clear values only from the data area.
        for cell in list(ws._cells.values()):
            if cell.row >= first_data_row and cell.value is not None:
                cell.value = None
        last_row = first_data_row - 1
        for offset, item in enumerate(grouped[key], 1):
            row_num = last_row + offset
            _copy_row_style(ws, source_style_row, row_num, min(ws.max_column, 30))
            item["serial"] = row_num - header_row
            _write_mapped_row(ws, header_row, row_num, item, is_near)

    try:
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
    except Exception:
        pass
    out = io.BytesIO()
    wb.save(out)
    wb.close()
    out.seek(0)
    return out.getvalue()
