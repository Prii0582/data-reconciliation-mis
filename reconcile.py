"""
Multi-Source Data Reconciliation & Monthly MIS
==============================================

Reads four source-system extracts that *should* agree about every student,
checks each one, reconciles them on student_id, and produces:

    output/reconciliation_report.xlsx   - discrepancy log + department work-lists
    output/monthly_mis_dashboard.html   - one-page monthly MIS dashboard
    output/monthly_mis_dashboard.pdf    - same, printed by headless Chrome
    docs/dashboard_preview.png          - screenshot used in the README

Run:  python reconcile.py            (defaults to the October 2026 MIS)
      python reconcile.py --as-of 2026-10-31
"""
from __future__ import annotations

import argparse
import base64
import io
import os
import re
import shutil
import subprocess
from datetime import date, timedelta
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
import numpy as np  # noqa: E402
import pandas as pd  # noqa: E402
from jinja2 import Environment, FileSystemLoader, select_autoescape  # noqa: E402
from openpyxl import Workbook  # noqa: E402
from openpyxl.chart import BarChart, Reference  # noqa: E402
from openpyxl.formatting.rule import CellIsRule  # noqa: E402
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side  # noqa: E402
from openpyxl.utils import get_column_letter  # noqa: E402

ROOT = Path(__file__).resolve().parent
DATA, OUT, DOCS, TEMPLATES = ROOT / "data", ROOT / "output", ROOT / "docs", ROOT / "templates"
INSTITUTE = "Eastvale Institute of Applied Skills"  # fictional
DEFAULT_CHROME = r"C:\Program Files\Google\Chrome\Application\chrome.exe"

NAVY, TEAL, LIGHT, AMBER, RED, GREY = "1B2A4A", "0F8B8D", "EEF3F7", "E0A030", "C0392B", "6B7785"

ADMISSIONS, FINANCE, ACADEMICS, INDUSTRY = "Admissions", "Finance", "Academics", "Industry Relations"
DEPARTMENTS = [ADMISSIONS, FINANCE, ACADEMICS, INDUSTRY]

# issue type -> (default owning department, severity, check family)
ISSUES = {
    # single-source validation checks
    "Missing mandatory field":            (ADMISSIONS, "Medium", "Validation"),
    "Invalid code value":                 (ADMISSIONS, "Low",    "Validation"),
    "Invalid email format":               (ADMISSIONS, "Low",    "Validation"),
    "Invalid phone format":               (ADMISSIONS, "Low",    "Validation"),
    "Date of birth out of range":         (ADMISSIONS, "Medium", "Validation"),
    "Duplicate student ID":               (ADMISSIONS, "High",   "Validation"),
    "Invalid student ID format":          (FINANCE,    "Low",    "Validation"),
    "Duplicate receipt":                  (FINANCE,    "High",   "Validation"),
    "Invalid date":                       (FINANCE,    "Medium", "Validation"),
    "Negative fee amount":                (FINANCE,    "Medium", "Validation"),
    "Fee amount out of range":            (FINANCE,    "Medium", "Validation"),
    "Marks out of range":                 (ACADEMICS,  "High",   "Validation"),
    "Result published but marks missing": (ACADEMICS,  "High",   "Validation"),
    "Result inconsistent with marks":     (ACADEMICS,  "High",   "Validation"),
    "OJT end date before start date":     (INDUSTRY,   "Medium", "Validation"),
    "Partner not in master list":         (INDUSTRY,   "Medium", "Validation"),
    "CTC out of range":                   (INDUSTRY,   "Low",    "Validation"),
    # cross-source reconciliation checks
    "Admitted but missing in ERP":        (ACADEMICS,  "High",   "Reconciliation"),
    "Fee paid but not enrolled":          (ACADEMICS,  "High",   "Reconciliation"),
    "Fee paid but enrolment withdrawn":   (FINANCE,    "Medium", "Reconciliation"),
    "In ERP but no admission record":     (ADMISSIONS, "High",   "Reconciliation"),
    "Fee receipt for unknown student":    (FINANCE,    "High",   "Reconciliation"),
    "Placement record for unknown student": (INDUSTRY, "High",   "Reconciliation"),
    "Programme mismatch":                 (ACADEMICS,  "Medium", "Reconciliation"),
    "Programme name spelling variant":    (ADMISSIONS, "Low",    "Reconciliation"),
    "Date of birth mismatch":             (ADMISSIONS, "Medium", "Reconciliation"),
    "Name mismatch":                      (ADMISSIONS, "Medium", "Reconciliation"),
}
SLA_DAYS = {"High": 3, "Medium": 7, "Low": 14}
SEVERITY_ORDER = {"High": 0, "Medium": 1, "Low": 2}

STUDENT_ID_RE = re.compile(r"^ST20\d{2}-\d{4}$")
EMAIL_RE = re.compile(r"^[\w.+-]+@[\w-]+(\.[\w-]+)+$")
PHONE_RE = re.compile(r"^[6-9]\d{9}$")


# ==========================================================================
# Discrepancy log
# ==========================================================================
class IssueLog:
    def __init__(self):
        self.rows: list[dict] = []

    def add(self, student_id, issue_type, detail, source_a, value_a="", source_b="", value_b="", owner=None):
        dept, severity, family = ISSUES[issue_type]
        self.rows.append({
            "student_id": "" if pd.isna(student_id) else str(student_id).strip().upper(),
            "issue_type": issue_type, "check_family": family, "issue_detail": detail,
            "source_a": source_a, "value_a": _txt(value_a),
            "source_b": source_b, "value_b": _txt(value_b),
            "owner_dept": owner or dept, "severity": severity,
        })

    def to_frame(self, as_of: date) -> pd.DataFrame:
        df = pd.DataFrame(self.rows)
        df["_sev"] = df["severity"].map(SEVERITY_ORDER)
        df["_dept"] = df["owner_dept"].map({d: i for i, d in enumerate(DEPARTMENTS)})
        df = df.sort_values(["_dept", "_sev", "issue_type", "student_id"]).drop(columns=["_sev", "_dept"])
        df.insert(0, "log_id", [f"DSC-{i:04d}" for i in range(1, len(df) + 1)])
        df["status"] = "Open"
        df["raised_on"] = as_of
        df["target_close_date"] = df["severity"].map(lambda s: as_of + timedelta(days=SLA_DAYS[s]))
        return df.reset_index(drop=True)


def _txt(v) -> str:
    if v is None or (isinstance(v, float) and np.isnan(v)) or v is pd.NA:
        return "(blank)"
    return str(v)


def blank(s: pd.Series) -> pd.Series:
    return s.isna() | (s.astype("string").str.strip() == "")


def norm_programme(name) -> str:
    """Normalise programme spellings: 'BVoc Retail Mgmt' == 'B.Voc Retail Management'."""
    if pd.isna(name):
        return ""
    s = str(name).lower().replace("&", " and ")
    s = re.sub(r"\bmgmt\b", "management", s)
    s = re.sub(r"\btech\b", "technology", s)
    return re.sub(r"[^a-z0-9]", "", s)


# ==========================================================================
# Load
# ==========================================================================
def load_sources():
    read = lambda f: pd.read_csv(DATA / f, dtype=str, keep_default_na=False).replace({"": None})
    return {
        "adm": read("admissions.csv"),
        "fin": read("finance_fees.csv"),
        "erp": read("academic_erp.csv"),
        "plc": read("industry_partner_placements.csv"),
        "prog": read("programme_master.csv"),
        "partners": read("partner_master.csv"),
    }


# ==========================================================================
# Step 1 - validation checks on each source
# ==========================================================================
def check_mandatory(df, cols, source, owner, log, id_col="student_id"):
    for col in cols:
        for _, r in df[blank(df[col])].iterrows():
            log.add(r.get(id_col), "Missing mandatory field", f"{col} is blank in {source}",
                    f"{source}.{col}", None, owner=owner)


def check_duplicates(df, source, owner, log, name_col):
    dup = df[df.duplicated("student_id", keep=False) & ~blank(df["student_id"])]
    for sid, grp in dup.groupby("student_id"):
        exact = len(grp.drop_duplicates()) == 1
        detail = (f"exact duplicate row ({len(grp)} copies) in {source}" if exact
                  else f"same ID used for {len(grp)} different records in {source}")
        log.add(sid, "Duplicate student ID", detail, f"{source}.student_id",
                " | ".join(grp[name_col].fillna("(blank)")), owner=owner)


def validate_admissions(adm, log):
    check_mandatory(adm, ["student_id", "full_name", "gender", "date_of_birth", "category",
                          "programme_name", "admission_date"], "Admissions", ADMISSIONS, log)
    check_duplicates(adm, "Admissions", ADMISSIONS, log, "full_name")
    for col, allowed in (("gender", {"Female", "Male", "Other"}), ("category", {"GEN", "OBC", "SC", "ST", "EWS"})):
        bad = adm[~blank(adm[col]) & ~adm[col].isin(allowed)]
        for _, r in bad.iterrows():
            log.add(r.student_id, "Invalid code value", f"{col} must be one of {sorted(allowed)}",
                    f"Admissions.{col}", r[col])
    for _, r in adm[~blank(adm.email) & ~adm.email.fillna("").str.match(EMAIL_RE)].iterrows():
        log.add(r.student_id, "Invalid email format", "e-mail address is not valid", "Admissions.email", r.email)
    for _, r in adm[~blank(adm.phone) & ~adm.phone.fillna("").str.match(PHONE_RE)].iterrows():
        log.add(r.student_id, "Invalid phone format", "phone must be 10 digits starting 6-9", "Admissions.phone", r.phone)
    dob = pd.to_datetime(adm.date_of_birth, format="%Y-%m-%d", errors="coerce")
    admitted = pd.to_datetime(adm.admission_date, format="%Y-%m-%d", errors="coerce")
    age = (admitted - dob).dt.days / 365.25
    for _, r in adm[(age < 16) | (age > 45)].iterrows():
        a = (pd.Timestamp(r.admission_date) - pd.Timestamp(r.date_of_birth)).days / 365.25
        log.add(r.student_id, "Date of birth out of range", f"age at admission is {a:.1f} yrs (allowed 16-45)",
                "Admissions.date_of_birth", r.date_of_birth)


def validate_finance(fin, max_fee, log):
    check_mandatory(fin, ["receipt_no", "student_id", "amount_paid", "payment_date"], "Finance", FINANCE, log)
    for _, r in fin[~blank(fin.student_id) & ~fin.student_id.fillna("").str.match(STUDENT_ID_RE)].iterrows():
        log.add(r.student_id, "Invalid student ID format",
                f"receipt {r.receipt_no}: ID has wrong case / stray spaces", "Finance.student_id", repr(r.student_id))
    dup = fin[fin.duplicated("receipt_no", keep=False)]
    for rno, grp in dup.groupby("receipt_no"):
        log.add(grp.student_id.iloc[0], "Duplicate receipt",
                f"receipt {rno} appears {len(grp)} times (amount counted twice)", "Finance.receipt_no",
                f"{rno} x{len(grp)}", "Finance.amount_paid", grp.amount_paid.iloc[0])
    pdate = pd.to_datetime(fin.payment_date, format="%Y-%m-%d", errors="coerce")
    for _, r in fin[pdate.isna() & ~blank(fin.payment_date)].iterrows():
        log.add(r.student_id, "Invalid date", f"receipt {r.receipt_no}: payment date is not a real date",
                "Finance.payment_date", r.payment_date)
    amt = pd.to_numeric(fin.amount_paid, errors="coerce")
    for _, r in fin[amt < 0].iterrows():
        log.add(r.student_id, "Negative fee amount", f"receipt {r.receipt_no}: reversal posted as a receipt",
                "Finance.amount_paid", r.amount_paid)
    for _, r in fin[amt > 2 * max_fee].iterrows():
        log.add(r.student_id, "Fee amount out of range", f"receipt {r.receipt_no}: above 2x the highest annual fee",
                "Finance.amount_paid", r.amount_paid)


def validate_erp(erp, log):
    check_mandatory(erp, ["student_id", "student_name", "programme_name", "enrolment_status"], "ERP", ACADEMICS, log)
    check_duplicates(erp, "ERP", ACADEMICS, log, "student_name")
    marks = pd.to_numeric(erp.marks_obtained, errors="coerce")
    maxm = pd.to_numeric(erp.max_marks, errors="coerce")
    published = erp.exam_status.eq("Result Published")
    seen = set()
    for _, r in erp[published & blank(erp.marks_obtained)].drop_duplicates("student_id").iterrows():
        log.add(r.student_id, "Result published but marks missing", f"{r.exam_session} result is '{r.result}' but marks are blank",
                "ERP.exam_status", r.exam_status, "ERP.marks_obtained", None)
    for i, r in erp[(marks > maxm) | (marks < 0)].iterrows():
        if r.student_id in seen:
            continue
        seen.add(r.student_id)
        log.add(r.student_id, "Marks out of range", f"marks must be between 0 and {r.max_marks}",
                "ERP.marks_obtained", r.marks_obtained, "ERP.max_marks", r.max_marks)
    pct = marks / maxm
    bad = erp[published & pct.notna() & (marks <= maxm) &
              (((pct >= 0.4) & erp.result.eq("Fail")) | ((pct < 0.4) & erp.result.eq("Pass")))]
    for _, r in bad.drop_duplicates("student_id").iterrows():
        log.add(r.student_id, "Result inconsistent with marks", "pass mark is 40%", "ERP.result", r.result,
                "ERP.marks_obtained", r.marks_obtained)


def validate_placements(plc, partners, log):
    check_mandatory(plc, ["record_id", "student_id", "partner_name", "engagement_type", "start_date"],
                    "Industry", INDUSTRY, log)
    approved = set(partners.partner_name.str.strip().str.casefold())
    for _, r in plc[~plc.partner_name.fillna("").str.strip().str.casefold().isin(approved)].iterrows():
        log.add(r.student_id, "Partner not in master list", f"record {r.record_id}: partner has no MoU on file",
                "Industry.partner_name", r.partner_name, "Partner master", "no match")
    start = pd.to_datetime(plc.start_date, format="%Y-%m-%d", errors="coerce")
    end = pd.to_datetime(plc.end_date, format="%Y-%m-%d", errors="coerce")
    for _, r in plc[end < start].iterrows():
        log.add(r.student_id, "OJT end date before start date", f"record {r.record_id}", "Industry.start_date",
                r.start_date, "Industry.end_date", r.end_date)
    ctc = pd.to_numeric(plc.annual_ctc_lpa, errors="coerce")
    for _, r in plc[(ctc <= 0) | (ctc > 50)].iterrows():
        log.add(r.student_id, "CTC out of range", f"record {r.record_id}: CTC must be 0-50 LPA",
                "Industry.annual_ctc_lpa", r.annual_ctc_lpa)


# ==========================================================================
# Step 2 - clean, then reconcile across sources on student_id
# ==========================================================================
def clean(src):
    adm = src["adm"].drop_duplicates().drop_duplicates("student_id", keep="first").copy()
    erp = src["erp"].drop_duplicates().drop_duplicates("student_id", keep="first").copy()

    fin = src["fin"].copy()
    fin["student_id"] = fin.student_id.str.strip().str.upper()
    fin = fin.drop_duplicates("receipt_no", keep="first")
    fin["amount_paid"] = pd.to_numeric(fin.amount_paid, errors="coerce")
    fin["payment_date"] = pd.to_datetime(fin.payment_date, format="%Y-%m-%d", errors="coerce")
    fin["valid_amount"] = fin.amount_paid.where(fin.amount_paid > 0, 0)

    plc = src["plc"].copy()
    plc["student_id"] = plc.student_id.str.strip().str.upper()
    return adm, erp, fin, plc


def reconcile(adm, erp, fin, plc, log):
    A, E = set(adm.student_id), set(erp.student_id)
    paid = fin.groupby("student_id").valid_amount.sum()
    F, P = set(paid.index), set(plc.student_id)

    for sid in sorted(A - E):
        r = adm.loc[adm.student_id == sid].iloc[0]
        if paid.get(sid, 0) > 0:
            log.add(sid, "Fee paid but not enrolled", "fee received but no enrolment in ERP",
                    "Admissions.admission_date", r.admission_date, "Finance.amount_paid", f"{paid[sid]:,.0f}")
        else:
            log.add(sid, "Admitted but missing in ERP", "admission confirmed but student not created in ERP",
                    "Admissions.admission_date", r.admission_date, "ERP", "not found")
    for sid in sorted(E - A):
        r = erp.loc[erp.student_id == sid].iloc[0]
        log.add(sid, "In ERP but no admission record", "student is enrolled but has no admission file",
                "ERP.student_name", r.student_name, "Admissions", "not found")
    for sid in sorted(F - A - E):
        log.add(sid, "Fee receipt for unknown student", "money received against an ID no system recognises",
                "Finance.amount_paid", f"{paid[sid]:,.0f}", "Admissions / ERP", "not found")
    for sid in sorted(P - A - E):
        rec = plc.loc[plc.student_id == sid].iloc[0]
        log.add(sid, "Placement record for unknown student", f"record {rec.record_id}",
                "Industry.partner_name", rec.partner_name, "Admissions / ERP", "not found")

    withdrawn = erp[erp.enrolment_status.eq("Withdrawn")]
    for sid in withdrawn.student_id:
        if paid.get(sid, 0) > 0:
            log.add(sid, "Fee paid but enrolment withdrawn", "check refund / fee adjustment",
                    "ERP.enrolment_status", "Withdrawn", "Finance.amount_paid", f"{paid[sid]:,.0f}")

    both = adm.merge(erp, on="student_id", suffixes=("_adm", "_erp"))
    for _, r in both.iterrows():
        pa, pe = r.programme_name_adm, r.programme_name_erp
        if (pa or "") != (pe or ""):
            if norm_programme(pa) == norm_programme(pe):
                log.add(r.student_id, "Programme name spelling variant", "same programme, different spelling",
                        "Admissions.programme_name", pa, "ERP.programme_name", pe)
            else:
                log.add(r.student_id, "Programme mismatch", "student shows a different programme in each system",
                        "Admissions.programme_name", pa, "ERP.programme_name", pe)
        if pd.notna(r.date_of_birth_adm) and pd.notna(r.date_of_birth_erp) and r.date_of_birth_adm != r.date_of_birth_erp:
            log.add(r.student_id, "Date of birth mismatch", "verify against 10th-standard certificate",
                    "Admissions.date_of_birth", r.date_of_birth_adm, "ERP.date_of_birth", r.date_of_birth_erp)
        if (r.full_name or "").casefold().strip() != (r.student_name or "").casefold().strip():
            log.add(r.student_id, "Name mismatch", "name differs between systems",
                    "Admissions.full_name", r.full_name, "ERP.student_name", r.student_name)
    return A | E | F | P


# ==========================================================================
# Step 3 - Excel report
# ==========================================================================
HDR_FILL = PatternFill("solid", fgColor=NAVY)
HDR_FONT = Font(bold=True, color="FFFFFF", name="Calibri", size=11)
ZEBRA = PatternFill("solid", fgColor="F5F8FA")
THIN = Border(bottom=Side(style="thin", color="D5DEE6"))
SEV_FILL = {"High": PatternFill("solid", fgColor="F8D7D3"), "Medium": PatternFill("solid", fgColor="FBEBC8"),
            "Low": PatternFill("solid", fgColor="DDF0EF")}


def write_table(ws, df: pd.DataFrame, row: int = 1, col: int = 1, widths: dict | None = None, zebra=True):
    for j, name in enumerate(df.columns, start=col):
        c = ws.cell(row=row, column=j, value=name)
        c.fill, c.font = HDR_FILL, HDR_FONT
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for i, rec in enumerate(df.itertuples(index=False), start=row + 1):
        for j, v in enumerate(rec, start=col):
            if isinstance(v, (np.integer,)):
                v = int(v)
            elif isinstance(v, (np.floating,)):
                v = None if np.isnan(v) else float(v)
            elif v is pd.NA or v is pd.NaT:
                v = None
            elif isinstance(v, pd.Timestamp):
                v = v.date()
            c = ws.cell(row=i, column=j, value=v)
            c.border = THIN
            if isinstance(v, date):
                c.number_format = "dd-mmm-yyyy"
            if zebra and (i - row) % 2 == 0:
                c.fill = ZEBRA
    for j, name in enumerate(df.columns, start=col):
        letter = get_column_letter(j)
        if widths and name in widths:
            w = widths[name]
        else:
            longest = max([len(str(name))] + [len(str(v)) for v in df[name].head(300)])
            w = min(max(10, longest + 2), 48)
        ws.column_dimensions[letter].width = max(ws.column_dimensions[letter].width or 0, w)
    ws.row_dimensions[row].height = 30
    return row + len(df)


def style_log_sheet(ws, df):
    end = write_table(ws, df, widths={"issue_detail": 46, "value_a": 26, "value_b": 26, "issue_type": 32})
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(df.columns))}{max(end, 2)}"
    sev_col = get_column_letter(list(df.columns).index("severity") + 1)
    for sev, fill in SEV_FILL.items():
        ws.conditional_formatting.add(f"{sev_col}2:{sev_col}{max(end, 2)}",
                                      CellIsRule(operator="equal", formula=[f'"{sev}"'], fill=fill))


def title_block(ws, title, subtitle):
    ws["A1"] = title
    ws["A1"].font = Font(bold=True, size=16, color=NAVY)
    ws["A2"] = subtitle
    ws["A2"].font = Font(italic=True, size=10, color=GREY)


LOG_COLUMNS = ["log_id", "student_id", "issue_type", "issue_detail", "source_a", "value_a", "source_b",
               "value_b", "owner_dept", "severity", "status", "raised_on", "target_close_date", "check_family"]

DICTIONARY = [
    ("Summary", "Metric / issue_type / owner_dept", "Headline counts. % fully matched = students with zero open issues / all unique student IDs seen in any source."),
    ("Discrepancy Log", "log_id", "Unique reference for the discrepancy (DSC-0001...). Quote it when closing the item."),
    ("Discrepancy Log", "student_id", "Student ID the issue relates to, normalised to upper case without spaces."),
    ("Discrepancy Log", "issue_type", "Standard name of the check that failed (see the Summary sheet for the full list)."),
    ("Discrepancy Log", "issue_detail", "Plain-language explanation of what is wrong."),
    ("Discrepancy Log", "source_a", "System and field holding the first value (e.g. Admissions.date_of_birth)."),
    ("Discrepancy Log", "value_a", "Value found in source A ('(blank)' if empty)."),
    ("Discrepancy Log", "source_b", "System and field holding the value being compared against (blank for single-source checks)."),
    ("Discrepancy Log", "value_b", "Value found in source B."),
    ("Discrepancy Log", "owner_dept", "Department responsible for correcting the record: Admissions, Finance, Academics or Industry Relations."),
    ("Discrepancy Log", "severity", "High = affects money, enrolment or results; Medium = affects statutory reporting; Low = cosmetic / format."),
    ("Discrepancy Log", "status", "Open when raised. Department updates to In Progress / Closed."),
    ("Discrepancy Log", "raised_on", "Date the reconciliation run raised the item (MIS cut-off date)."),
    ("Discrepancy Log", "target_close_date", "raised_on + SLA (High 3 days, Medium 7 days, Low 14 days)."),
    ("Discrepancy Log", "check_family", "Validation = problem inside one system; Reconciliation = two systems disagree."),
    ("<Department> sheets", "(same as Discrepancy Log)", "Filtered work-list of the Discrepancy Log for one owning department."),
    ("Master Data", "student_id", "Student ID present and consistent in Admissions and ERP."),
    ("Master Data", "full_name", "Student name (identical in Admissions and ERP)."),
    ("Master Data", "gender", "Female / Male / Other, from Admissions."),
    ("Master Data", "category", "Social category (GEN, OBC, SC, ST, EWS), from Admissions."),
    ("Master Data", "date_of_birth", "Date of birth (identical in Admissions and ERP)."),
    ("Master Data", "home_state", "Home state, from Admissions."),
    ("Master Data", "programme_name", "Programme (identical in Admissions and ERP)."),
    ("Master Data", "batch", "Year of admission."),
    ("Master Data", "current_semester", "Semester the student is in now, from ERP."),
    ("Master Data", "enrolment_status", "Active / Completed / Withdrawn, from ERP."),
    ("Master Data", "result", "Latest exam result (May-2026 session), from ERP."),
    ("Master Data", "marks_pct", "Marks obtained / maximum marks, latest exam."),
    ("Master Data", "fee_due_2026_27", "Annual fee from the programme master (only for students on roll in 2026-27)."),
    ("Master Data", "fee_paid_2026_27", "Sum of valid receipts in Finance for 2026-27."),
    ("Master Data", "fee_balance", "fee_due - fee_paid."),
    ("Master Data", "engagement_type", "OJT or Placement, from the industry partner file."),
    ("Master Data", "partner_name", "Approved industry partner."),
    ("Master Data", "engagement_status", "Ongoing / Completed (OJT) or Joined / Offer Accepted (placement)."),
    ("Master Data", "annual_ctc_lpa", "Annual salary offered, in lakh rupees (placements only)."),
    ("Source: admissions.csv", "student_id, full_name, gender, date_of_birth, category, programme_name, admission_date, email, phone, home_state", "Admissions office register; one row per admitted student."),
    ("Source: finance_fees.csv", "receipt_no, student_id, academic_year, amount_paid, payment_date, payment_mode", "Fee receipts; one row per receipt."),
    ("Source: academic_erp.csv", "student_id, student_name, date_of_birth, programme_name, batch, current_semester, enrolment_status, exam_session, exam_status, marks_obtained, max_marks, result", "Academic ERP; one row per enrolled student with latest exam result."),
    ("Source: industry_partner_placements.csv", "record_id, student_id, partner_name, engagement_type, start_date, end_date, status, annual_ctc_lpa", "OJT and placement records from Industry Relations."),
    ("Source: programme_master.csv / partner_master.csv", "programme_code, programme_name, annual_fee_inr / partner_name, sector_programme, mou_status", "Reference lists the checks compare against."),
]


def build_master(adm, erp, fin, plc, prog, log_df):
    flagged = set(log_df.student_id)
    m = adm.merge(erp, on="student_id", suffixes=("", "_erp"))
    m = m[~m.student_id.isin(flagged)].copy()
    fee = dict(zip(prog.programme_name, pd.to_numeric(prog.annual_fee_inr)))
    paid = fin.groupby("student_id").valid_amount.sum()
    on_roll = m.enrolment_status.eq("Active")
    m["fee_due_2026_27"] = np.where(on_roll, m.programme_name.map(fee), np.nan)
    m["fee_paid_2026_27"] = m.student_id.map(paid).fillna(0)
    m["fee_balance"] = m.fee_due_2026_27 - m.fee_paid_2026_27
    m["marks_pct"] = (pd.to_numeric(m.marks_obtained, errors="coerce") / pd.to_numeric(m.max_marks, errors="coerce")).round(3)
    p = plc.drop_duplicates("student_id").set_index("student_id")
    m["engagement_type"] = m.student_id.map(p.engagement_type)
    m["partner_name"] = m.student_id.map(p.partner_name)
    m["engagement_status"] = m.student_id.map(p.status)
    m["annual_ctc_lpa"] = pd.to_numeric(m.student_id.map(p.annual_ctc_lpa), errors="coerce")
    m["date_of_birth"] = pd.to_datetime(m.date_of_birth).dt.date
    m["batch"] = pd.to_numeric(m.batch)
    m["current_semester"] = pd.to_numeric(m.current_semester)
    cols = ["student_id", "full_name", "gender", "category", "date_of_birth", "home_state", "programme_name",
            "batch", "current_semester", "enrolment_status", "result", "marks_pct", "fee_due_2026_27",
            "fee_paid_2026_27", "fee_balance", "engagement_type", "partner_name", "engagement_status", "annual_ctc_lpa"]
    return m[cols].sort_values("student_id").reset_index(drop=True)


def write_excel(path, log_df, master, universe_n, as_of):
    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    ws.sheet_view.showGridLines = False
    title_block(ws, f"{INSTITUTE} - Data Reconciliation Summary",
                f"MIS cut-off {as_of:%d %b %Y}  |  Synthetic demo data  |  Sources: Admissions, Finance, Academic ERP, Industry Partners")

    flagged = log_df.student_id.nunique()
    matched = universe_n - flagged
    kpis = pd.DataFrame({
        "Metric": ["Unique student IDs across all sources", "Students with at least one issue",
                   "Students fully matched", "% records fully matched", "Open discrepancies",
                   "High-severity discrepancies"],
        "Value": [universe_n, flagged, matched, matched / universe_n, len(log_df),
                  int((log_df.severity == "High").sum())],
    })
    end = write_table(ws, kpis, row=4, widths={"Metric": 40, "Value": 14}, zebra=True)
    ws.cell(row=8, column=2).number_format = "0.0%"
    for r in range(5, end + 1):
        ws.cell(row=r, column=2).font = Font(bold=True, color=NAVY)

    by_type = (log_df.groupby(["issue_type", "check_family", "owner_dept", "severity"]).size()
               .reset_index(name="count").sort_values(["count"], ascending=False))
    by_type.columns = ["Issue type", "Check family", "Owner dept", "Severity", "Count"]
    start = end + 3
    ws.cell(row=start - 1, column=1, value="Discrepancies by issue type").font = Font(bold=True, size=12, color=TEAL)
    end2 = write_table(ws, by_type, row=start, widths={"Issue type": 40, "Check family": 16, "Owner dept": 20, "Severity": 12, "Count": 10})

    pivot = (log_df.pivot_table(index="owner_dept", columns="severity", values="log_id", aggfunc="count", fill_value=0)
             .reindex(index=DEPARTMENTS, columns=["High", "Medium", "Low"], fill_value=0))
    pivot["Total"] = pivot.sum(axis=1)
    pivot = pivot.reset_index().rename(columns={"owner_dept": "Department"})
    col0 = 8
    ws.cell(row=start - 1, column=col0, value="Discrepancies by owning department").font = Font(bold=True, size=12, color=TEAL)
    pend = write_table(ws, pivot, row=start, col=col0, widths={"Department": 22, "High": 10, "Medium": 10, "Low": 10, "Total": 10})

    chart = BarChart()
    chart.type, chart.grouping, chart.overlap = "bar", "stacked", 100
    chart.title, chart.height, chart.width = "Open items by department", 7.5, 15
    chart.add_data(Reference(ws, min_col=col0 + 1, max_col=col0 + 3, min_row=start, max_row=pend), titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=col0, min_row=start + 1, max_row=pend))
    for s, colr in zip(chart.series, [RED, AMBER, TEAL]):
        s.graphicalProperties.solidFill = colr
        s.graphicalProperties.line.solidFill = colr
    ws.add_chart(chart, f"{get_column_letter(col0)}{pend + 3}")

    log_out = log_df[LOG_COLUMNS]
    style_log_sheet(wb.create_sheet("Discrepancy Log"), log_out)
    for dept in DEPARTMENTS:
        wsd = wb.create_sheet(dept)
        style_log_sheet(wsd, log_out[log_out.owner_dept == dept])

    wsm = wb.create_sheet("Master Data")
    mend = write_table(wsm, master, widths={"full_name": 22, "programme_name": 30, "partner_name": 26})
    wsm.freeze_panes = "C2"
    wsm.auto_filter.ref = f"A1:{get_column_letter(len(master.columns))}{mend}"
    for name, fmt in (("marks_pct", "0.0%"), ("fee_due_2026_27", "#,##0"), ("fee_paid_2026_27", "#,##0"), ("fee_balance", "#,##0")):
        letter = get_column_letter(list(master.columns).index(name) + 1)
        for r in range(2, mend + 1):
            wsm[f"{letter}{r}"].number_format = fmt

    wdd = wb.create_sheet("Data Dictionary")
    dd = pd.DataFrame(DICTIONARY, columns=["Sheet / file", "Column", "Description"])
    write_table(wdd, dd, widths={"Sheet / file": 34, "Column": 48, "Description": 90})
    for row in wdd.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")
    wdd.freeze_panes = "A2"

    for w in wb.worksheets:
        w.sheet_properties.tabColor = TEAL if w.title in DEPARTMENTS else NAVY
    wb.save(path)


# ==========================================================================
# Step 4 - Monthly MIS dashboard
# ==========================================================================
def pct(n, d):
    return 100.0 * n / d if d else 0.0


def compute_mis(adm, erp, fin, plc, prog, log_df, universe_n, as_of):
    fee = dict(zip(prog.programme_name, pd.to_numeric(prog.annual_fee_inr)))
    ad = adm.set_index("student_id")
    erp = erp.copy()
    erp["gender"] = erp.student_id.map(ad.gender).fillna("Not recorded")
    erp["category"] = erp.student_id.map(ad.category).fillna("Not recorded")
    erp["marks"] = pd.to_numeric(erp.marks_obtained, errors="coerce")
    erp["maxm"] = pd.to_numeric(erp.max_marks, errors="coerce")
    active = erp[erp.enrolment_status.eq("Active")]
    paid = fin.groupby("student_id").valid_amount.sum()

    graded = erp[erp.exam_status.eq("Result Published") & erp.marks.notna() & (erp.marks <= erp.maxm)]
    passouts = erp[erp.batch.eq("2023") & erp.result.eq("Pass")]
    placed_ids = set(plc.loc[plc.engagement_type.eq("Placement"), "student_id"])
    ojt_ids = set(plc.loc[plc.engagement_type.eq("OJT"), "student_id"])
    ojt_pool = active[active.batch.eq("2024")]

    def prog_row(name, a, g, po):
        due = a.programme_name.map(fee).sum()
        got = a.student_id.map(paid).fillna(0).sum()
        return {
            "programme": name, "enrolled": len(a),
            "female_pct": pct((a.gender == "Female").sum(), len(a)),
            "collected_lakh": got / 1e5, "collection_pct": pct(got, due),
            "pass_pct": pct((g.result == "Pass").sum(), len(g)),
            "placement_pct": pct(po.student_id.isin(placed_ids).sum(), len(po)),
        }

    rows = []
    open_by_prog = {}
    prog_of = pd.concat([erp.set_index("student_id").programme_name, ad.programme_name]).groupby(level=0).first()
    for sid in log_df.student_id:
        p = prog_of.get(sid)
        if p is not None:
            key = next((n for n in fee if norm_programme(n) == norm_programme(p)), None)
            open_by_prog[key] = open_by_prog.get(key, 0) + 1
    for name in prog.programme_name:
        r = prog_row(name, active[active.programme_name == name], graded[graded.programme_name == name],
                     passouts[passouts.programme_name == name])
        r["open_issues"] = open_by_prog.get(name, 0)
        rows.append(r)
    total = prog_row("All programmes", active, graded, passouts)
    total["open_issues"] = int(len(log_df))

    due_all = active.programme_name.map(fee).sum()
    got_all = active.student_id.map(paid).fillna(0).sum()
    matched = universe_n - log_df.student_id.nunique()

    kpis = [
        {"label": "Total enrolled", "value": f"{len(active):,}", "sub": f"Active students, AY 2026-27"},
        {"label": "Fee collection", "value": f"{pct(got_all, due_all):.1f}%",
         "sub": f"Rs {got_all/1e7:.2f} Cr of Rs {due_all/1e7:.2f} Cr billed"},
        {"label": "Pass rate", "value": f"{pct((graded.result == 'Pass').sum(), len(graded)):.1f}%",
         "sub": f"{len(graded):,} results, May-2026 exams"},
        {"label": "Placement / OJT", "value": f"{pct(passouts.student_id.isin(placed_ids).sum(), len(passouts)):.1f}%",
         "sub": f"pass-outs placed | OJT cover {pct(ojt_pool.student_id.isin(ojt_ids).sum(), len(ojt_pool)):.0f}% (Sem 5)"},
        {"label": "Data match rate", "value": f"{pct(matched, universe_n):.1f}%",
         "sub": f"{matched:,} of {universe_n:,} students clean in all 4 systems", "accent": True},
    ]

    # Regulatory-ready (illustrative) tables
    gcols = ["Male", "Female", "Other", "Not recorded"]
    by_gender = (active.pivot_table(index="programme_name", columns="gender", values="student_id", aggfunc="count", fill_value=0)
                 .reindex(index=list(prog.programme_name), columns=gcols, fill_value=0))
    by_gender["Total"] = by_gender.sum(axis=1)
    by_gender.loc["Total"] = by_gender.sum()
    ccols = ["GEN", "EWS", "OBC", "SC", "ST", "Not recorded"]
    by_cat = active.category.value_counts().reindex(ccols, fill_value=0)
    po_placed = plc[plc.engagement_type.eq("Placement") & plc.student_id.isin(passouts.student_id)]
    graduates = {
        "pass_outs": len(passouts),
        "placed": int(passouts.student_id.isin(placed_ids).sum()),
        "median_ctc": float(pd.to_numeric(po_placed.annual_ctc_lpa, errors="coerce").median()),
        "final_year": int(erp.batch.eq("2023").sum()),
    }

    dept = (log_df.pivot_table(index="owner_dept", columns="severity", values="log_id", aggfunc="count", fill_value=0)
            .reindex(index=DEPARTMENTS, columns=["High", "Medium", "Low"], fill_value=0))
    dept["Total"] = dept.sum(axis=1)

    # Monthly trend (Apr - as_of month)
    months = pd.period_range("2026-04", pd.Period(as_of, "M"), freq="M")
    f = fin[fin.student_id.isin(active.student_id) & fin.payment_date.notna()]
    coll = f.groupby(f.payment_date.dt.to_period("M")).valid_amount.sum().reindex(months, fill_value=0)
    a26 = adm[adm.student_id.str.startswith("ST2026")]
    adm_m = pd.to_datetime(a26.admission_date).dt.to_period("M").value_counts().reindex(months, fill_value=0)
    trend = {"months": months, "collected": coll / 1e5, "cum_pct": coll.cumsum() / due_all * 100, "admissions": adm_m}

    top = (log_df.groupby(["issue_type", "owner_dept"]).size().reset_index(name="n")
           .sort_values("n", ascending=False).head(6))
    return {"top_issues": top, "kpis": kpis, "programmes": rows, "total": total, "by_gender": by_gender, "by_cat": by_cat,
            "graduates": graduates, "dept": dept, "trend": trend}


def trend_chart(trend) -> bytes:
    plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 10, "axes.edgecolor": "#C9D3DD",
                         "axes.labelcolor": "#33415C", "xtick.color": "#33415C", "ytick.color": "#33415C"})
    labels = [m.strftime("%b") for m in trend["months"]]
    x = np.arange(len(labels))
    fig, (ax1, ax3) = plt.subplots(1, 2, figsize=(11.6, 3.9), dpi=150, gridspec_kw={"width_ratios": [1.6, 1]})

    bars = ax1.bar(x, trend["collected"].values, color="#1B2A4A", width=0.6, label="Fee collected (Rs lakh)")
    ax1.bar_label(bars, fmt="%.0f", fontsize=8, color="#1B2A4A", padding=2)
    ax1.set_xticks(x, labels)
    ax1.set_ylabel("Rs lakh")
    ax1.set_ylim(0, trend["collected"].max() * 1.25)
    ax2 = ax1.twinx()
    ax2.plot(x, trend["cum_pct"].values, color="#0F8B8D", marker="o", linewidth=2.2, label="Cumulative collection %")
    ax2.set_ylim(0, 100)
    ax2.set_ylabel("% of annual fee billed", color="#0F8B8D")
    ax2.tick_params(axis="y", colors="#0F8B8D")
    for xi, yi in zip(x, trend["cum_pct"].values):
        ax2.annotate(f"{yi:.0f}%", (xi, yi), textcoords="offset points", xytext=(0, 9), ha="center", fontsize=8,
                     color="#0F8B8D", fontweight="bold", bbox=dict(boxstyle="round,pad=0.2", fc="white", ec="#0F8B8D", lw=0.6))
    ax1.set_title("Fee collection by month, AY 2026-27", loc="left", fontsize=11, color="#1B2A4A", fontweight="bold")
    h1, l1 = ax1.get_legend_handles_labels()
    h2, l2 = ax2.get_legend_handles_labels()
    ax1.legend(h1 + h2, l1 + l2, loc="upper left", frameon=False, fontsize=8)

    bars = ax3.bar(x, trend["admissions"].values, color="#0F8B8D", width=0.6)
    ax3.bar_label(bars, fontsize=8, color="#0F8B8D", padding=2)
    ax3.set_xticks(x, labels)
    ax3.set_ylim(0, max(trend["admissions"].max() * 1.25, 1))
    ax3.set_title("New admissions by month, 2026 intake", loc="left", fontsize=11, color="#1B2A4A", fontweight="bold")
    for ax in (ax1, ax2, ax3):
        ax.spines["top"].set_visible(False)
    ax1.spines["right"].set_visible(False)
    ax3.spines["right"].set_visible(False)
    ax1.grid(axis="y", color="#E6ECF2")
    ax3.grid(axis="y", color="#E6ECF2")
    ax1.set_axisbelow(True)
    ax3.set_axisbelow(True)
    fig.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png", facecolor="white")
    plt.close(fig)
    return buf.getvalue()


def find_chrome() -> str | None:
    for c in (os.environ.get("CHROME_PATH"), DEFAULT_CHROME, shutil.which("chrome"), shutil.which("google-chrome"),
              shutil.which("chromium")):
        if c and Path(c).exists():
            return c
    return None


def render_dashboard(mis, as_of, chart_png: bytes):
    env = Environment(loader=FileSystemLoader(TEMPLATES), autoescape=select_autoescape(["html", "j2"]))
    html = env.get_template("dashboard.html.j2").render(
        institute=INSTITUTE, month=as_of.strftime("%B %Y"), as_of=as_of.strftime("%d %b %Y"),
        chart=base64.b64encode(chart_png).decode(), **mis)
    html_path = OUT / "monthly_mis_dashboard.html"
    html_path.write_text(html, encoding="utf-8")
    (OUT / "monthly_trend.png").write_bytes(chart_png)

    chrome = find_chrome()
    if not chrome:
        print("  ! Chrome not found - skipped PDF/PNG (set CHROME_PATH to enable)")
        return
    uri = html_path.resolve().as_uri()
    common = [chrome, "--headless=new", "--disable-gpu", "--hide-scrollbars", "--no-first-run",
              "--run-all-compositor-stages-before-draw", "--virtual-time-budget=3000"]
    subprocess.run(common + ["--no-pdf-header-footer", "--print-to-pdf-no-header",
                             f"--print-to-pdf={OUT / 'monthly_mis_dashboard.pdf'}", uri],
                   check=True, capture_output=True, timeout=120)
    subprocess.run(common + ["--window-size=1240,1754", f"--screenshot={DOCS / 'dashboard_preview.png'}", uri],
                   check=True, capture_output=True, timeout=120)


# ==========================================================================
# Main
# ==========================================================================
def detection_check(log_df):
    key_path = ROOT / "tests" / "expected_issues.csv"
    if not key_path.exists():
        return
    key = pd.read_csv(key_path, dtype=str)
    found = set(zip(log_df.student_id, log_df.issue_type))
    hit = key.apply(lambda r: (r.student_id, r.planted_issue) in found, axis=1)
    print(f"  Planted issues detected: {hit.sum()}/{len(key)}")
    if not hit.all():
        print(key[~hit].to_string(index=False))


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--as-of", default="2026-10-31", help="MIS cut-off date (YYYY-MM-DD)")
    args = ap.parse_args()
    as_of = date.fromisoformat(args.as_of)
    OUT.mkdir(exist_ok=True)
    DOCS.mkdir(exist_ok=True)

    src = load_sources()
    log = IssueLog()
    max_fee = pd.to_numeric(src["prog"].annual_fee_inr).max()
    print("1. Validating each source ...")
    validate_admissions(src["adm"], log)
    validate_finance(src["fin"], max_fee, log)
    validate_erp(src["erp"], log)
    validate_placements(src["plc"], src["partners"], log)
    n_val = len(log.rows)
    print(f"   {n_val} validation issues")

    print("2. Reconciling across sources on student_id ...")
    adm, erp, fin, plc = clean(src)
    universe = reconcile(adm, erp, fin, plc, log)
    print(f"   {len(log.rows) - n_val} reconciliation issues")
    log_df = log.to_frame(as_of)
    detection_check(log_df)

    print("3. Writing output/reconciliation_report.xlsx ...")
    master = build_master(adm, erp, fin, plc, src["prog"], log_df)
    write_excel(OUT / "reconciliation_report.xlsx", log_df, master, len(universe), as_of)

    print("4. Building monthly MIS dashboard ...")
    mis = compute_mis(adm, erp, fin, plc, src["prog"], log_df, len(universe), as_of)
    render_dashboard(mis, as_of, trend_chart(mis["trend"]))

    matched = len(universe) - log_df.student_id.nunique()
    print(f"\nDone. {len(universe):,} students | {len(log_df)} open discrepancies | "
          f"{matched / len(universe):.1%} fully matched | master data: {len(master):,} clean records")


if __name__ == "__main__":
    main()
