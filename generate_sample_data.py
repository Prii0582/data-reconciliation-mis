"""
Generate synthetic source-system extracts for the reconciliation demo.

Everything produced here is FICTIONAL: the institute, students, partners,
phone numbers and e-mail addresses are randomly generated. A fixed seed
makes every run identical, so the reports in /output are reproducible.

Outputs (in ./data):
    admissions.csv                   - Admissions office register
    finance_fees.csv                 - Fee receipts for academic year 2026-27
    academic_erp.csv                 - Enrolment, programme, semester, exam results
    industry_partner_placements.csv  - On-the-job training (OJT) and placements
    programme_master.csv             - Approved programme list + annual fee
    partner_master.csv               - Approved industry partner list
    tests/expected_issues.csv        - The problems deliberately planted, used to
                                       prove that reconcile.py finds all of them
"""
from __future__ import annotations

from datetime import date, timedelta
from pathlib import Path

import numpy as np
import pandas as pd

SEED = 2026
ROOT = Path(__file__).resolve().parent
DATA = ROOT / "data"

rng = np.random.default_rng(SEED)

# --------------------------------------------------------------------------
# Reference data
# --------------------------------------------------------------------------
PROGRAMMES = [  # code, canonical name, annual fee (INR), share of intake
    ("BVRM", "B.Voc Retail Management", 68000, 0.19),
    ("BVHC", "B.Voc Healthcare Services", 82000, 0.18),
    ("BVLS", "B.Voc Logistics & Supply Chain", 72000, 0.16),
    ("BBAS", "BBA Applied Skills", 90000, 0.14),
    ("BSDA", "B.Sc Data Analytics", 95000, 0.18),
    ("BVAT", "B.Voc Automotive Technology", 76000, 0.15),
]
PROG_NAMES = [p[1] for p in PROGRAMMES]
PROG_SHARE = np.array([p[3] for p in PROGRAMMES])
PROG_SHARE = PROG_SHARE / PROG_SHARE.sum()

# How the Admissions team sometimes types the same programme
SPELLING_VARIANTS = {
    "B.Voc Retail Management": ["BVoc Retail Mgmt", "B.Voc. Retail Management", "B Voc Retail Management"],
    "B.Voc Healthcare Services": ["BVoc Healthcare Services", "B.Voc. Healthcare Services"],
    "B.Voc Logistics & Supply Chain": ["B.Voc Logistics and Supply Chain", "BVoc Logistics & Supply Chain"],
    "BBA Applied Skills": ["BBA - Applied Skills", "B.B.A Applied Skills"],
    "B.Sc Data Analytics": ["BSc Data Analytics", "B.Sc. Data Analytics"],
    "B.Voc Automotive Technology": ["B.Voc Automotive Tech", "BVoc Automotive Technology"],
}

# Approved partners per programme (invented names)
PARTNERS = {
    "B.Voc Retail Management": ["Brindlecove Retail", "Marlowe Mart Stores"],
    "B.Voc Healthcare Services": ["Tarnwick Healthcare", "Sorrelby Diagnostics"],
    "B.Voc Logistics & Supply Chain": ["Quillmarsh Logistics", "Fenmoor Freight"],
    "BBA Applied Skills": ["Pembric Hospitality", "Cardale Business Services"],
    "B.Sc Data Analytics": ["Ostrella Analytics", "Vantrel Data Labs"],
    "B.Voc Automotive Technology": ["Velloran Motors", "Drayholt Auto Components"],
}
UNLISTED_PARTNERS = [
    "Brindlecove Retail Pvt Ltd",  # variant of an approved partner
    "Velloran Motor",              # typo of an approved partner
    "Kestrafield Warehousing",     # never onboarded
    "Lumora Clinics",              # never onboarded
    "Zenthic Auto Works",          # never onboarded
]

FEMALE = ["Aarohi", "Ananya", "Bhavna", "Charvi", "Diya", "Esha", "Farida", "Gauri", "Harini",
          "Ishita", "Jhanvi", "Kavya", "Lavanya", "Meera", "Nandini", "Oindrila", "Pallavi", "Riya",
          "Sanjana", "Tanvi", "Urmila", "Vaishnavi", "Yamini", "Zoya", "Sneha", "Rupa", "Ankita",
          "Mitali", "Payal", "Swati"]
MALE = ["Aarav", "Arjun", "Bikram", "Chirag", "Dev", "Farhan", "Gaurav", "Harsh", "Imran", "Jay",
        "Kunal", "Lokesh", "Manav", "Nikhil", "Omkar", "Pranav", "Rahul", "Sahil", "Tarun", "Uday",
        "Varun", "Yash", "Zubin", "Ritesh", "Sourav", "Abhik", "Debashis", "Kartik", "Rohan", "Vikram"]
SURNAMES = ["Sharma", "Das", "Iyer", "Khan", "Patel", "Roy", "Banerjee", "Nair", "Ghosh", "Reddy",
            "Mehta", "Sen", "Kulkarni", "Bose", "Mishra", "Pillai", "Chatterjee", "Joshi", "Saxena",
            "Mondal", "Verma", "Ali", "Pandey", "Rao", "Kapoor", "Mukherjee", "Shetty", "Gupta",
            "Desai", "Sinha"]
STATES = ["West Bengal", "Odisha", "Jharkhand", "Bihar", "Assam", "Maharashtra", "Karnataka",
          "Tamil Nadu", "Uttar Pradesh", "Telangana"]

BATCHES = {2023: 260, 2024: 300, 2025: 310, 2026: 330}   # = 1,200 students
SEMESTER = {2023: 6, 2024: 5, 2025: 3, 2026: 1}
REPORT_END = date(2026, 10, 31)


def rand_date(start: date, end: date) -> date:
    return start + timedelta(days=int(rng.integers(0, (end - start).days + 1)))


def iso(d) -> str:
    return d.isoformat() if isinstance(d, date) else ""


# --------------------------------------------------------------------------
# 1. Clean "truth" for every student
# --------------------------------------------------------------------------
def build_students() -> pd.DataFrame:
    rows = []
    for batch, n in BATCHES.items():
        for i in range(1, n + 1):
            gender = rng.choice(["Female", "Male", "Other"], p=[0.46, 0.52, 0.02])
            pool = FEMALE if gender == "Female" else MALE if gender == "Male" else FEMALE + MALE
            first, last = rng.choice(pool), rng.choice(SURNAMES)
            if batch == 2026:
                admitted = rand_date(date(2026, 5, 15), date(2026, 9, 30))
            else:
                admitted = rand_date(date(batch, 6, 1), date(batch, 8, 31))
            dob = admitted - timedelta(days=int(rng.integers(18 * 365, 21 * 365)))
            rows.append({
                "student_id": f"ST{batch}-{i:04d}",
                "full_name": f"{first} {last}",
                "gender": gender,
                "date_of_birth": dob,
                "category": rng.choice(["GEN", "OBC", "SC", "ST", "EWS"], p=[0.38, 0.32, 0.15, 0.07, 0.08]),
                "programme_name": rng.choice(PROG_NAMES, p=PROG_SHARE),
                "admission_date": admitted,
                "email": f"{first}.{last}{int(rng.integers(10, 99))}@mail.example".lower(),
                "phone": "90000" + f"{int(rng.integers(0, 100000)):05d}",
                "home_state": rng.choice(STATES),
                "batch": batch,
            })
    return pd.DataFrame(rows)


def build_erp(students: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for s in students.itertuples():
        status = "Completed" if s.batch == 2023 else ("Withdrawn" if rng.random() < 0.03 else "Active")
        marks, exam_status, result = "", "Not Due", ""
        if s.batch < 2026 and status != "Withdrawn":
            exam_status = "Result Published"
            marks = int(np.clip(rng.normal(355, 80), 110, 598))
            result = "Pass" if marks >= 240 else "Fail"
        elif status == "Withdrawn":
            exam_status = "Not Applicable"
        rows.append({
            "student_id": s.student_id,
            "student_name": s.full_name,
            "date_of_birth": s.date_of_birth,
            "programme_name": s.programme_name,
            "batch": s.batch,
            "current_semester": SEMESTER[s.batch],
            "enrolment_status": status,
            "exam_session": "May-2026" if s.batch < 2026 else "",
            "exam_status": exam_status,
            "marks_obtained": marks,
            "max_marks": 600 if s.batch < 2026 else "",
            "result": result,
        })
    return pd.DataFrame(rows)


def build_fees(students: pd.DataFrame, erp: pd.DataFrame) -> pd.DataFrame:
    fee = {p[1]: p[2] for p in PROGRAMMES}
    rows = []
    modes = ["UPI", "Net Banking", "Card", "Bank Transfer", "Demand Draft"]
    for s in students[students.batch >= 2024].itertuples():
        due = fee[s.programme_name]
        r = rng.random()
        plan = "full" if r < 0.72 else "partial" if r < 0.89 else "none"
        if plan == "none":
            continue
        if s.batch == 2026:
            d1 = min(s.admission_date + timedelta(days=int(rng.integers(0, 10))), REPORT_END)
        else:
            d1 = rand_date(date(2026, 4, 1), date(2026, 6, 30))
        if plan == "partial":
            rows.append((s.student_id, due // 2, d1))
        elif d1 < date(2026, 8, 15) and rng.random() < 0.6:
            first = int(round(due * 0.6, -2))
            rows.append((s.student_id, first, d1))
            rows.append((s.student_id, due - first, rand_date(max(d1 + timedelta(days=30), date(2026, 8, 15)), REPORT_END)))
        else:
            rows.append((s.student_id, due, d1))
    df = pd.DataFrame(rows, columns=["student_id", "amount_paid", "payment_date"]).sort_values("payment_date")
    df.insert(0, "receipt_no", [f"RCPT-26-{i:05d}" for i in range(1, len(df) + 1)])
    df["academic_year"] = "2026-27"
    df["payment_mode"] = rng.choice(modes, size=len(df), p=[0.45, 0.2, 0.12, 0.18, 0.05])
    return df[["receipt_no", "student_id", "academic_year", "amount_paid", "payment_date", "payment_mode"]].reset_index(drop=True)


def build_placements(students: pd.DataFrame, erp: pd.DataFrame) -> pd.DataFrame:
    status = erp.set_index("student_id")["enrolment_status"]
    rows = []
    for s in students.itertuples():
        partner = rng.choice(PARTNERS[s.programme_name])
        if s.batch == 2023 and rng.random() < 0.72:
            start = rand_date(date(2026, 7, 1), date(2026, 9, 30))
            rows.append((s.student_id, partner, "Placement", start, None,
                         rng.choice(["Joined", "Offer Accepted"], p=[0.8, 0.2]),
                         round(float(rng.uniform(2.2, 4.8)), 2)))
        elif s.batch == 2024 and status[s.student_id] == "Active" and rng.random() < 0.85:
            start = rand_date(date(2026, 7, 1), date(2026, 8, 20))
            rows.append((s.student_id, partner, "OJT", start, start + timedelta(days=180), "Ongoing", None))
        elif s.batch == 2025 and status[s.student_id] == "Active" and rng.random() < 0.60:
            start = rand_date(date(2026, 5, 4), date(2026, 6, 1))
            rows.append((s.student_id, partner, "OJT", start, start + timedelta(days=42), "Completed", None))
    df = pd.DataFrame(rows, columns=["student_id", "partner_name", "engagement_type", "start_date",
                                     "end_date", "status", "annual_ctc_lpa"])
    df.insert(0, "record_id", [f"PL-{i:05d}" for i in range(1, len(df) + 1)])
    return df


# --------------------------------------------------------------------------
# 2. Plant realistic problems (each student gets at most one planted issue)
# --------------------------------------------------------------------------
class Planter:
    def __init__(self):
        self.used: set[str] = set()
        self.key: list[tuple[str, str]] = []

    def pick(self, candidates, n: int) -> list[str]:
        pool = [c for c in candidates if c not in self.used]
        chosen = list(rng.choice(pool, size=n, replace=False))
        self.used.update(chosen)
        return chosen

    def note(self, sid: str, issue: str):
        self.key.append((sid, issue))


def plant_issues(adm, erp, fin, plc):
    p = Planter()
    ids_2026 = adm.loc[adm.batch == 2026, "student_id"].tolist()
    paid = set(fin.student_id)
    erp_status = erp.set_index("student_id")["enrolment_status"]
    active_2026 = [i for i in ids_2026 if erp_status[i] == "Active"]
    # natural "Withdrawn but paid" cases are left alone (detected as a bonus finding)
    for sid, st in erp_status.items():
        if st == "Withdrawn":
            p.used.add(sid)

    # Student admitted but missing in ERP (some have paid fees)
    for sid in p.pick([i for i in active_2026 if i in paid], 12):
        erp = erp[erp.student_id != sid]
        p.note(sid, "Fee paid but not enrolled")
    for sid in p.pick([i for i in active_2026 if i not in paid], 10):
        erp = erp[erp.student_id != sid]
        p.note(sid, "Admitted but missing in ERP")

    # In ERP but no admission record
    extra = []
    for k in range(6):
        sid = f"ST2025-{401 + k:04d}"
        g = rng.choice(["Female", "Male"])
        name = f"{rng.choice(FEMALE if g == 'Female' else MALE)} {rng.choice(SURNAMES)}"
        marks = int(rng.integers(260, 520))
        extra.append({"student_id": sid, "student_name": name,
                      "date_of_birth": date(2006, int(rng.integers(1, 13)), int(rng.integers(1, 28))),
                      "programme_name": rng.choice(PROG_NAMES), "batch": 2025, "current_semester": 3,
                      "enrolment_status": "Active", "exam_session": "May-2026",
                      "exam_status": "Result Published", "marks_obtained": marks, "max_marks": 600,
                      "result": "Pass"})
        p.note(sid, "In ERP but no admission record")
    erp = pd.concat([erp, pd.DataFrame(extra)], ignore_index=True)

    eligible = [i for i in adm.student_id if i in set(erp.student_id)]

    # Programme name spelled differently in Admissions
    for sid in p.pick(eligible, 45):
        prog = adm.loc[adm.student_id == sid, "programme_name"].iloc[0]
        adm.loc[adm.student_id == sid, "programme_name"] = rng.choice(SPELLING_VARIANTS[prog])
        p.note(sid, "Programme name spelling variant")

    # Programme genuinely different (transfer not updated in ERP)
    for sid in p.pick(eligible, 8):
        cur = erp.loc[erp.student_id == sid, "programme_name"].iloc[0]
        erp.loc[erp.student_id == sid, "programme_name"] = rng.choice([x for x in PROG_NAMES if x != cur])
        p.note(sid, "Programme mismatch")

    # Date-of-birth mismatch between Admissions and ERP
    for sid in p.pick(eligible, 18):
        d = adm.loc[adm.student_id == sid, "date_of_birth"].iloc[0]
        new = date(d.year, d.day, d.month) if (d.day <= 12 and d.day != d.month) else date(d.year + 1, d.month, min(d.day, 28))
        erp.loc[erp.student_id == sid, "date_of_birth"] = new
        p.note(sid, "Date of birth mismatch")

    published = erp.loc[erp.exam_status == "Result Published", "student_id"].tolist()
    # Result published but marks missing
    for sid in p.pick([i for i in published if i in eligible], 14):
        erp.loc[erp.student_id == sid, "marks_obtained"] = ""
        p.note(sid, "Result published but marks missing")
    # Marks above maximum
    for sid in p.pick([i for i in published if i in eligible], 4):
        erp.loc[erp.student_id == sid, ["marks_obtained", "result"]] = [int(rng.integers(605, 660)), "Pass"]
        p.note(sid, "Marks out of range")

    # Duplicate IDs: exact duplicate rows in Admissions, re-used IDs, ERP duplicates
    for sid in p.pick(eligible, 5):
        adm = pd.concat([adm, adm[adm.student_id == sid]], ignore_index=True)
        p.note(sid, "Duplicate student ID")
    for sid in p.pick(eligible, 3):
        clone = adm[adm.student_id == sid].copy()
        clone["full_name"] = f"{rng.choice(FEMALE + MALE)} {rng.choice(SURNAMES)}"
        clone["email"] = "dup." + clone["email"]
        adm = pd.concat([adm, clone], ignore_index=True)
        p.note(sid, "Duplicate student ID")
    for sid in p.pick(eligible, 4):
        erp = pd.concat([erp, erp[erp.student_id == sid]], ignore_index=True)
        p.note(sid, "Duplicate student ID")

    # Missing mandatory fields / invalid formats / out-of-range in Admissions
    for col, n in (("gender", 4), ("category", 5), ("date_of_birth", 3)):
        for sid in p.pick(eligible, n):
            adm.loc[adm.student_id == sid, col] = None
            p.note(sid, "Missing mandatory field")
    for sid in p.pick(eligible, 6):
        adm.loc[adm.student_id == sid, "email"] = adm.loc[adm.student_id == sid, "email"].str.replace("@", "", regex=False)
        p.note(sid, "Invalid email format")
    for sid in p.pick(eligible, 5):
        adm.loc[adm.student_id == sid, "phone"] = rng.choice(["900001234", "90000-12AB5", "9000012345678"])
        p.note(sid, "Invalid phone format")
    for sid in p.pick([i for i in eligible if not i.startswith("ST2023")], 3):
        adm.loc[adm.student_id == sid, "date_of_birth"] = date(2014, int(rng.integers(1, 13)), 15)
        erp.loc[erp.student_id == sid, "date_of_birth"] = adm.loc[adm.student_id == sid, "date_of_birth"].iloc[0]
        p.note(sid, "Date of birth out of range")

    # Finance problems
    fin_ids = [i for i in fin.student_id.unique() if i in eligible]
    for sid in p.pick(fin_ids, 6):
        row = fin[fin.student_id == sid].iloc[[0]]
        fin = pd.concat([fin, row], ignore_index=True)
        p.note(sid, "Duplicate receipt")
    for sid in p.pick(fin_ids, 5):
        idx = fin.index[fin.student_id == sid][0]
        fin.loc[idx, "student_id"] = rng.choice([sid.lower(), f" {sid}", f"{sid} "])
        p.note(sid, "Invalid student ID format")
    for sid in p.pick(fin_ids, 4):
        idx = fin.index[fin.student_id == sid][0]
        fin.loc[idx, "payment_date"] = rng.choice(["31-09-2026", "2026-13-05", "TBC"])
        p.note(sid, "Invalid date")
    for sid in p.pick(fin_ids, 2):
        fin = pd.concat([fin, pd.DataFrame([{"receipt_no": f"RCPT-26-REV{len(fin):03d}", "student_id": sid,
                                              "academic_year": "2026-27", "amount_paid": -5000,
                                              "payment_date": date(2026, 10, 12), "payment_mode": "Bank Transfer"}])],
                        ignore_index=True)
        p.note(sid, "Negative fee amount")
    for k in range(4):
        sid = f"ST2026-{901 + k:04d}"
        fin = pd.concat([fin, pd.DataFrame([{"receipt_no": f"RCPT-26-X{k:04d}", "student_id": sid,
                                              "academic_year": "2026-27", "amount_paid": 34000,
                                              "payment_date": date(2026, 9, 2 + k), "payment_mode": "UPI"}])],
                        ignore_index=True)
        p.note(sid, "Fee receipt for unknown student")

    # Industry partner / placement problems
    plc_ids = [i for i in plc.student_id if i in eligible]
    for sid in p.pick(plc_ids, 11):
        plc.loc[plc.student_id == sid, "partner_name"] = rng.choice(UNLISTED_PARTNERS)
        p.note(sid, "Partner not in master list")
    for sid in p.pick([i for i in plc_ids if not i.startswith("ST2023")], 3):
        idx = plc.index[plc.student_id == sid][0]
        s, e = plc.loc[idx, "start_date"], plc.loc[idx, "end_date"]
        plc.loc[idx, ["start_date", "end_date"]] = [e, s]
        p.note(sid, "OJT end date before start date")
    for k in range(3):
        sid = f"ST2023-{951 + k:04d}"
        plc = pd.concat([plc, pd.DataFrame([{"record_id": f"PL-9{k:04d}", "student_id": sid,
                                              "partner_name": "Fenmoor Freight", "engagement_type": "Placement",
                                              "start_date": date(2026, 8, 1), "end_date": None,
                                              "status": "Joined", "annual_ctc_lpa": 3.1}])], ignore_index=True)
        p.note(sid, "Placement record for unknown student")

    key = pd.DataFrame(p.key, columns=["student_id", "planted_issue"]).drop_duplicates()
    return adm, erp, fin, plc, key


def main():
    DATA.mkdir(exist_ok=True)
    students = build_students()
    erp = build_erp(students)
    fin = build_fees(students, erp)
    plc = build_placements(students, erp)
    adm = students.copy()

    adm, erp, fin, plc, key = plant_issues(adm, erp, fin, plc)

    # Shuffle rows so planted problems are not grouped together
    adm = adm.sample(frac=1, random_state=SEED).reset_index(drop=True)
    erp = erp.sample(frac=1, random_state=SEED + 1).reset_index(drop=True)
    for df, cols in ((adm, ["date_of_birth", "admission_date"]), (erp, ["date_of_birth"]),
                     (fin, ["payment_date"]), (plc, ["start_date", "end_date"])):
        for c in cols:
            df[c] = df[c].map(lambda v: iso(v) if isinstance(v, date) else ("" if v is None or (isinstance(v, float) and np.isnan(v)) else v))

    adm.drop(columns=["batch"]).to_csv(DATA / "admissions.csv", index=False)
    erp.to_csv(DATA / "academic_erp.csv", index=False)
    fin.to_csv(DATA / "finance_fees.csv", index=False)
    plc.to_csv(DATA / "industry_partner_placements.csv", index=False)
    pd.DataFrame(PROGRAMMES, columns=["programme_code", "programme_name", "annual_fee_inr", "intake_share"]) \
        .drop(columns="intake_share").assign(duration_years=3).to_csv(DATA / "programme_master.csv", index=False)
    pd.DataFrame([(p, prog) for prog, ps in PARTNERS.items() for p in ps], columns=["partner_name", "sector_programme"]) \
        .assign(mou_status="Active").to_csv(DATA / "partner_master.csv", index=False)
    (ROOT / "tests").mkdir(exist_ok=True)
    key.to_csv(ROOT / "tests" / "expected_issues.csv", index=False)

    print(f"admissions.csv                  {len(adm):>5} rows")
    print(f"academic_erp.csv                {len(erp):>5} rows")
    print(f"finance_fees.csv                {len(fin):>5} rows")
    print(f"industry_partner_placements.csv {len(plc):>5} rows")
    print(f"planted issues                  {len(key):>5}")


if __name__ == "__main__":
    main()
