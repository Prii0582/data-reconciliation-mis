# Data Reconciliation and Monthly MIS

This project checks student data across four systems of a university, logs every mismatch against the department that owns it, and builds a one-page monthly MIS dashboard at the end.

The data here is dummy data. The institute (Eastvale Institute of Applied Skills), students, industry partners and figures are all made up. I have done similar reconciliation work in my job; I can't share that, so I rebuilt the idea with made-up data.

## Sample output

![Monthly MIS dashboard](docs/dashboard_preview.png)

Sample report generated from the dummy data. Full PDF: [output/monthly_mis_dashboard.pdf](output/monthly_mis_dashboard.pdf). Mismatch workbook: [output/reconciliation_report.xlsx](output/reconciliation_report.xlsx).

## Why I made this

In a university or skills institute, student data sits in different places. Admissions has its register, Finance has fee receipts, the Academic ERP has enrolment and exam results, and Industry Relations tracks OJT and placements. Over time these go out of sync. A student is admitted but never created in the ERP, a fee is paid for someone who is not enrolled, a programme name is typed three different ways, dates of birth don't match. Then the monthly MIS and returns like NIRF / AISHE are prepared on numbers that don't agree.

## What it does

1. Checks each system on its own for blank mandatory fields, duplicates, wrong formats and impossible values.
2. Matches the four systems on Student ID (like SQL joins) and flags every difference.
3. Gives each issue a type, a severity and an owner department (Admissions, Finance, Academics or Industry Relations), with a target close date.
4. Writes an Excel workbook with a summary, the full discrepancy log, one work-list per department, a clean master data sheet and a data dictionary.
5. Builds a one-page monthly MIS dashboard (HTML, PDF and PNG) with KPIs, a programme-wise table, a monthly trend chart and a sample regulatory figures block.

Results on the sample data (1,200 students, about 4,200 source rows):

- 1,213 unique student IDs checked across 4 systems
- 214 discrepancies logged, 71 of them high severity
- 82.4% of students fully matched in every system
- 193 of 193 planted test problems caught
- 999 clean master records
- The whole month reruns in about 30 seconds

## Checks

Inside each system:

- Required fields left blank (gender, category, date of birth, programme, etc.)
- Same Student ID used twice, as an exact copy or by two different people
- Same fee receipt entered twice, so money gets counted double
- Email without "@", phone number not 10 digits
- Student IDs with wrong case or extra spaces
- Dates that don't exist (for example 31 September)
- Age at admission under 16 or over 45
- Negative fee amounts (a reversal posted as a receipt)
- Marks above the maximum, or Pass/Fail not matching the marks
- Result shown as published but marks missing
- OJT end date before start date
- Industry partner not on the approved MoU list

Across systems:

- Admitted, but missing in the Academic ERP
- Fee paid, but student not enrolled
- Fee paid, but student has withdrawn (refund to be checked)
- In the ERP, but no admission record
- Fee or placement record for a student ID that doesn't exist anywhere else
- Same programme spelled differently ("BVoc Retail Mgmt" vs "B.Voc Retail Management")
- Student in a different programme in each system
- Date of birth or name different between Admissions and the ERP

Each issue gets an owner department, a severity (High, Medium or Low) and a target close date (3, 7 or 14 days).

## Excel report

| Sheet | What it has |
|---|---|
| Summary | Headline counts, % fully matched, counts by issue type and department, a stacked bar chart |
| Discrepancy Log | Every issue with student, values in both sources, owner, severity, status, raised on and target date |
| Admissions / Finance / Academics / Industry Relations | Each department's own work-list |
| Master Data | 999 clean records that passed every check |
| Data Dictionary | What each column means and which file it comes from |

## Dashboard

- KPI cards: total enrolled, fee collection %, pass %, placement / OJT %, data match rate
- Monthly trend: fee collected per month with cumulative collection %, and new admissions by month
- Programme-wise table: enrolment, female %, fees, pass %, placement %, open data issues
- Regulatory figures (marked illustrative): students by programme and gender, by social category, pass-outs, placements and median salary, in the layout NIRF / AISHE returns ask for
- Data quality: open items by department and severity, and the most common issue types

## How to run

```bash
pip install -r requirements.txt
python generate_sample_data.py
python reconcile.py
python reconcile.py --as-of 2026-10-31   # to pick the MIS cut-off date
```

The PDF and PNG are made with headless Google Chrome. If Chrome is not in the default Windows location, set `CHROME_PATH`. Without Chrome you still get the HTML dashboard and the Excel file.

## Notes

- `data/` has the dummy source files and the programme and partner master lists.
- `tests/expected_issues.csv` lists the problems I planted in the data on purpose. `reconcile.py` prints how many of them it caught.
- The sample data uses a fixed seed, so the numbers come out the same every time.
