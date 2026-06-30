import os
import pandas as pd
from openpyxl import load_workbook

PLAN_FILE     = os.environ.get("PLAN_FILE",     "/mnt/user-data/outputs/Integrated_Supply_Plan.xlsx")
FORECAST_FILE = os.environ.get("FORECAST_FILE", "/mnt/project/TEMPLATE_Forecast_SOH_1.xlsx")
BOM_FILE      = os.environ.get("BOM_FILE",      "/mnt/project/TEMPLATE_BOM_1.xlsx")
COMP_SOH_FILE = os.environ.get("COMP_SOH_FILE", "/mnt/project/TEMPLATE_BOM_Component_SOH_2.xlsx")

MONTHS = ['Jun-25', 'Jul-25', 'Aug-25', 'Sep-25']


def run_checks():
    """
    Runs every check and returns {check_num: {passed, label, note, diagnosis}}.

    This is the ONLY check shipped by default — it exists as a working,
    minimal template. New checks (added via the UI) get appended below
    by claude_agent.py, following this same shape:

      results[N] = dict(
          passed=<bool>,
          label="<short human-readable description>",
          note="<one-line detail, e.g. the actual value found>",
          diagnosis="PASS" if passed else "<CALC ERROR: ...> or <GENUINE BUSINESS ALERT: ...>",
      )

    Use "CALC ERROR" in diagnosis when a failure means something is broken
    in the plan logic (the pipeline will try to auto-fix it).
    Use "GENUINE BUSINESS ALERT" when a failure reflects real business
    conditions, not a bug (the pipeline will just flag it, not fix it).
    """
    results = {}

    # ── Check 1: Executive Summary totals match source forecast ──────────

    try:
        ex = pd.read_excel(PLAN_FILE, sheet_name='Executive Summary', header=None).iloc[10:, :]
        fc = pd.read_excel(FORECAST_FILE, header=None, skiprows=4)

        fc_row = ex[ex.iloc[:, 1].astype(str).str.contains('Total Forecast', na=False)]

        passed = True
        diffs = {}
        for i, m in enumerate(MONTHS):
            expected = float(fc_row.iloc[0, i + 2]) if len(fc_row) and not pd.isna(fc_row.iloc[0, i + 2]) else 0
            # NOTE: column letters in the raw forecast file vary by template —
            # adjust this lookup to match your actual forecast file layout.
            actual = fc.iloc[:, i + 2].sum() if fc.shape[1] > i + 2 else 0
            ok = abs(expected - actual) < 1
            diffs[m] = ok
            if not ok:
                passed = False

        results[1] = dict(
            passed=passed,
            label="Executive Summary totals match source forecast",
            note=f"checked months: {diffs}",
            diagnosis="PASS" if passed else "CALC ERROR: executive summary totals do not match source forecast input",
        )
    except Exception as e:
        results[1] = dict(
            passed=False,
            label="Executive Summary totals match source forecast",
            note=str(e),
            diagnosis=f"CALC ERROR: check could not run — {e}",
        )

    

    return results


def write_alerts_sheet(results):
    """Writes a simple Alerts & Flags sheet into the plan workbook, listing
    every failed check with its diagnosis."""
    wb = load_workbook(PLAN_FILE)

    if 'Alerts & Flags' in wb.sheetnames:
        del wb['Alerts & Flags']
    ws = wb.create_sheet('Alerts & Flags')

    ws['A1'] = 'Check #'
    ws['B1'] = 'Label'
    ws['C1'] = 'Status'
    ws['D1'] = 'Diagnosis'

    row = 2
    for num, r in sorted(results.items()):
        ws.cell(row=row, column=1, value=num)
        ws.cell(row=row, column=2, value=r['label'])
        ws.cell(row=row, column=3, value='PASS' if r['passed'] else 'FAIL')
        ws.cell(row=row, column=4, value=r.get('diagnosis', ''))
        row += 1

    wb.save(PLAN_FILE)



