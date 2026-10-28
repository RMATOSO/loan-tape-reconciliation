"""
Loan tape vs capital ledger reconciliation + 12-month cash forecast.

Reads data/loan_tape.csv and data/capital_ledger.csv and writes
output/Loan_Tape_Review.xlsx with:
  Summary        - what the data establishes, what is assumed, what to verify, next steps
  Issues         - every discrepancy with severity and a proposed resolution
  Reconciliation - one row per loan: tape vs ledger
  Portfolio      - duration, concentration and loans past maturity
  Assumptions    - forecast inputs (edit the yellow cells)
  Loans          - active loans with expected payoff month (base and delayed, by formula)
  Forecast       - 12-month cash forecast, base vs delayed payoffs, capital needed (formulas)

Usage: python scripts/reconcile.py [--as-of 2026-10-31]
"""
import argparse
import datetime as dt
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parents[1]
DATE_TOL_DAYS = 3        # funding/payoff dates may differ by up to 3 days
AMOUNT_TOL = 1.00        # amounts may differ by up to $1 (rounding)

HEAD = PatternFill("solid", fgColor="1F3A5F")
INPUT = PatternFill("solid", fgColor="FFF2CC")
SEV_FILL = {"High": "F4CCCC", "Medium": "FCE5CD", "Low": "D9EAD3"}
MONEY = '#,##0;[Red]-#,##0'
MONEY2 = '#,##0.00;[Red]-#,##0.00'


def header(ws, row, cols):
    for i, c in enumerate(cols, 1):
        cell = ws.cell(row, i, c)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = HEAD
        cell.alignment = Alignment(vertical="center", wrap_text=True)


def widths(ws, ws_widths):
    for i, w in enumerate(ws_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def reconcile(tape, ledger, as_of):
    issues = []
    rows = []

    def add(sev, loan, issue, evidence, resolution):
        issues.append({"Severity": sev, "Loan": loan, "Issue": issue,
                       "Evidence": evidence, "Proposed resolution": resolution})

    fund = ledger[ledger.type == "Funding"]
    pay = ledger[ledger.type == "Payoff"]

    for _, ln in tape.iterrows():
        lid = ln.loan_id
        f = fund[fund.loan_id == lid]
        p = pay[pay.loan_id == lid]
        f_amt = -f.amount.sum()
        p_amt = p.amount.sum()
        status = "OK"

        if len(f) == 0:
            status = "Check"
            add("High", lid, "Funding in tape but not in ledger",
                f"Tape: {ln.principal:,.0f} funded {ln.funding_date:%Y-%m-%d}; ledger: no funding row",
                "Request bank statement for the funding date; book the missing ledger entry or correct the tape")
        elif len(f) > 1:
            status = "Check"
            add("High", lid, "Duplicate funding in ledger",
                f"{len(f)} funding rows totalling {f_amt:,.0f} vs tape principal {ln.principal:,.0f}",
                "Confirm against bank statement and reverse the duplicate ledger row")
        else:
            if abs(f_amt - ln.principal) > AMOUNT_TOL:
                status = "Check"
                add("High", lid, "Funding amount differs",
                    f"Tape {ln.principal:,.0f} vs ledger {f_amt:,.0f}", "Confirm wire amount with bank statement")
            gap = abs((f.date.iloc[0] - ln.funding_date).days)
            if gap > DATE_TOL_DAYS:
                status = "Check"
                add("Low", lid, "Funding date differs",
                    f"Tape {ln.funding_date:%Y-%m-%d} vs ledger {f.date.iloc[0]:%Y-%m-%d} ({gap} days)",
                    "Use the bank value date in both records; affects interest accrual")

        if ln.status == "Paid off":
            if len(p) == 0:
                status = "Check"
                add("High", lid, "Tape says paid off, ledger has no payoff",
                    f"Tape payoff {ln.payoff_amount:,.2f} on {ln.payoff_date:%Y-%m-%d}",
                    "Find the payoff receipt; if not received, the loan is still outstanding")
            elif abs(p_amt - ln.payoff_amount) > AMOUNT_TOL:
                status = "Check"
                add("Medium", lid, "Payoff amount differs",
                    f"Tape {ln.payoff_amount:,.2f} vs ledger {p_amt:,.2f} (diff {p_amt - ln.payoff_amount:,.2f})",
                    "Check payoff statement: fees, interest days or a short payment")
        elif ln.maturity_date < as_of:
            status = "Check"
            add("Medium", lid, "Active loan past maturity",
                f"Matured {ln.maturity_date:%Y-%m-%d}, {(as_of - ln.maturity_date).days} days ago",
                "Confirm extension terms or start default process; treat payoff as delayed in the forecast")

        rows.append({"Loan": lid, "Borrower": ln.borrower_id, "Status (tape)": ln.status,
                     "Principal (tape)": ln.principal, "Funding (ledger)": f_amt,
                     "Payoff (tape)": ln.payoff_amount if ln.status == "Paid off" else None,
                     "Payoff (ledger)": p_amt if len(p) else None, "Result": status})

    unknown = sorted(set(fund.loan_id) - set(tape.loan_id))
    for lid in unknown:
        amt = -fund[fund.loan_id == lid].amount.sum()
        add("High", lid, "Funding in ledger for a loan not in tape",
            f"Ledger funding {amt:,.0f}", "Obtain loan documents and add to tape, or reverse if posted in error")

    order = {"High": 0, "Medium": 1, "Low": 2}
    issues.sort(key=lambda r: (order[r["Severity"]], r["Loan"]))
    return pd.DataFrame(rows), pd.DataFrame(issues)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--as-of", default="2026-10-31")
    args = ap.parse_args()
    as_of = pd.Timestamp(args.as_of)

    tape = pd.read_csv(ROOT / "data" / "loan_tape.csv",
                       parse_dates=["auction_date", "funding_date", "maturity_date", "payoff_date"])
    ledger = pd.read_csv(ROOT / "data" / "capital_ledger.csv", parse_dates=["date"])
    ledger.loan_id = ledger.loan_id.fillna("")

    recon, issues = reconcile(tape, ledger, as_of)
    cash = ledger[ledger.date <= as_of].amount.sum()

    active = tape[tape.status == "Active"].copy()
    paid = tape[tape.status == "Paid off"].copy()
    active_principal = active.principal.sum()
    conc = (active.groupby("borrower_id").principal.sum().sort_values(ascending=False) / active_principal)
    paid_days = (paid.payoff_date - paid.funding_date).dt.days
    overdue = active[active.maturity_date < as_of]

    wb = Workbook()

    # ---------------- Summary ----------------
    ws = wb.active
    ws.title = "Summary"
    ws["A1"] = "Loan tape review (demo data)"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = f"As of {as_of:%Y-%m-%d}. All names, amounts and dates are fictitious."
    n_high = (issues.Severity == "High").sum() if len(issues) else 0
    blocks = [
        ("What the information establishes", [
            f"{len(tape)} loans in the tape; {len(active)} active with {active_principal:,.0f} outstanding principal.",
            f"Cash per ledger: {cash:,.0f}.",
            f"{len(paid)} loans paid off; average time from funding to payoff {paid_days.mean():.0f} days.",
            f"Largest borrower exposure: {conc.index[0]} with {conc.iloc[0]:.0%} of active principal.",
        ]),
        ("What is assumed", [
            "Interest is paid at payoff (principal x 12% x term).",
            "Active loans pay off at maturity in the base case; loans already past maturity pay next month.",
            "Delayed case: every payoff moves by the number of months in Assumptions!B3.",
            "New loans: count, size, term and rate as set in Assumptions; they pay off after their term.",
            "Investor distributions and operating costs as set in Assumptions.",
        ]),
        ("What needs to be verified", [
            f"{len(issues)} discrepancies between tape and ledger ({n_high} high severity) - see Issues.",
            f"{len(overdue)} active loans past maturity: extension terms or default status.",
            "Bank statements for the flagged funding and payoff dates.",
        ]),
        ("Recommended next steps", [
            "Resolve high-severity issues before using the forecast with investors.",
            "Agree the delay assumption with the lender, based on its own payoff history.",
            "Use 'Capital needed' in Forecast as the size of the next capital raise, plus the cash buffer.",
        ]),
    ]
    r = 4
    for title, lines in blocks:
        ws.cell(r, 1, title).font = Font(bold=True, color="1F3A5F", size=12)
        r += 1
        for line in lines:
            ws.cell(r, 1, "- " + line)
            r += 1
        r += 1
    ws.cell(r, 1, "Capital needed (base case)").font = Font(bold=True)
    ws.cell(r, 2, "=Forecast!I18").number_format = MONEY
    ws.cell(r + 1, 1, "Capital needed (delayed payoffs)").font = Font(bold=True)
    ws.cell(r + 1, 2, "=Forecast!J18").number_format = MONEY
    widths(ws, [95, 16])

    # ---------------- Issues ----------------
    wi = wb.create_sheet("Issues")
    cols = ["Severity", "Loan", "Issue", "Evidence", "Proposed resolution"]
    header(wi, 1, cols)
    for i, rec in enumerate(issues.to_dict("records"), 2):
        for j, c in enumerate(cols, 1):
            cell = wi.cell(i, j, rec[c])
            cell.alignment = Alignment(wrap_text=True, vertical="top")
        wi.cell(i, 1).fill = PatternFill("solid", fgColor=SEV_FILL[rec["Severity"]])
    widths(wi, [10, 9, 34, 52, 58])
    wi.freeze_panes = "A2"

    # ---------------- Reconciliation ----------------
    wr = wb.create_sheet("Reconciliation")
    cols = list(recon.columns)
    header(wr, 1, cols)
    for i, rec in enumerate(recon.to_dict("records"), 2):
        for j, c in enumerate(cols, 1):
            v = rec[c]
            cell = wr.cell(i, j, None if pd.isna(v) else v)
            if j in (4, 5, 6, 7):
                cell.number_format = MONEY2
        if rec["Result"] != "OK":
            wr.cell(i, len(cols)).fill = PatternFill("solid", fgColor="FCE5CD")
    widths(wr, [9, 10, 13, 16, 16, 15, 15, 9])
    wr.freeze_panes = "A2"

    # ---------------- Portfolio ----------------
    wp = wb.create_sheet("Portfolio")
    wp["A1"] = "Concentration by borrower (active principal)"
    wp["A1"].font = Font(bold=True)
    header(wp, 2, ["Borrower", "Active principal", "Share"])
    for i, (b, share) in enumerate(conc.items(), 3):
        wp.cell(i, 1, b)
        wp.cell(i, 2, float(active[active.borrower_id == b].principal.sum())).number_format = MONEY
        wp.cell(i, 3, float(share)).number_format = "0.0%"
    r = len(conc) + 4
    wp.cell(r, 1, "Loans past maturity").font = Font(bold=True)
    header(wp, r + 1, ["Loan", "Borrower", "Principal", "Maturity", "Days past"])
    for i, ln in enumerate(overdue.itertuples(), r + 2):
        wp.cell(i, 1, ln.loan_id)
        wp.cell(i, 2, ln.borrower_id)
        wp.cell(i, 3, ln.principal).number_format = MONEY
        wp.cell(i, 4, ln.maturity_date.date()).number_format = "yyyy-mm-dd"
        wp.cell(i, 5, (as_of - ln.maturity_date).days)
    r = r + 3 + len(overdue)
    wp.cell(r, 1, "Paid-off loans: days from funding to payoff").font = Font(bold=True)
    wp.cell(r + 1, 1, "Average")
    wp.cell(r + 1, 2, round(float(paid_days.mean()), 0))
    wp.cell(r + 2, 1, "Shortest")
    wp.cell(r + 2, 2, int(paid_days.min()))
    wp.cell(r + 3, 1, "Longest")
    wp.cell(r + 3, 2, int(paid_days.max()))
    widths(wp, [22, 17, 12, 12, 10])

    # ---------------- Assumptions ----------------
    wa = wb.create_sheet("Assumptions")
    header(wa, 1, ["Input", "Value", "Note"])
    inputs = [
        ("As-of date", as_of.date(), "Ledger cash is measured on this date", "yyyy-mm-dd"),
        ("Payoff delay (months)", 2, "Delayed case: months added to every expected payoff", "0"),
        ("New loans per month", 2, "Pipeline assumption - confirm with lender", "0"),
        ("Average new loan", 150000, "Average principal of new loans", MONEY),
        ("Investor distributions per month", 20000, "From ledger history", MONEY),
        ("Operating costs per month", 8000, "From ledger history", MONEY),
        ("Minimum cash buffer", 100000, "Cash the lender wants to keep at all times", MONEY),
        ("Opening cash", float(round(cash, 2)), "Ledger balance on the as-of date", MONEY),
        ("Term of new loans (months)", 6, "New loans pay off after this many months", "0"),
        ("Interest rate on new loans", 0.12, "Annual, paid at payoff", "0%"),
    ]
    for i, (k, v, note, fmt) in enumerate(inputs, 2):
        wa.cell(i, 1, k)
        c = wa.cell(i, 2, v)
        c.number_format = fmt
        c.fill = INPUT
        wa.cell(i, 3, note)
    widths(wa, [32, 16, 52])

    # ---------------- Loans (active, with formulas) ----------------
    wl = wb.create_sheet("Loans")
    header(wl, 1, ["Loan", "Borrower", "Principal", "Maturity", "Expected payoff month",
                   "Delayed payoff month", "Payoff amount"])
    for i, ln in enumerate(active.itertuples(), 2):
        wl.cell(i, 1, ln.loan_id)
        wl.cell(i, 2, ln.borrower_id)
        wl.cell(i, 3, ln.principal).number_format = MONEY
        wl.cell(i, 4, ln.maturity_date.date()).number_format = "yyyy-mm-dd"
        wl.cell(i, 5, f"=IF(D{i}<=Assumptions!$B$2,EOMONTH(Assumptions!$B$2,1),EOMONTH(D{i},0))").number_format = "mmm-yy"
        wl.cell(i, 6, f"=EOMONTH(E{i},Assumptions!$B$3)").number_format = "mmm-yy"
        wl.cell(i, 7, f"=C{i}*(1+{ln.annual_rate}*{ln.term_months}/12)").number_format = MONEY2
    widths(wl, [9, 10, 13, 12, 14, 14, 15])

    # ---------------- Forecast (formulas) ----------------
    wf = wb.create_sheet("Forecast")
    wf["A1"] = "12-month cash forecast: base vs delayed payoffs"
    wf["A1"].font = Font(bold=True, size=12)
    wf["A2"] = "Edit the yellow inputs in Assumptions; every figure below recalculates."
    header(wf, 3, ["Month", "Payoffs, existing loans (base)", "Payoffs, existing loans (delayed)",
                   "New fundings", "Payoffs, new loans (base)", "Payoffs, new loans (delayed)",
                   "Investor distributions", "Operating costs", "Closing cash (base)",
                   "Closing cash (delayed)"])
    wf.row_dimensions[3].height = 45
    last = len(active) + 1
    term, rate, delay = "Assumptions!$B$10", "Assumptions!$B$11", "Assumptions!$B$3"
    for k in range(12):
        r = 4 + k
        wf.cell(r, 1, f"=EOMONTH(Assumptions!$B$2,{k + 1})").number_format = "mmm-yy"
        wf.cell(r, 2, f"=SUMIFS(Loans!$G$2:$G${last},Loans!$E$2:$E${last},A{r})")
        wf.cell(r, 3, f"=SUMIFS(Loans!$G$2:$G${last},Loans!$F$2:$F${last},A{r})")
        wf.cell(r, 4, "=Assumptions!$B$4*Assumptions!$B$5")
        wf.cell(r, 5, f"=IF(ROW()-3>{term},OFFSET(D{r},-{term},0)*(1+{rate}*{term}/12),0)")
        wf.cell(r, 6, f"=IF(ROW()-3>{term}+{delay},OFFSET(D{r},-({term}+{delay}),0)*(1+{rate}*{term}/12),0)")
        wf.cell(r, 7, "=Assumptions!$B$6")
        wf.cell(r, 8, "=Assumptions!$B$7")
        prev_b = "Assumptions!$B$9" if k == 0 else f"I{r - 1}"
        prev_d = "Assumptions!$B$9" if k == 0 else f"J{r - 1}"
        wf.cell(r, 9, f"={prev_b}+B{r}+E{r}-D{r}-G{r}-H{r}")
        wf.cell(r, 10, f"={prev_d}+C{r}+F{r}-D{r}-G{r}-H{r}")
        for c in range(2, 11):
            wf.cell(r, c).number_format = MONEY
    wf.cell(16, 1, "Lowest closing cash").font = Font(bold=True)
    wf.cell(17, 1, "Minimum cash buffer")
    wf.cell(18, 1, "Capital needed").font = Font(bold=True)
    for col in ("I", "J"):
        wf[f"{col}16"] = f"=MIN({col}4:{col}15)"
        wf[f"{col}17"] = "=Assumptions!$B$8"
        c = wf[f"{col}18"]
        c.value = f"=MAX(0,{col}17-{col}16)"
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor="FFF2CC")
        for rr in (16, 17, 18):
            wf[f"{col}{rr}"].number_format = MONEY
    widths(wf, [10, 15, 15, 13, 14, 14, 13, 12, 15, 15])

    chart = LineChart()
    chart.title = "Closing cash: base vs delayed payoffs"
    chart.title.overlay = False  # keep the title above the plot, not on top of it
    chart.x_axis.number_format = "mmm-yy"
    chart.y_axis.number_format = "#,##0"
    # Newer openpyxl hides both axes unless told otherwise; Excel then shows no labels
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.x_axis.tickLblPos = "low"  # month labels stay at the bottom when cash is negative
    chart.legend.position = "b"
    chart.legend.overlay = False  # legend gets its own space below the month labels
    data = Reference(wf, min_col=9, max_col=10, min_row=3, max_row=15)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(Reference(wf, min_col=1, min_row=4, max_row=15))
    for series in chart.series:
        series.smooth = False
    chart.height, chart.width = 10, 20
    wf.add_chart(chart, "A21")

    out = ROOT / "output" / "Loan_Tape_Review.xlsx"
    out.parent.mkdir(exist_ok=True)
    wb.save(out)
    print(f"{len(issues)} issues found. Workbook: {out}")
    print(issues[["Severity", "Loan", "Issue"]].to_string(index=False))


if __name__ == "__main__":
    main()
