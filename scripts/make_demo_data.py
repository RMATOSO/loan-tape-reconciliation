"""
Builds a fictitious loan tape and capital ledger for a small short-term lender
that finances investors buying properties at foreclosure auctions.

Several problems are planted on purpose so the reconciliation has something to find.
All names, amounts and dates are invented.
"""
import csv
import datetime as dt
import random
from pathlib import Path

random.seed(42)
ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / "data"
DATA.mkdir(exist_ok=True)

AS_OF = dt.date(2026, 10, 31)
RATE = 0.12  # annual interest, paid at payoff
BORROWERS = ["B01", "B02", "B03", "B03", "B03", "B04", "B05", "B06", "B03", "B07",
             "B08", "B02", "B09", "B03", "B10", "B11", "B12", "B02", "B13", "B14",
             "B15", "B16", "B17", "B18"]
COUNTIES = ["Maple County", "Cedar County", "Pine County", "Oak County"]


def add_months(d, m):
    y, mo = divmod(d.month - 1 + m, 12)
    y += d.year
    mo += 1
    last = [31, 29 if y % 4 == 0 else 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][mo - 1]
    return dt.date(y, mo, min(d.day, last))


loans = []
start = dt.date(2025, 11, 5)
for i in range(24):
    loan_id = f"L-{101 + i:04d}"
    funded = start + dt.timedelta(days=int(i * 14.5))
    term = 6 if i % 3 else 12
    principal = random.choice([85000, 110000, 125000, 140000, 160000, 185000, 210000])
    maturity = add_months(funded, term)
    payoff_amt = round(principal * (1 + RATE * term / 12), 2)
    status, payoff_date, paid = "Active", "", ""
    if maturity <= AS_OF and i not in (4, 10):
        status = "Paid off"
        payoff_date = (maturity - dt.timedelta(days=random.randint(0, 40))).isoformat()
        days = (dt.date.fromisoformat(payoff_date) - funded).days
        paid = round(principal * (1 + RATE * days / 365), 2)
    loans.append({
        "loan_id": loan_id, "borrower_id": BORROWERS[i], "county": COUNTIES[i % 4],
        "auction_date": (funded - dt.timedelta(days=3)).isoformat(),
        "funding_date": funded.isoformat(), "principal": principal,
        "annual_rate": RATE, "term_months": term, "maturity_date": maturity.isoformat(),
        "status": status, "payoff_date": payoff_date, "payoff_amount": paid,
    })

# Ledger built from the tape, then planted problems
ledger = [{"date": "2025-10-15", "type": "Investor capital in", "loan_id": "", "amount": 2_400_000},
          {"date": "2026-05-29", "type": "Investor capital in", "loan_id": "", "amount": 900_000}]
for ln in loans:
    ledger.append({"date": ln["funding_date"], "type": "Funding", "loan_id": ln["loan_id"],
                   "amount": -ln["principal"]})
    if ln["status"] == "Paid off":
        ledger.append({"date": ln["payoff_date"], "type": "Payoff", "loan_id": ln["loan_id"],
                       "amount": ln["payoff_amount"]})
m = dt.date(2025, 11, 30)
while m <= AS_OF:
    ledger.append({"date": m.isoformat(), "type": "Investor distribution", "loan_id": "", "amount": -20000})
    ledger.append({"date": m.isoformat(), "type": "Operating costs", "loan_id": "", "amount": -8000})
    m = add_months(m.replace(day=1), 2) - dt.timedelta(days=1)

def find(loan_id, typ):
    return next(r for r in ledger if r["loan_id"] == loan_id and r["type"] == typ)

# 1. Funding in the tape but missing from the ledger
ledger.remove(find("L-0107", "Funding"))
# 2. Payoff amount differs from the tape
find("L-0112", "Payoff")["amount"] = round(find("L-0112", "Payoff")["amount"] - 2500, 2)
# 3. Funding in the ledger for a loan that is not in the tape
ledger.append({"date": "2026-09-18", "type": "Funding", "loan_id": "L-0125", "amount": -95000})
# 4. Funding date in the ledger 6 days after the tape
r = find("L-0103", "Funding")
r["date"] = (dt.date.fromisoformat(r["date"]) + dt.timedelta(days=6)).isoformat()
# 5. Duplicate funding row
ledger.append(dict(find("L-0115", "Funding")))
# 6. Tape says paid off but the ledger has no payoff
ledger.remove(find("L-0109", "Payoff"))

ledger.sort(key=lambda r: r["date"])

with open(DATA / "loan_tape.csv", "w", newline="") as f:
    w = csv.DictWriter(f, fieldnames=list(loans[0].keys()))
    w.writeheader()
    w.writerows(loans)
with open(DATA / "capital_ledger.csv", "w", newline="") as f:
    w = csv.DictWriter(f, fieldnames=["date", "type", "loan_id", "amount"])
    w.writeheader()
    w.writerows(ledger)
print(f"{len(loans)} loans, {len(ledger)} ledger rows written to {DATA}")
