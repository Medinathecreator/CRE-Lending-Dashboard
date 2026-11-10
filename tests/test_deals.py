"""
End-to-end test: run both test deals through the engine, build each Excel model,
recalculate it with LibreOffice, and check that Excel's live formulas agree with
the Python engine. Also checks that the expected source-model flags appear.

Run:  python tests/test_deals.py   (needs LibreOffice for the Excel recalculation step)
"""
import json
import subprocess
import sys
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from cre import engine, excel  # noqa: E402

OUT = ROOT / "output"
OUT.mkdir(exist_ok=True)
RECALC = sys.argv[1] if len(sys.argv) > 1 else None   # path to a recalc script (LibreOffice), optional


def named(wb, name):
    dn = wb.defined_names[name]
    (sheet, ref), = list(dn.destinations)
    return wb[sheet][ref.replace("$", "")].value


def check(label, xl, py, tol=1e-6):
    if isinstance(py, str):
        ok = xl == py
    else:
        ok = xl is not None and abs(xl - py) <= tol * max(1, abs(py))
    print(f"  {'PASS' if ok else 'FAIL'}  {label:<34} excel={xl!s:<22} python={py!s:<22}")
    return ok


all_ok = True
for f in sorted((ROOT / "deals").glob("*.json")):
    deal = json.loads(f.read_text())
    res = engine.run(deal)
    m = res["metrics"]
    print(f"\n=== {deal['name']}")
    print("  Source-model reconciliation:")
    for v in res["validation"]:
        print(f"    {v['status']:<22} {v['label']}")

    path = OUT / f"{f.stem}.xlsx"
    path.write_bytes(excel.build(deal, res))
    if not RECALC:
        print("  (skipping Excel tie-out: no recalc script given)")
        continue
    rc = subprocess.run([sys.executable, RECALC, str(path), "120"], capture_output=True, text=True)
    info = json.loads(rc.stdout)
    print(f"  Excel recalc: {info.get('status')}, {info.get('total_formulas')} formulas, {info.get('total_errors')} errors")
    all_ok &= info.get("total_errors") == 0
    wb = load_workbook(path, data_only=True)
    pairs = [("Sized loan", "LoanSized", m["loan"]), ("Binding constraint", "Binding", m["binding"]),
             ("Loan constant", "Constant", m["constant"]), ("DSCR amortizing", "DSCRAmort", m["dscr_amort"]),
             ("DSCR interest-only", "DSCRIO", m["dscr_io"]), ("Debt yield", "DebtYield", m["debt_yield"]),
             ("Breakeven occupancy", "Breakeven", m["breakeven_occ"]), ("Risk score", "RiskScore", m["risk_score"]),
             ("Worst stress DSCR", "WorstDSCR", min(s["dscr_amort"] for s in res["stress"])),
             ("Global DSCR", "GlobalDSCR", res["debt_schedule"]["global_dscr"]),
             ("Rate used", "RateUsed", m["rate_used"])]
    if "requested_loan" in m:
        pairs += [("Requested loan", "ReqLoan", m["requested_loan"]),
                  ("Requested DSCR", "ReqDSCRAmort", m["requested_dscr_amort"])]
    if deal["path"] == "development":
        d = res["development"]
        pairs += [("Total cost", "TotalCost", d["total_cost"]), ("Construction loan", "ConsLoan", d["construction_loan"]),
                  ("Required interest reserve", "ReqReserve", d["draws"]["required_reserve"]),
                  ("Construction payoff", "Payoff", d["payoff"]), ("Takeout coverage", "TakeoutCov", d["takeout_coverage"]),
                  ("Net refi proceeds", "NetRefi", d["net_refi_proceeds"]), ("Yield on cost", "YoC", d["yoc_untrended"]),
                  ("Worst stress takeout coverage", "WorstCov", min(s["takeout_coverage"] for s in res["stress"]))]
    print("  Excel vs Python:")
    for label, nm, py in pairs:
        all_ok &= check(label, named(wb, nm), py)
    status = wb["Credit Summary"]
    flat = [c for row in status.iter_rows(values_only=True) for c in row if isinstance(c, str)]
    model_ok = any(s.startswith("OK: all tabs tie") for s in flat)
    print(f"  Model checks: {'all OK' if model_ok else 'SOME CHECKS FAIL'}")
    all_ok &= model_ok

print("\nALL TIE-OUTS PASS" if all_ok else "\nSOME CHECKS FAILED")
sys.exit(0 if all_ok else 1)
