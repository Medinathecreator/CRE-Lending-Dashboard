"""
Calculation engine: everything is recomputed from the deal inputs.

The Excel generator (excel.py) writes the same logic as live formulas, and
tests/test_deals.py checks that the two agree after Excel recalculates.
"""
from __future__ import annotations

import math

from . import pricing

CONSTRAINTS = ["LTV", "DSCR", "Debt Yield"]


# --------------------------------------------------------------------------- basics
def monthly_payment(loan, rate, years):
    n, r = years * 12, rate / 12
    if loan == 0:
        return 0.0
    return loan / n if r == 0 else loan * r / (1 - (1 + r) ** -n)


def loan_constant(rate, years):
    return monthly_payment(1.0, rate, years) * 12


def income(prop):
    """Gross potential rent, vacancy, other income, opex -> EGI and NOI.
    'noi' mode: NOI is entered directly and an assumed opex ratio (share of EGI) is used
    only to split it into income and expenses for the stress tests and breakeven."""
    vac = prop.get("vacancy_pct", 0.05)
    if prop.get("income_mode", "detail") == "noi":
        noi, ratio = prop["noi"], prop["assumed_opex_ratio"]
        egi = noi / (1 - ratio)
        other = 0.0
        gpr = egi / (1 - vac)
        opex = egi - noi
    else:
        gpr, other, opex = prop["gpr"], prop.get("other_income", 0.0), prop["opex"]
    egi = gpr * (1 - vac) + other
    noi = egi - opex
    return {"gpr": gpr, "other_income": other, "vacancy_pct": vac, "vacancy_loss": gpr * vac,
            "egi": egi, "opex": opex, "noi": noi, "opex_ratio": opex / egi if egi else 0.0}


def size_loan(noi, value, rate, amort_years, max_ltv, min_dscr, min_dy, dscr_basis="amortizing"):
    """Loan = minimum of the LTV, DSCR and debt-yield constraints."""
    const = loan_constant(rate, amort_years)
    ds_factor = const if dscr_basis == "amortizing" else rate
    by = {"LTV": max_ltv * value, "DSCR": noi / min_dscr / ds_factor, "Debt Yield": noi / min_dy}
    loan = min(by.values())
    binding = min(by, key=by.get)
    return {"by_constraint": by, "loan": loan, "binding": binding, "constant": const,
            **loan_metrics(loan, noi, value, rate, amort_years)}


def loan_metrics(loan, noi, value, rate, amort_years):
    const = loan_constant(rate, amort_years)
    ds_amort, ds_io = loan * const, loan * rate
    return {"ltv": loan / value if value else None, "debt_yield": noi / loan if loan else None,
            "ds_amort": ds_amort, "ds_io": ds_io,
            "dscr_amort": noi / ds_amort if ds_amort else None,
            "dscr_io": noi / ds_io if ds_io else None}


def dscr_gap_flag(m, min_dscr):
    """Flag when interest-only and amortizing DSCR tell a materially different story."""
    gap = m["dscr_io"] - m["dscr_amort"]
    if m["dscr_io"] >= min_dscr > m["dscr_amort"]:
        return f"Passes {min_dscr:.2f}x interest-only ({m['dscr_io']:.2f}x) but fails amortizing ({m['dscr_amort']:.2f}x)"
    if gap >= 0.15:
        return f"IO DSCR {m['dscr_io']:.2f}x vs amortizing {m['dscr_amort']:.2f}x: {gap:.2f}x gap, amortization uses much of the cushion"
    return ""


# --------------------------------------------------------------------------- construction
def draw_schedule(dev, cons, loan, noi_stab, sofr):
    """Monthly construction funding: equity first, then loan. Interest accrues on the loan
    balance; lease-up NOI pays interest first, the interest reserve covers the rest."""
    C, L = cons["construction_months"], cons["leaseup_months"]
    T = C + L
    rate = sofr + cons["spread_bps"] / 1e4
    orig = cons["origination_pct"] * loan
    total_cost = dev["land"] + dev["hard"] + dev["soft"] + dev["financing"]
    equity = total_cost - loan
    hs = dev["hard"] + dev["soft"]
    s_curve = cons.get("draw_curve", "straight") == "s-curve"

    rows, eq_left, bal = [], equity, 0.0
    for m in range(T + 1):
        if m == 0:
            cost = dev["land"] + orig
        elif m <= C:
            w = (1 - math.cos(2 * math.pi * (m - 0.5) / C)) / C if s_curve else 1 / C
            cost = hs * w
        else:
            cost = 0.0
        eq = min(cost, eq_left)
        eq_left -= eq
        loan_cost = cost - eq
        beg = bal
        interest = beg * rate / 12 if m > 0 else 0.0
        noi_m = noi_stab / 12 * (m - C) / L if (m > C and cons.get("leaseup_noi_pays_interest", True)) else 0.0
        noi_applied = min(interest, noi_m)
        reserve = interest - noi_applied
        bal = beg + loan_cost + reserve
        rows.append({"month": m, "cost_draw": cost, "equity_draw": eq, "loan_cost_draw": loan_cost,
                     "beg_balance": beg, "interest": interest, "leaseup_noi": noi_m,
                     "noi_applied": noi_applied, "reserve_draw": reserve, "end_balance": bal})
    req = sum(r["reserve_draw"] for r in rows)
    budget = dev["financing"] - orig
    return {"rows": rows, "rate": rate, "origination_fee": orig, "total_interest": sum(r["interest"] for r in rows),
            "interest_paid_by_noi": sum(r["noi_applied"] for r in rows), "required_reserve": req,
            "budgeted_reserve": budget, "reserve_shortfall": max(0.0, req - budget),
            "ending_balance": bal, "payoff": min(bal, loan), "over_commitment": max(0.0, bal - loan)}


# --------------------------------------------------------------------------- stress tests
def scenarios(path):
    s = [("Base", 0, 0, 0, 0, 0), ("Rate +100 bps", 0.01, 0, 0, 0, 0), ("Rate +200 bps", 0.02, 0, 0, 0, 0),
         ("Vacancy +5 pts", 0, 0.05, 0, 0, 0), ("Vacancy +10 pts", 0, 0.10, 0, 0, 0),
         ("Rent -5%", 0, 0, -0.05, 0, 0), ("Rent -10%", 0, 0, -0.10, 0, 0),
         ("Combined: +100 bps, +5 pts vac, -5% rent", 0.01, 0.05, -0.05, 0, 0)]
    if path == "development":
        s += [("Hard/soft cost +10%", 0, 0, 0, 0.10, 0), ("Hard/soft cost +15%", 0, 0, 0, 0.15, 0),
              ("Exit cap +50 bps", 0, 0, 0, 0, 0.005),
              ("Combined + cost +10% + cap +50 bps", 0.01, 0.05, -0.05, 0.10, 0.005)]
    return [dict(zip(["name", "d_rate", "d_vac", "d_rent", "d_cost", "d_cap"], x)) for x in s]


def stress(inc, perm, loan, deal, dev_ctx=None):
    out = []
    for sc in scenarios(deal["path"]):
        gpr = inc["gpr"] * (1 + sc["d_rent"])
        vac = inc["vacancy_pct"] + sc["d_vac"]
        noi = gpr * (1 - vac) + inc["other_income"] - inc["opex"]
        rate = perm["rate_used"] + sc["d_rate"]
        m = loan_metrics(loan, noi, 1.0, rate, perm["amort_years"])
        row = {**sc, "noi": noi, "rate": rate, "dscr_io": m["dscr_io"], "dscr_amort": m["dscr_amort"],
               "debt_yield": m["debt_yield"],
               "breakeven_occ": (inc["opex"] + m["ds_amort"]) / (gpr + inc["other_income"])}
        if dev_ctx:
            d = deal["development"]
            cost = d["land"] + (d["hard"] + d["soft"]) * (1 + sc["d_cost"]) + d["financing"]
            cap = d["exit_cap"] + sc["d_cap"]
            takeout = size_loan(noi, noi / cap, rate, perm["amort_years"], perm["max_ltv"], perm["min_dscr"],
                                perm["min_debt_yield"], perm.get("dscr_basis", "amortizing"))
            row.update(total_cost=cost, yoc=noi / cost, equity_required=cost - dev_ctx["construction_loan"],
                       value=noi / cap, takeout_loan=takeout["loan"],
                       takeout_coverage=takeout["loan"] / dev_ctx["payoff"])
        out.append(row)
    return out


# --------------------------------------------------------------------------- risk scorecard
# (factor, weight perm, weight dev, direction, thresholds for scores 5,4,3,2)
SCORECARD = [
    ("DSCR (amortizing)", 0.20, 0.15, "higher", [1.50, 1.35, 1.25, 1.15]),
    ("LTV / LTC", 0.15, 0.15, "lower", [0.55, 0.60, 0.65, 0.75]),
    ("Debt yield", 0.15, 0.10, "higher", [0.11, 0.10, 0.09, 0.08]),
    ("Breakeven occupancy", 0.10, 0.05, "lower", [0.70, 0.75, 0.80, 0.85]),
    ("Worst stress DSCR", 0.10, 0.10, "higher", [1.25, 1.15, 1.05, 1.00]),
    ("Takeout coverage (worst stress)", 0.00, 0.15, "higher", [1.30, 1.15, 1.05, 1.00]),
    ("Credit score", 0.10, 0.10, "higher", [760, 720, 680, 640]),
    ("Liquidity / loan", 0.10, 0.10, "higher", [0.15, 0.10, 0.075, 0.05]),
    ("Net worth / loan", 0.05, 0.05, "higher", [1.50, 1.00, 0.75, 0.50]),
    ("Experience (deals completed)", 0.05, 0.05, "higher", [10, 5, 3, 1]),
]
GRADES = [(4.25, "A - Low risk"), (3.50, "B - Moderate"), (2.75, "C - Watch"), (2.00, "D - Elevated"), (0, "E - High risk")]


def factor_score(x, direction, t):
    for score, thr in zip([5, 4, 3, 2], t):
        if (x >= thr) if direction == "higher" else (x <= thr):
            return score
    return 1


def scorecard(values, path):
    rows, total = [], 0.0
    for name, wp, wd, direction, t in SCORECARD:
        w = wd if path == "development" else wp
        x = values[name]
        s = factor_score(x, direction, t)
        total += w * s
        rows.append({"factor": name, "value": x, "weight": w, "score": s, "weighted": w * s,
                     "direction": direction, "thresholds": t})
    grade = next(g for cut, g in GRADES if total >= cut)
    return {"rows": rows, "score": total, "grade": grade}


# --------------------------------------------------------------------------- personal debt schedule
def debt_schedule(debts, borrower, prop_loan, prop_rate, amort_years, noi, equity_required, recourse_pct,
                  guaranty_base=None, years=10):
    pay = monthly_payment(prop_loan, prop_rate, amort_years)
    personal_annual = sum(d["monthly_payment"] * 12 for d in debts)
    prop_annual = pay * 12
    inc = borrower["annual_income"]
    proj = []
    for y in range(years + 1):
        row = {"year": y}
        for d in debts:
            r, n, b, p = d["rate"] / 12, 12 * y, d["balance"], d["monthly_payment"]
            fv = b * (1 + r) ** n - p * (((1 + r) ** n - 1) / r if r else n)
            row[d["name"]] = max(0.0, fv)
        r, n = prop_rate / 12, 12 * y
        row["New property loan"] = max(0.0, prop_loan * (1 + r) ** n - pay * (((1 + r) ** n - 1) / r if r else n))
        proj.append(row)
    liq_after = borrower["liquidity"] - equity_required
    return {
        "personal_annual_payments": personal_annual,
        "property_annual_ds": prop_annual,
        "dti_before": personal_annual / inc if inc else None,
        "global_dscr": (inc + noi) / (personal_annual + prop_annual) if (personal_annual + prop_annual) else None,
        "post_closing_liquidity": liq_after,
        "reserve_months": liq_after / pay if pay else None,
        "contingent_liability": recourse_pct * (guaranty_base if guaranty_base is not None else prop_loan),
        "projection": proj,
    }


# --------------------------------------------------------------------------- main
def run(deal, benchmarks=None):
    benchmarks = benchmarks or pricing.FALLBACK_BENCHMARKS
    prop, perm, bor = deal["property"], dict(deal["perm"]), deal["borrower"]
    path = deal["path"]
    inc = income(prop)
    noi = inc["noi"]
    res = {"path": path, "income": inc, "warnings": []}

    # pricing: leverage band from requested (or max) LTV
    lev = perm.get("requested_ltv") or perm["max_ltv"]
    res["pricing_perm"] = pricing.indicative_rates(benchmarks, lev, prop["type"])
    sel = next(r for r in res["pricing_perm"] if r["lender_type"] == perm["lender_type"])
    perm["indicative_rate"] = sel["rate"]
    perm["rate_used"] = perm["rate"] if perm.get("rate") is not None else sel["rate"]
    if perm["rate_used"] is None:
        raise ValueError(f"{perm['lender_type']} has no spread at {lev:.0%} leverage; enter a rate override.")
    res["perm"] = perm
    basis = perm.get("dscr_basis", "amortizing")

    if path == "permanent":
        value = prop.get("appraised_value") or prop["purchase_price"]
        sizing = size_loan(noi, value, perm["rate_used"], perm["amort_years"], perm["max_ltv"],
                           perm["min_dscr"], perm["min_debt_yield"], basis)
        closing = prop.get("closing_cost_pct", 0.0) * prop["purchase_price"]
        uses = prop["purchase_price"] + closing
        su = {"uses": {"Purchase price": prop["purchase_price"], "Closing costs": closing}, "total_uses": uses,
              "sources": {"Senior loan": sizing["loan"], "Sponsor equity": uses - sizing["loan"]}}
        equity_required = uses - sizing["loan"]
        prop_loan = sizing["loan"]
        dev_ctx = None
    else:
        d, cons = deal["development"], deal["construction"]
        total_cost = d["land"] + d["hard"] + d["soft"] + d["financing"]
        c_loan = cons.get("loan_amount") or cons["ltc_max"] * total_cost
        sofr = benchmarks["SOFR"]["value"] if cons.get("sofr") is None else cons["sofr"]
        draws = draw_schedule(d, cons, c_loan, noi, sofr)
        value = noi / d["exit_cap"]
        sizing = size_loan(noi, value, perm["rate_used"], perm["amort_years"], perm["max_ltv"],
                           perm["min_dscr"], perm["min_debt_yield"], basis)
        perm_costs = perm.get("closing_cost_pct", 0.0) * sizing["loan"]
        noi_tr = d.get("noi_trended")
        trended = None
        if noi_tr:
            trended = size_loan(noi_tr, noi_tr / d["exit_cap"], perm["rate_used"], perm["amort_years"],
                                perm["max_ltv"], perm["min_dscr"], perm["min_debt_yield"], basis)
        dev_ctx = {"construction_loan": c_loan, "payoff": draws["payoff"]}
        res["development"] = {
            "total_cost": total_cost, "construction_loan": c_loan, "ltc": c_loan / total_cost,
            "equity": total_cost - c_loan, "equity_pct": (total_cost - c_loan) / total_cost,
            "sofr": sofr, "construction_rate": draws["rate"], "draws": draws,
            "yoc_untrended": noi / total_cost, "yoc_trended": noi_tr / total_cost if noi_tr else None,
            "dev_spread_bps": (noi / total_cost - d["exit_cap"]) * 1e4,
            "construction_debt_yield": noi / c_loan,
            "stabilized_value": value, "trended_value": noi_tr / d["exit_cap"] if noi_tr else None,
            "takeout_loan": sizing["loan"], "payoff": draws["payoff"],
            "takeout_coverage": sizing["loan"] / draws["payoff"],
            "gross_refi_excess": sizing["loan"] - draws["payoff"],
            "net_refi_proceeds": sizing["loan"] - draws["payoff"] - perm_costs,
            "takeout_pass": sizing["loan"] >= draws["payoff"],
            "trended_takeout": trended,
        }
        su = {"uses": {"Land": d["land"], "Hard costs": d["hard"], "Soft costs": d["soft"],
                       "Financing (origination + interest reserve)": d["financing"]}, "total_uses": total_cost,
              "sources": {"Construction loan": c_loan, "Sponsor equity": total_cost - c_loan}}
        equity_required = total_cost - c_loan
        prop_loan = sizing["loan"]
        if draws["reserve_shortfall"] > 0:
            res["warnings"].append(
                f"Interest reserve short by ${draws['reserve_shortfall']:,.0f}: required ${draws['required_reserve']:,.0f} "
                f"vs ${draws['budgeted_reserve']:,.0f} budgeted (financing line less origination fee).")
        if not res["development"]["takeout_pass"]:
            res["warnings"].append("Takeout test fails: permanent loan does not cover the construction payoff.")
        res["pricing_construction"] = pricing.indicative_rates(benchmarks, c_loan / total_cost, prop["type"], "construction")

    res["value"] = value
    res["sizing"] = sizing
    res["sources_uses"] = su
    res["dscr_flag"] = dscr_gap_flag(sizing, perm["min_dscr"])
    if res["dscr_flag"]:
        res["warnings"].append("DSCR: " + res["dscr_flag"])
    res["breakeven_occ"] = (inc["opex"] + sizing["ds_amort"]) / (inc["gpr"] + inc["other_income"])

    if perm.get("requested_ltv"):
        rl = perm["requested_ltv"] * value
        rm = loan_metrics(rl, noi, value, perm["rate_used"], perm["amort_years"])
        dscr_used = rm["dscr_amort"] if basis == "amortizing" else rm["dscr_io"]
        fails = [n for n, ok in [("LTV", perm["requested_ltv"] <= perm["max_ltv"] + 1e-9),
                                 ("DSCR", dscr_used >= perm["min_dscr"] - 1e-9),
                                 ("Debt Yield", rm["debt_yield"] >= perm["min_debt_yield"] - 1e-9)] if not ok]
        res["requested"] = {"loan": rl, **rm, "fails": fails}
        if fails:
            res["warnings"].append(f"Requested {perm['requested_ltv']:.0%} LTV loan (${rl:,.0f}) fails: {', '.join(fails)}. "
                                   f"Loan sizes down to ${sizing['loan']:,.0f} ({sizing['binding']} binds).")

    res["stress"] = stress(inc, perm, sizing["loan"], deal, dev_ctx)
    worst_dscr = min(r["dscr_amort"] for r in res["stress"])
    worst_cov = min(r["takeout_coverage"] for r in res["stress"]) if dev_ctx else 0.0
    lev_metric = res["development"]["ltc"] if dev_ctx else sizing["ltv"]
    ref_loan = res["development"]["construction_loan"] if dev_ctx else sizing["loan"]
    res["scorecard"] = scorecard({
        "DSCR (amortizing)": sizing["dscr_amort"], "LTV / LTC": lev_metric, "Debt yield": sizing["debt_yield"],
        "Breakeven occupancy": res["breakeven_occ"], "Worst stress DSCR": worst_dscr,
        "Takeout coverage (worst stress)": worst_cov, "Credit score": bor["credit_score"],
        "Liquidity / loan": bor["liquidity"] / ref_loan, "Net worth / loan": bor["net_worth"] / ref_loan,
        "Experience (deals completed)": bor["deals_completed"]}, path)

    sponsor_equity = equity_required * bor.get("sponsor_equity_share", 1.0)   # rest comes from LP investors
    res["debt_schedule"] = debt_schedule(deal.get("personal_debts", []), bor, prop_loan, perm["rate_used"],
                                         perm["amort_years"], noi, sponsor_equity, bor.get("recourse_pct", 0.0),
                                         guaranty_base=ref_loan)
    res["equity_required"] = equity_required
    res["metrics"] = flat_metrics(res)
    res["validation"] = validate(deal.get("expected", []), res["metrics"])
    return res


def flat_metrics(res):
    s = res["sizing"]
    m = {"noi": res["income"]["noi"], "value": res["value"], "loan": s["loan"], "binding": s["binding"],
         "ltv": s["ltv"], "debt_yield": s["debt_yield"], "constant": s["constant"],
         "dscr_io": s["dscr_io"], "dscr_amort": s["dscr_amort"], "breakeven_occ": res["breakeven_occ"],
         "equity_required": res["equity_required"], "risk_score": res["scorecard"]["score"],
         "rate_used": res["perm"]["rate_used"]}
    if "requested" in res:
        r = res["requested"]
        m.update(requested_loan=r["loan"], requested_dscr_amort=r["dscr_amort"], requested_debt_yield=r["debt_yield"])
    if "development" in res:
        d = res["development"]
        m.update({k: d[k] for k in ["total_cost", "construction_loan", "ltc", "equity", "equity_pct", "yoc_untrended",
                                    "yoc_trended", "takeout_coverage", "net_refi_proceeds", "gross_refi_excess",
                                    "stabilized_value", "construction_rate"]})
        m["required_interest_reserve"] = d["draws"]["required_reserve"]
    return m


def validate(expected, metrics):
    """Compare numbers from the user's source/test model with the engine's own result.
    Never adjusts the engine to match: a difference is reported, not hidden."""
    out = []
    for e in expected:
        calc = metrics.get(e["key"])
        if calc is None:
            out.append({**e, "calc": None, "diff": None, "status": "Not computed"})
            continue
        if isinstance(e["value"], str):
            ok = str(calc) == e["value"]
            out.append({**e, "calc": calc, "diff": None, "status": "Ties" if ok else "MISMATCH"})
            continue
        diff = calc - e["value"]
        tol = e.get("tol", abs(e["value"]) * 0.001)
        if abs(diff) <= tol:
            status = "Ties"
        elif abs(diff) <= abs(e["value"]) * 0.01:
            status = "Minor mismatch (<1%)"
        else:
            status = "MISMATCH"
        out.append({**e, "calc": calc, "diff": diff, "status": status})
    return out
