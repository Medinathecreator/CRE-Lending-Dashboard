"""
CRE Credit Dashboard (Streamlit)
Enter a property + borrower -> loan sizing, stress tests, risk score, indicative pricing,
personal debt schedule, and a downloadable Excel model with live formulas.

Run on Replit:  streamlit run app.py --server.port 8080 --server.address 0.0.0.0
"""
import copy
import json
from pathlib import Path

import altair as alt
import pandas as pd
import streamlit as st

from cre import engine, excel, pricing

ROOT = Path(__file__).parent
DEALS = {
    "Test Deal 1: Chicago development": ROOT / "deals" / "deal1_chicago_development.json",
    "Test Deal 2: Nashville acquisition": ROOT / "deals" / "deal2_nashville_acquisition.json",
    "New deal: acquisition / refi (starter values)": ROOT / "deals" / "starter_permanent.json",
    "New deal: development (starter values)": ROOT / "deals" / "starter_development.json",
}
PROPERTY_TYPES = ["Multifamily", "Mixed-use", "Office", "Retail", "Industrial", "Hotel", "Self-storage"]
ACCENT, GREY = "#2a78d6", "#b9b8b3"

st.set_page_config(page_title="CRE Credit Dashboard", layout="wide")


def usd(x):
    return "-" if x is None else f"${x:,.0f}"


def pct(x, d=2):
    return "-" if x is None else f"{x:.{d}%}"


def xx(x):
    return "-" if x is None else f"{x:.2f}x"


@st.cache_data(ttl=3600, show_spinner="Fetching SOFR and Treasury yields from FRED...")
def get_benchmarks():
    return pricing.fetch_benchmarks()


# ------------------------------------------------------------------ sidebar
st.sidebar.title("CRE Credit Dashboard")
choice = st.sidebar.selectbox("Deal", list(DEALS))
if st.session_state.get("deal_choice") != choice:
    st.session_state["deal_choice"] = choice
    st.session_state["base"] = json.loads(DEALS[choice].read_text())
base = st.session_state["base"]
kid = str(list(DEALS).index(choice))  # widget key prefix so inputs reset when you switch deals

use_fred = st.sidebar.toggle("Live benchmarks from FRED", value=True,
                             help="Uses FRED_API_KEY from Replit Secrets if set, otherwise FRED's public CSV. Falls back to stored values.")
if st.sidebar.button("Refresh benchmarks"):
    get_benchmarks.clear()
bench = get_benchmarks() if use_fred else copy.deepcopy(pricing.FALLBACK_BENCHMARKS)
for k, b in bench.items():
    st.sidebar.caption(f"{k}: {b['value']:.2%}  ({b['as_of']}, {b['source']})")
st.sidebar.divider()
st.sidebar.caption("Spreads by lender type are ASSUMPTIONS from public market reports (see Pricing tab), not quotes.")

# ------------------------------------------------------------------ inputs
st.title(base["name"])
if base.get("assumptions_note"):
    st.caption("Assumed inputs: " + base["assumptions_note"])

deal = copy.deepcopy(base)
P, L, Bw = deal["property"], deal["perm"], deal["borrower"]

with st.expander("Inputs", expanded=False):
    t_prop, t_loan, t_bor, t_debts = st.tabs(["Property", "Loan terms", "Borrower", "Personal debts"])

    with t_prop:
        c1, c2, c3 = st.columns(3)
        deal["path"] = c1.radio("Loan path", ["permanent", "development"],
                                index=0 if base["path"] == "permanent" else 1, key=f"{kid}path",
                                format_func=lambda s: "A. Stabilized acquisition / refi" if s == "permanent" else "B. Ground-up development")
        P["type"] = c2.selectbox("Property type", PROPERTY_TYPES,
                                 index=PROPERTY_TYPES.index(P.get("type", "Multifamily")), key=f"{kid}type")
        P["market"] = c3.text_input("Market", P.get("market", ""), key=f"{kid}mkt")
        P["units"] = c3.number_input("Units", 0, 10000, int(P.get("units") or 0), key=f"{kid}units")

        mode = c1.radio("Income input", ["detail", "noi"], index=0 if P.get("income_mode") == "detail" else 1,
                        key=f"{kid}mode", format_func=lambda s: "Rent roll / gross income" if s == "detail" else "NOI only")
        P["income_mode"] = mode
        if mode == "detail":
            st.markdown("**Rent roll** (monthly rents). Gross potential rent = sum of units x rent x 12.")
            rr = pd.DataFrame(P.get("rent_roll") or [{"Unit type": "1BR", "Units": int(P.get("units") or 0), "Monthly rent": 0.0}])
            rr = st.data_editor(rr, num_rows="dynamic", key=f"{kid}rr", use_container_width=True)
            gpr_rr = float((rr["Units"].fillna(0) * rr["Monthly rent"].fillna(0)).sum() * 12)
            P["rent_roll"] = rr.fillna(0).to_dict("records")
            a, b, c, d = st.columns(4)
            if gpr_rr > 0:
                P["gpr"] = gpr_rr
                a.metric("Gross potential rent (from rent roll)", usd(gpr_rr))
            else:
                P["gpr"] = a.number_input("Gross potential rent (annual)", 0.0, value=float(P.get("gpr") or 0), step=10000.0,
                                          key=f"{kid}gpr", help="Fill in the rent roll above instead to build this from units x rent")
            P["other_income"] = b.number_input("Other income (annual)", 0.0, value=float(P.get("other_income") or 0), step=5000.0, key=f"{kid}oi")
            P["vacancy_pct"] = c.number_input("Vacancy & credit loss", 0.0, 0.9, float(P.get("vacancy_pct", 0.05)), 0.005, key=f"{kid}vac", format="%.3f")
            P["opex"] = d.number_input("Operating expenses (annual)", 0.0, value=float(P.get("opex") or 0), step=10000.0, key=f"{kid}opex")
        else:
            a, b, c = st.columns(3)
            P["noi"] = a.number_input("Stabilized NOI (annual)", 0.0, value=float(P.get("noi") or 0), step=10000.0, key=f"{kid}noi")
            P["vacancy_pct"] = b.number_input("Vacancy & credit loss", 0.0, 0.9, float(P.get("vacancy_pct", 0.05)), 0.005, key=f"{kid}vac2", format="%.3f")
            P["assumed_opex_ratio"] = c.number_input("Assumed opex ratio (share of EGI)", 0.05, 0.9,
                                                     float(P.get("assumed_opex_ratio", 0.40)), 0.01, key=f"{kid}oxr",
                                                     help="Only used to split NOI into income and expenses for stress tests and breakeven.")

        if deal["path"] == "permanent":
            a, b, c = st.columns(3)
            P["purchase_price"] = a.number_input("Purchase price (or value for a refi)", 0.0, value=float(P.get("purchase_price") or 0), step=100000.0, key=f"{kid}pp")
            P["closing_cost_pct"] = b.number_input("Closing costs (% of price)", 0.0, 0.2, float(P.get("closing_cost_pct", 0.015)), 0.0025, key=f"{kid}cc", format="%.4f")
        else:
            D = deal.setdefault("development", {"land": 0, "hard": 0, "soft": 0, "financing": 0, "noi_trended": 0, "exit_cap": 0.055})
            Cn = deal.setdefault("construction", {"ltc_max": 0.60, "spread_bps": 300, "origination_pct": 0.01, "term_months": 36,
                                                  "construction_months": 24, "leaseup_months": 12, "draw_curve": "straight",
                                                  "leaseup_noi_pays_interest": True, "sofr": None, "lender_type": "Bank construction"})
            st.markdown("**Development budget**")
            a, b, c, d = st.columns(4)
            D["land"] = a.number_input("Land", 0.0, value=float(D["land"]), step=100000.0, key=f"{kid}land")
            D["hard"] = b.number_input("Hard costs", 0.0, value=float(D["hard"]), step=100000.0, key=f"{kid}hard")
            D["soft"] = c.number_input("Soft costs", 0.0, value=float(D["soft"]), step=100000.0, key=f"{kid}soft")
            D["financing"] = d.number_input("Financing (origination + interest reserve)", 0.0, value=float(D["financing"]), step=50000.0, key=f"{kid}fin")
            a, b = st.columns(2)
            D["noi_trended"] = a.number_input("Trended stabilized NOI (optional)", 0.0, value=float(D.get("noi_trended") or 0), step=10000.0, key=f"{kid}noitr")
            D["exit_cap"] = b.number_input("Exit / valuation cap rate", 0.02, 0.15, float(D["exit_cap"]), 0.0025, key=f"{kid}cap", format="%.4f")
            st.markdown("**Construction loan**")
            a, b, c, d = st.columns(4)
            Cn["ltc_max"] = a.number_input("Max LTC", 0.1, 0.95, float(Cn["ltc_max"]), 0.01, key=f"{kid}ltc")
            Cn["spread_bps"] = b.number_input("Spread over SOFR (bps)", 0, 1500, int(Cn["spread_bps"]), 5, key=f"{kid}spr")
            sofr_default = Cn.get("sofr") if Cn.get("sofr") is not None else bench["SOFR"]["value"]
            Cn["sofr"] = c.number_input("SOFR assumption", 0.0, 0.15, float(sofr_default), 0.0005, key=f"{kid}sofr", format="%.4f",
                                        help=f"Latest SOFR: {bench['SOFR']['value']:.2%} ({bench['SOFR']['as_of']})")
            Cn["origination_pct"] = d.number_input("Origination fee", 0.0, 0.05, float(Cn["origination_pct"]), 0.0025, key=f"{kid}orig", format="%.4f")
            a, b, c, d = st.columns(4)
            Cn["construction_months"] = a.number_input("Construction months", 1, 72, int(Cn["construction_months"]), key=f"{kid}cm")
            Cn["leaseup_months"] = b.number_input("Lease-up months", 1, 48, int(Cn["leaseup_months"]), key=f"{kid}lm")
            Cn["term_months"] = c.number_input("Loan term (months)", 1, 96, int(Cn["term_months"]), key=f"{kid}tm")
            Cn["draw_curve"] = d.selectbox("Draw curve", ["straight", "s-curve"], index=["straight", "s-curve"].index(Cn.get("draw_curve", "straight")), key=f"{kid}dc")
            Cn["leaseup_noi_pays_interest"] = st.checkbox("Lease-up NOI pays interest before the reserve", Cn.get("leaseup_noi_pays_interest", True), key=f"{kid}lnpi")

    with t_loan:
        names = list(pricing.PERM_LENDERS)
        a, b, c = st.columns(3)
        L["lender_type"] = a.selectbox("Permanent lender type", names, index=names.index(L["lender_type"]), key=f"{kid}lt")
        use_override = b.checkbox("Override rate", L.get("rate") is not None, key=f"{kid}ovr")
        L["rate"] = c.number_input("Rate", 0.0, 0.2, float(L.get("rate") or 0.065), 0.0005, key=f"{kid}rate", format="%.4f") if use_override else None
        a, b, c, d = st.columns(4)
        L["amort_years"] = a.number_input("Amortization (years)", 5, 40, int(L["amort_years"]), key=f"{kid}am")
        L["max_ltv"] = b.number_input("Max LTV", 0.1, 0.95, float(L["max_ltv"]), 0.01, key=f"{kid}mltv")
        L["min_dscr"] = c.slider("Min DSCR", 1.00, 1.60, float(L["min_dscr"]), 0.05, key=f"{kid}mdscr")
        L["min_debt_yield"] = d.number_input("Min debt yield", 0.0, 0.2, float(L["min_debt_yield"]), 0.0025, key=f"{kid}mdy", format="%.4f")
        a, b, c = st.columns(3)
        L["dscr_basis"] = a.selectbox("Size DSCR on", ["amortizing", "io"], index=["amortizing", "io"].index(L.get("dscr_basis", "amortizing")), key=f"{kid}basis",
                                      format_func=lambda s: "Amortizing payment" if s == "amortizing" else "Interest-only payment")
        req = b.number_input("Requested LTV (0 = none)", 0.0, 0.95, float(L.get("requested_ltv") or 0), 0.01, key=f"{kid}req")
        L["requested_ltv"] = req or None
        L["closing_cost_pct"] = c.number_input("Perm closing costs (% of loan)", 0.0, 0.05, float(L.get("closing_cost_pct", 0)), 0.0025, key=f"{kid}pcc", format="%.4f")

    with t_bor:
        a, b, c, d = st.columns(4)
        Bw["credit_score"] = a.number_input("Credit score", 300, 850, int(Bw["credit_score"]), key=f"{kid}fico")
        Bw["liquidity"] = b.number_input("Liquidity", 0.0, value=float(Bw["liquidity"]), step=50000.0, key=f"{kid}liq")
        Bw["net_worth"] = c.number_input("Net worth", 0.0, value=float(Bw["net_worth"]), step=100000.0, key=f"{kid}nw")
        Bw["existing_debt"] = d.number_input("Existing debt", 0.0, value=float(Bw.get("existing_debt", 0)), step=100000.0, key=f"{kid}ed")
        a, b, c, d = st.columns(4)
        Bw["deals_completed"] = a.number_input("Deals completed", 0, 500, int(Bw["deals_completed"]), key=f"{kid}deals")
        Bw["experience_years"] = b.number_input("Years of experience", 0, 60, int(Bw.get("experience_years", 0)), key=f"{kid}yrs")
        Bw["annual_income"] = c.number_input("Annual personal income", 0.0, value=float(Bw["annual_income"]), step=5000.0, key=f"{kid}inc")
        Bw["recourse_pct"] = d.number_input("Recourse / guaranty share", 0.0, 1.0, float(Bw.get("recourse_pct", 0)), 0.05, key=f"{kid}rec")
        Bw["sponsor_equity_share"] = a.number_input("Sponsor's share of equity", 0.0, 1.0, float(Bw.get("sponsor_equity_share", 1.0)), 0.05, key=f"{kid}seq",
                                                    help="1.0 = you fund all the equity. Lower if LP investors fund part of it.")

    with t_debts:
        st.caption("Sample debts are placeholders. Replace them with yours; nothing here is saved anywhere.")
        dd = pd.DataFrame(deal.get("personal_debts", []), columns=["name", "balance", "rate", "monthly_payment"])
        dd = st.data_editor(dd, num_rows="dynamic", key=f"{kid}debts", use_container_width=True,
                            column_config={"rate": st.column_config.NumberColumn(format="%.4f")})
        deal["personal_debts"] = [r for r in dd.fillna(0).to_dict("records") if str(r.get("name", "")).strip()]

if deal["path"] == "permanent":
    deal.pop("development", None)
    deal.pop("construction", None)
if deal["path"] != base["path"]:
    deal["expected"] = []   # source-model figures belong to the original loan path

# ------------------------------------------------------------------ input sanity checks
problems = []
inc_check = engine.income(P) if (P.get("income_mode") == "detail" and P.get("gpr", 0) > 0) or \
    (P.get("income_mode") == "noi" and P.get("noi", 0) > 0) else None
if inc_check is None or inc_check["noi"] <= 0:
    problems.append("NOI must be positive: enter a rent roll / gross income and expenses, or an NOI.")
if deal["path"] == "permanent" and not P.get("purchase_price"):
    problems.append("Enter a purchase price (or the property value for a refinance).")
if deal["path"] == "development" and sum(deal["development"][k_] for k_ in ("land", "hard", "soft")) <= 0:
    problems.append("Enter the development budget (land, hard and soft costs).")
if problems:
    for p_ in problems:
        st.info(p_)
    st.stop()

# ------------------------------------------------------------------ run
try:
    res = engine.run(deal, bench)
except Exception as e:  # show input problems instead of a stack trace
    st.error(f"Can't size this deal yet: {e}")
    st.stop()

dev = "development" in res
s = res["sizing"]
k = st.columns(6)
k[0].metric("Permanent loan" if not dev else "Takeout loan", usd(s["loan"]), f"{s['binding']} binds", delta_color="off")
if dev:
    k[1].metric("Construction loan", usd(res["development"]["construction_loan"]), f"{res['development']['ltc']:.1%} LTC", delta_color="off")
else:
    k[1].metric("LTV", pct(s["ltv"], 1))
k[2].metric("DSCR amortizing / IO", f"{s['dscr_amort']:.2f}x / {s['dscr_io']:.2f}x")
k[3].metric("Debt yield", pct(s["debt_yield"]))
k[4].metric("Breakeven occupancy", pct(res["breakeven_occ"], 1))
k[5].metric("Risk score", f"{res['scorecard']['score']:.2f}", res["scorecard"]["grade"], delta_color="off")
for w in res["warnings"]:
    st.warning(w)

tabs = ["Loan sizing", "Risk", "Pricing"] + (["Development"] if dev else []) + ["Debt schedule", "Validation", "Excel model"]
T = dict(zip(tabs, st.tabs(tabs)))

# ------------------------------------------------------------------ loan sizing
with T["Loan sizing"]:
    c1, c2 = st.columns([3, 2])
    by = pd.DataFrame({"Constraint": list(s["by_constraint"]), "Max loan": list(s["by_constraint"].values())})
    by["Binds"] = by["Constraint"] == s["binding"]
    chart = alt.Chart(by).mark_bar(cornerRadiusEnd=4, size=28).encode(
        y=alt.Y("Constraint:N", sort=None, title=None),
        x=alt.X("Max loan:Q", title="Maximum loan ($)", axis=alt.Axis(format="$,.0f", grid=True)),
        color=alt.condition("datum.Binds", alt.value(ACCENT), alt.value(GREY)),
        tooltip=[alt.Tooltip("Constraint:N"), alt.Tooltip("Max loan:Q", format="$,.0f")])
    labels = chart.mark_text(align="left", dx=6, fontSize=12).encode(text=alt.Text("Max loan:Q", format="$,.0f"), color=alt.value("#333"))
    c1.markdown(f"**Loan sizes to {usd(s['loan'])}: {s['binding']} is the binding constraint**")
    c1.altair_chart((chart + labels).properties(height=170), use_container_width=True)
    c1.caption(f"Limits: max LTV {L['max_ltv']:.0%} on value {usd(res['value'])} | min DSCR {L['min_dscr']:.2f}x "
               f"({L.get('dscr_basis', 'amortizing')}) | min debt yield {L['min_debt_yield']:.1%} | rate {res['perm']['rate_used']:.2%}, "
               f"{L['amort_years']}-yr amortization, constant {s['constant']:.2%}")
    met = pd.DataFrame([
        ["LTV", pct(s["ltv"], 1)], ["Debt yield", pct(s["debt_yield"])],
        ["Debt service, amortizing", usd(s["ds_amort"])], ["Debt service, interest-only", usd(s["ds_io"])],
        ["DSCR, amortizing", xx(s["dscr_amort"])], ["DSCR, interest-only", xx(s["dscr_io"])],
        ["Breakeven occupancy", pct(res["breakeven_occ"], 1)]], columns=["Metric", "Value"])
    c2.dataframe(met, hide_index=True, use_container_width=True)
    if "requested" in res:
        r = res["requested"]
        msg = (f"Requested {L['requested_ltv']:.0%} LTV = {usd(r['loan'])}: DSCR {r['dscr_amort']:.2f}x amortizing, "
               f"debt yield {r['debt_yield']:.2%}. ")
        (c2.error if r["fails"] else c2.success)(msg + (f"Fails {', '.join(r['fails'])}." if r["fails"] else "Passes all three tests."))
    c3, c4 = st.columns(2)
    inc = res["income"]
    c3.markdown("**NOI**")
    c3.dataframe(pd.DataFrame([["Gross potential rent", usd(inc["gpr"])], ["Vacancy & credit loss", usd(-inc["vacancy_loss"])],
                               ["Other income", usd(inc["other_income"])], ["Effective gross income", usd(inc["egi"])],
                               ["Operating expenses", usd(-inc["opex"])], ["NOI", usd(inc["noi"])]], columns=["Line", "Amount"]),
                 hide_index=True, use_container_width=True)
    su = res["sources_uses"]
    c4.markdown("**Sources & uses**")
    rows = [[k_, usd(v)] for k_, v in su["uses"].items()] + [["Total uses", usd(su["total_uses"])]] + \
           [[k_, usd(v)] for k_, v in su["sources"].items()]
    c4.dataframe(pd.DataFrame(rows, columns=["Item", "Amount"]), hide_index=True, use_container_width=True)

# ------------------------------------------------------------------ risk
with T["Risk"]:
    st.markdown("**Stress tests** (loan held at the sized amount; rate shock = refinance-rate risk)")
    sdf = pd.DataFrame(res["stress"])
    show = pd.DataFrame({"Scenario": sdf["name"], "NOI": sdf["noi"].map(usd), "Rate": sdf["rate"].map(pct),
                         "DSCR IO": sdf["dscr_io"].map(xx), "DSCR amortizing": sdf["dscr_amort"].map(xx),
                         "Debt yield": sdf["debt_yield"].map(pct), "Breakeven occ.": sdf["breakeven_occ"].map(lambda v: pct(v, 1)),
                         "Passes min DSCR": sdf["dscr_amort"].map(lambda v: "Yes" if v >= L["min_dscr"] else "NO")})
    if dev:
        show["Yield on cost"] = sdf["yoc"].map(pct)
        show["Takeout loan"] = sdf["takeout_loan"].map(usd)
        show["Takeout coverage"] = sdf["takeout_coverage"].map(xx)
    st.dataframe(show, hide_index=True, use_container_width=True)
    sc = res["scorecard"]
    st.markdown(f"**Weighted risk scorecard: {sc['score']:.2f} / 5, {sc['grade']}**")
    st.dataframe(pd.DataFrame([{"Factor": r_["factor"], "Value": round(r_["value"], 4), "Weight": f"{r_['weight']:.0%}",
                                "Score (1-5)": r_["score"], "Weighted": round(r_["weighted"], 2),
                                "Thresholds for 5/4/3/2": " / ".join(str(t) for t in r_["thresholds"])} for r_ in sc["rows"]]),
                 hide_index=True, use_container_width=True)
    st.caption("Weights and thresholds are my assumptions, editable on the Excel Stress Tests tab.")

# ------------------------------------------------------------------ pricing
with T["Pricing"]:
    st.markdown("**Benchmarks**")
    st.dataframe(pd.DataFrame([{"Benchmark": k_, "FRED series": b["fred_id"], "Rate": pct(b["value"]), "As of": b["as_of"],
                                "Source": b["source"]} for k_, b in bench.items()]), hide_index=True, use_container_width=True)
    lev = L.get("requested_ltv") or L["max_ltv"]
    st.markdown(f"**Indicative permanent pricing at {lev:.0%} LTV**")
    st.dataframe(pd.DataFrame([{"Lender type": r_["lender_type"], "Benchmark": r_["benchmark"], "Spread (bps)": r_["spread_bps"],
                                "Indicative rate": pct(r_["rate"]), "Eligible": "Yes" if r_["eligible"] else f"No ({r_['note']})",
                                "Term": r_["term"], "Amortization": r_["amort"], "Recourse": r_["recourse"]}
                               for r_ in res["pricing_perm"]]), hide_index=True, use_container_width=True)
    if dev:
        st.markdown(f"**Indicative construction pricing at {res['development']['ltc']:.0%} LTC**")
        st.dataframe(pd.DataFrame([{"Lender type": r_["lender_type"], "Spread (bps)": r_["spread_bps"], "Indicative rate": pct(r_["rate"]),
                                    "Eligible": "Yes" if r_["eligible"] else f"No ({r_['note']})", "Recourse": r_["recourse"]}
                                   for r_ in res["pricing_construction"]]), hide_index=True, use_container_width=True)
    st.info("Spreads are ASSUMPTIONS compiled from public market reports, not lender quotes:\n\n" +
            "\n".join(f"- **{n}**: {v['source']}" for n, v in {**pricing.PERM_LENDERS, **pricing.CONSTRUCTION_LENDERS}.items()))
    st.caption(f"Rate used for sizing: {res['perm']['rate_used']:.2%} "
               f"({'override' if L.get('rate') is not None else 'indicative rate for ' + L['lender_type']}).")

# ------------------------------------------------------------------ development
if dev:
    with T["Development"]:
        d = res["development"]
        dr = d["draws"]
        a, b, c, e = st.columns(4)
        a.metric("Total project cost", usd(d["total_cost"]))
        b.metric("Yield on cost (untrended)", pct(d["yoc_untrended"]), f"{d['dev_spread_bps']:.0f} bps over exit cap", delta_color="off")
        c.metric("Takeout coverage", xx(d["takeout_coverage"]), "PASS" if d["takeout_pass"] else "FAIL", delta_color="off")
        e.metric("Interest reserve shortfall", usd(dr["reserve_shortfall"]))
        st.markdown(f"**Takeout test:** perm loan {usd(d['takeout_loan'])} vs construction payoff {usd(d['payoff'])} "
                    f"= {d['takeout_coverage']:.2f}x coverage, excess {usd(d['gross_refi_excess'])}, "
                    f"net refinancing proceeds {usd(d['net_refi_proceeds'])}.")
        if d.get("trended_takeout"):
            t_ = d["trended_takeout"]
            st.caption(f"On trended NOI: perm loan {usd(t_['loan'])} ({t_['binding']} binds), coverage {t_['loan'] / d['payoff']:.2f}x.")
        st.markdown(f"**Interest reserve:** total interest {usd(dr['total_interest'])} at {dr['rate']:.2%}; lease-up NOI pays "
                    f"{usd(dr['interest_paid_by_noi'])}; reserve needed {usd(dr['required_reserve'])} vs {usd(dr['budgeted_reserve'])} budgeted.")
        dfm = pd.DataFrame(dr["rows"])
        long = dfm.melt(id_vars="month", value_vars=["end_balance"], var_name="Series", value_name="Balance")
        line = alt.Chart(long).mark_line(strokeWidth=2, color=ACCENT).encode(
            x=alt.X("month:Q", title="Month"), y=alt.Y("Balance:Q", title="Construction loan balance ($)", axis=alt.Axis(format="$,.0f")),
            tooltip=[alt.Tooltip("month:Q"), alt.Tooltip("Balance:Q", format="$,.0f")])
        commit = alt.Chart(pd.DataFrame({"y": [d["construction_loan"]]})).mark_rule(strokeDash=[4, 3], color="#555").encode(y="y:Q")
        st.altair_chart((line + commit).properties(height=260, title="Loan balance vs commitment (dashed)"), use_container_width=True)
        with st.expander("Monthly draw schedule"):
            st.dataframe(dfm.round(0), hide_index=True, use_container_width=True)

# ------------------------------------------------------------------ debt schedule
with T["Debt schedule"]:
    ds = res["debt_schedule"]
    a, b, c, e = st.columns(4)
    a.metric("Personal DTI (before loan)", pct(ds["dti_before"], 1))
    b.metric("Global DSCR", xx(ds["global_dscr"]))
    c.metric("Post-closing liquidity", usd(ds["post_closing_liquidity"]))
    e.metric("Contingent liability (guaranty)", usd(ds["contingent_liability"]))
    st.caption(f"Global DSCR = (personal income + property NOI) / (personal debt payments {usd(ds['personal_annual_payments'])} "
               f"+ property debt service {usd(ds['property_annual_ds'])}). Reserves cover "
               f"{ds['reserve_months']:.1f} months of property debt service.")
    proj = pd.DataFrame(ds["projection"]).set_index("year")
    proj["Total"] = proj.sum(axis=1)
    st.dataframe(proj.map(usd), use_container_width=True)

# ------------------------------------------------------------------ validation
with T["Validation"]:
    st.markdown("**Reconciliation to your source / test model.** The tool recomputes everything from the inputs and "
                "reports differences; it never adjusts its math to match.")
    v = res["validation"]
    if not v:
        st.caption("No source-model figures for this deal. Add them in the deal JSON under 'expected'.")
    else:
        vdf = pd.DataFrame([{"Item": x["label"], "Source figure": x["value"], "This model": x["calc"],
                             "Difference": x["diff"], "Status": x["status"], "Source": x.get("source", "")} for x in v])

        def color(val):
            return "background-color: #F8CBAD" if val == "MISMATCH" else ("background-color: #FFF2CC" if str(val).startswith("Minor") else "background-color: #C6EFCE")
        st.dataframe(vdf.style.map(color, subset=["Status"]), hide_index=True, use_container_width=True)

# ------------------------------------------------------------------ excel
with T["Excel model"]:
    st.markdown("Download the credit model. Every number is a live Excel formula built from the blue input cells, "
                "and each tab has a check row (OK / CHECK). The Credit Summary tab is the one-page view.")
    xbytes = excel.build(deal, res, bench)
    fname = deal["name"].split(":")[0].strip().replace(" ", "_") + "_credit_model.xlsx"
    st.download_button("Download Excel model", xbytes, file_name=fname,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
    st.caption("Tabs: Credit Summary, Inputs, NOI, Sources & Uses, Loan Sizing, Pricing, Amortization, "
               + ("Construction, " if dev else "") + "Stress Tests, Debt Schedule.")
