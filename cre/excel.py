"""
Builds the credit model workbook with LIVE formulas (openpyxl).

Blue = input, black = formula, green = link to another tab. Every tab ends with a
check row that says OK or CHECK. Inputs are defined names (e.g. =NOI/MinDSCR), so
the formulas read like the credit memo.
"""
from __future__ import annotations

import io

from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as COL
from openpyxl.workbook.defined_name import DefinedName

from . import pricing

F = "Arial"
BLUE, BLACK, GREEN, WHITE = "0000FF", "000000", "008000", "FFFFFF"
f_in = Font(name=F, size=10, color=BLUE)
f_calc = Font(name=F, size=10, color=BLACK)
f_link = Font(name=F, size=10, color=GREEN)
f_bold = Font(name=F, size=10, bold=True)
f_hdr = Font(name=F, size=10, bold=True, color=WHITE)
f_title = Font(name=F, size=14, bold=True, color="1F3864")
f_note = Font(name=F, size=9, italic=True, color="595959")
fill_hdr = PatternFill("solid", fgColor="1F3864")
fill_sec = PatternFill("solid", fgColor="D9E1F2")
fill_key = PatternFill("solid", fgColor="FFF2CC")
fill_bad = PatternFill("solid", fgColor="F8CBAD")
fill_ok = PatternFill("solid", fgColor="C6EFCE")
thin = Side(style="thin", color="BFBFBF")

USD = '$#,##0;($#,##0);"-"'
PCT = '0.00%;(0.00%);"-"'
PCT1 = '0.0%;(0.0%);"-"'
X = '0.00"x"'
BPS = '#,##0" bps"'
NUM = '#,##0;(#,##0);"-"'
DATE = "mm/dd/yyyy"


class Book:
    def __init__(self):
        self.wb = Workbook()
        self.checks = []          # (sheet, cell) of every check result

    def sheet(self, title, heading, sub):
        ws = self.wb.create_sheet(title)
        ws["A1"], ws["A2"] = heading, sub
        ws["A1"].font, ws["A2"].font = f_title, f_note
        ws.sheet_view.showGridLines = False
        return ws

    def name(self, nm, ws, ref):
        self.wb.defined_names[nm] = DefinedName(nm, attr_text=f"'{ws.title}'!{ref}")

    def check(self, ws, row, label, ok_formula):
        ws.cell(row=row, column=1, value=f"CHECK: {label}").font = f_bold
        c = ws.cell(row=row, column=2, value=f'=IF({ok_formula},"OK","CHECK")')
        c.font = f_bold
        ws.conditional_formatting.add(c.coordinate, FormulaRule(formula=[f'{c.coordinate}="CHECK"'], fill=fill_bad))
        ws.conditional_formatting.add(c.coordinate, FormulaRule(formula=[f'{c.coordinate}="OK"'], fill=fill_ok))
        self.checks.append((ws.title, c.coordinate))
        return c


def put(ws, ref, v, font=f_calc, fmt=None, fill=None):
    c = ws[ref]
    c.value, c.font = v, font
    if fmt:
        c.number_format = fmt
    if fill:
        c.fill = fill
    return c


def header(ws, row, labels, col=1):
    for i, lab in enumerate(labels):
        c = ws.cell(row=row, column=col + i, value=lab)
        c.font, c.fill = f_hdr, fill_hdr
        c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
    ws.row_dimensions[row].height = 30


def section(ws, row, text, ncols=4):
    for i in range(ncols):
        ws.cell(row=row, column=1 + i).fill = fill_sec
    ws.cell(row=row, column=1, value=text).font = f_bold


def widths(ws, w):
    for k, v in w.items():
        ws.column_dimensions[k].width = v


def build(deal, res, benchmarks=None) -> bytes:
    benchmarks = benchmarks or pricing.FALLBACK_BENCHMARKS
    dev = deal["path"] == "development"
    B = Book()
    wb = B.wb
    wb.remove(wb.active)
    prop, perm, bor = deal["property"], deal["perm"], deal["borrower"]
    inc = res["income"]

    # ================================================================== Inputs
    wi = B.sheet("Inputs", "Inputs", "Blue cells are inputs. Every other tab is formulas built from these cells.")
    header(wi, 4, ["Input", "Value", "Unit", "Note / source"])
    r = 5

    def inp(label, nm, val, fmt=None, unit="", note="", key=False):
        nonlocal r
        put(wi, f"A{r}", label)
        put(wi, f"B{r}", val, f_in, fmt, fill_key if key else None)
        put(wi, f"C{r}", unit)
        put(wi, f"D{r}", note, f_note)
        if nm:
            B.name(nm, wi, f"$B${r}")
        r += 1

    def sec(t):
        nonlocal r
        r += 1
        section(wi, r, t)
        r += 1

    sec("Deal")
    inp("Deal name", "DealName", deal["name"])
    inp("Loan path", "Path", "Development" if dev else "Permanent", note="Workbook layout is built for this path")
    inp("Property type", "PropType", prop["type"])
    inp("Market", "Market", prop.get("market", ""))
    inp("Units", "Units", prop.get("units") or 0, NUM)
    sec("Income (stabilized)")
    derived = prop.get("income_mode") == "noi"
    dnote = (f"Derived: NOI of ${prop['noi']:,.0f} entered directly; split with an assumed "
             f"{prop['assumed_opex_ratio']:.0%} opex ratio for stress tests") if derived else ""
    inp("Gross potential rent (annual)", "GPR", inc["gpr"], USD, "$", dnote)
    inp("Other income (annual)", "OtherInc", inc["other_income"], USD, "$")
    inp("Vacancy & credit loss", "VacPct", inc["vacancy_pct"], PCT1, "% of GPR", key=True)
    inp("Operating expenses (annual)", "Opex", inc["opex"], USD, "$", dnote)
    if dev:
        d, c = deal["development"], deal["construction"]
        sec("Development budget")
        inp("Land", "Land", d["land"], USD, "$")
        inp("Hard costs", "Hard", d["hard"], USD, "$")
        inp("Soft costs", "Soft", d["soft"], USD, "$")
        inp("Financing costs (origination + interest reserve)", "Financing", d["financing"], USD, "$")
        inp("Trended stabilized NOI", "NOITrended", d.get("noi_trended") or 0, USD, "$")
        inp("Exit / valuation cap rate", "ExitCap", d["exit_cap"], PCT, key=True)
        sec("Construction loan")
        inp("Max loan-to-cost", "LTCMax", c["ltc_max"], PCT1, key=True)
        inp("Spread over SOFR", "ConsSpread", c["spread_bps"], NUM, "bps")
        inp("SOFR (flat forward assumption)", "SOFRCons", res["development"]["sofr"], PCT, note="Latest SOFR at build time; see Pricing")
        inp("Origination fee", "OrigPct", c["origination_pct"], PCT, "% of loan")
        inp("Construction period", "ConsMonths", c["construction_months"], NUM, "months", "ASSUMPTION unless given")
        inp("Lease-up period", "LeaseupMonths", c["leaseup_months"], NUM, "months", "ASSUMPTION unless given")
        inp("Loan term", "ConsTerm", c["term_months"], NUM, "months")
        inp("Draw curve (straight / s-curve)", "DrawCurve", c.get("draw_curve", "straight"))
        inp("Lease-up NOI pays interest (1 = yes)", "NOIPaysInt", 1 if c.get("leaseup_noi_pays_interest", True) else 0, NUM)
    else:
        sec("Acquisition")
        inp("Purchase price", "PurchasePrice", prop["purchase_price"], USD, "$")
        inp("Closing costs", "ClosingPct", prop.get("closing_cost_pct", 0.0), PCT1, "% of price")
    sec("Permanent loan" + (" (takeout)" if dev else ""))
    inp("Rate override (blank = indicative rate from Pricing)", "RateOverride", perm.get("rate"), PCT, key=True)
    inp("Amortization", "AmortYrs", perm["amort_years"], NUM, "years")
    inp("Max LTV", "MaxLTV", perm["max_ltv"], PCT1, key=True)
    inp("Min DSCR", "MinDSCR", perm["min_dscr"], X, key=True)
    inp("Min debt yield", "MinDY", perm["min_debt_yield"], PCT1, key=True)
    inp("DSCR sizing basis (amortizing / io)", "DSCRBasis", perm.get("dscr_basis", "amortizing"))
    inp("Requested LTV (blank = none)", "ReqLTV", perm.get("requested_ltv"), PCT1)
    inp("Perm closing costs", "PermClosingPct", perm.get("closing_cost_pct", 0.0), PCT1, "% of loan")
    inp("Lender type", "LenderType", perm["lender_type"], note="Must match a row on the Pricing tab")
    sec("Borrower")
    inp("Credit score", "CreditScore", bor["credit_score"], NUM)
    inp("Liquidity", "Liquidity", bor["liquidity"], USD, "$")
    inp("Net worth", "NetWorth", bor["net_worth"], USD, "$")
    inp("Existing debt", "ExistingDebt", bor.get("existing_debt", 0), USD, "$")
    inp("Deals completed", "DealsDone", bor["deals_completed"], NUM)
    inp("Years of experience", "ExpYears", bor.get("experience_years", 0), NUM)
    inp("Annual personal income", "AnnualIncome", bor["annual_income"], USD, "$")
    inp("Recourse / guaranty share", "RecoursePct", bor.get("recourse_pct", 0.0), PCT1)
    inp("Sponsor's share of equity (rest from LP investors)", "SponsorEqShare", bor.get("sponsor_equity_share", 1.0), PCT1)
    r += 1
    put(wi, f"A{r}", "Assumptions: " + deal.get("assumptions_note", ""), f_note)
    r += 2
    B.check(wi, r, "required inputs are filled in",
            "AND(GPR>0,Opex>=0,MaxLTV>0,MinDSCR>0,MinDY>0,AmortYrs>0" + (",Land+Hard+Soft>0)" if dev else ",PurchasePrice>0)"))
    widths(wi, {"A": 46, "B": 22, "C": 12, "D": 80})
    wi.freeze_panes = "A5"

    # ================================================================== NOI
    wn = B.sheet("NOI", "Net Operating Income", "Stabilized operating statement, $ annual")
    header(wn, 4, ["Line", "Amount", "% of EGI"])
    rows = [("Gross potential rent", "=GPR"), ("Less: vacancy & credit loss", "=-GPR*VacPct"),
            ("Other income", "=OtherInc"), ("Effective gross income (EGI)", "=B5+B6+B7"),
            ("Less: operating expenses", "=-Opex"), ("Net operating income (NOI)", "=B8+B9")]
    for i, (lab, f) in enumerate(rows):
        rr = 5 + i
        put(wn, f"A{rr}", lab, f_bold if lab.startswith(("Effective", "Net")) else f_calc)
        put(wn, f"B{rr}", f, f_link if f.count("=") == 1 and f[1:].replace("-", "").isalpha() else f_calc, USD)
        put(wn, f"C{rr}", f"=IF($B$8=0,0,B{rr}/$B$8)", fmt=PCT1)
    B.name("EGI", wn, "$B$8")
    B.name("NOI", wn, "$B$10")
    put(wn, "A12", "NOI per unit"); put(wn, "B12", "=IF(Units=0,0,NOI/Units)", fmt=USD)
    put(wn, "A13", "Opex ratio (opex / EGI)"); put(wn, "B13", "=IF(EGI=0,0,Opex/EGI)", fmt=PCT1)
    if dev:
        put(wn, "A15", "Development returns", f_bold)
        put(wn, "A16", "Total project cost"); put(wn, "B16", "=TotalCost", f_link, USD)
        put(wn, "A17", "Yield on cost, untrended (NOI / total cost)"); put(wn, "B17", "=NOI/TotalCost", fmt=PCT)
        put(wn, "A18", "Yield on cost, trended"); put(wn, "B18", "=IF(NOITrended=0,0,NOITrended/TotalCost)", fmt=PCT)
        put(wn, "A19", "Exit cap rate"); put(wn, "B19", "=ExitCap", f_link, PCT)
        put(wn, "A20", "Development spread, untrended (bps)"); put(wn, "B20", "=(B17-B19)*10000", fmt=NUM)
        put(wn, "A21", "Stabilized value (NOI / exit cap)"); put(wn, "B21", "=NOI/ExitCap", fmt=USD)
        put(wn, "A22", "Trended value"); put(wn, "B22", "=IF(NOITrended=0,0,NOITrended/ExitCap)", fmt=USD)
        put(wn, "A23", "Construction loan debt yield (NOI / loan)"); put(wn, "B23", "=NOI/ConsLoan", fmt=PCT)
        B.name("YoC", wn, "$B$17"); B.name("YoCTrended", wn, "$B$18"); B.name("Value", wn, "$B$21")
        chk_row = 25
    else:
        put(wn, "A15", "Value (purchase price)"); put(wn, "B15", "=PurchasePrice", f_link, USD)
        put(wn, "A16", "Going-in cap rate (NOI / price)"); put(wn, "B16", "=NOI/PurchasePrice", fmt=PCT)
        B.name("Value", wn, "$B$15")
        chk_row = 18
    B.check(wn, chk_row, "NOI = GPR x (1 - vacancy) + other income - opex",
            "ABS(NOI-(GPR*(1-VacPct)+OtherInc-Opex))<1")
    widths(wn, {"A": 44, "B": 18, "C": 10})

    # ================================================================== Sources & Uses
    wsu = B.sheet("Sources & Uses", "Sources & Uses", "$")
    header(wsu, 4, ["Uses", "Amount", "% of total"])
    if dev:
        uses = [("Land", "=Land"), ("Hard costs", "=Hard"), ("Soft costs", "=Soft"),
                ("Financing (origination + interest reserve)", "=Financing")]
        src = [("Construction loan", "=ConsLoan"), ("Sponsor equity", "=B9-B13")]
    else:
        uses = [("Purchase price", "=PurchasePrice"), ("Closing costs", "=PurchasePrice*ClosingPct")]
        src = [("Senior loan (sized)", "=LoanSized"), ("Sponsor equity", None)]
    for i, (lab, f) in enumerate(uses):
        put(wsu, f"A{5+i}", lab); put(wsu, f"B{5+i}", f, f_link, USD)
    tu = 5 + len(uses)
    if not dev:
        tu = 9  # keep the same row for totals in both layouts
    put(wsu, f"A{tu}", "Total uses", f_bold); put(wsu, f"B{tu}", f"=SUM(B5:B{tu-1})", f_bold, USD)
    if dev:
        B.name("TotalCost", wsu, f"$B${tu}")
    header(wsu, 11, ["Sources", "Amount", "% of total"])
    put(wsu, "A12", src[0][0]); put(wsu, "B12", src[0][1], f_link, USD)
    put(wsu, "A13", src[1][0]); put(wsu, "B13", f"=B{tu}-B12", fmt=USD)
    put(wsu, "A14", "Total sources", f_bold); put(wsu, "B14", "=B12+B13", f_bold, USD)
    for rr in list(range(5, tu + 1)):
        put(wsu, f"C{rr}", f"=IF($B${tu}=0,0,B{rr}/$B${tu})", fmt=PCT1)
    for rr in (12, 13, 14):
        put(wsu, f"C{rr}", f"=IF($B$14=0,0,B{rr}/$B$14)", fmt=PCT1)
    B.name("EquityReq", wsu, "$B$13")
    if dev:
        put(wsu, "B12", "=LTCMax*B9", fmt=USD)  # loan computed here, named ConsLoan
        put(wsu, "B13", "=B9-B12", fmt=USD)
        B.name("ConsLoan", wsu, "$B$12")
    B.check(wsu, 16, "total sources = total uses", f"ABS(B14-B{tu})<1")
    widths(wsu, {"A": 44, "B": 18, "C": 10})

    # ================================================================== Pricing
    wp = B.sheet("Pricing", "Pricing", "Benchmarks from FRED + spread grid. SPREADS ARE ASSUMPTIONS from public market reports, not lender quotes.")
    header(wp, 4, ["Benchmark", "FRED series", "Rate", "As of", "Source"])
    for i, (k, b) in enumerate(benchmarks.items()):
        rr = 5 + i
        put(wp, f"A{rr}", k); put(wp, f"B{rr}", b["fred_id"])
        put(wp, f"C{rr}", b["value"], f_in, PCT); put(wp, f"D{rr}", b["as_of"], f_in); put(wp, f"E{rr}", b["source"], f_note)
        B.name(k, wp, f"$C${rr}")
    bench_rng = f"$A$5:$C${4+len(benchmarks)}"
    put(wp, "A11", "Leverage used to pick the band (requested LTV, else max LTV)")
    put(wp, "C11", '=IF(ISNUMBER(ReqLTV),IF(ReqLTV>0,ReqLTV,MaxLTV),MaxLTV)', fmt=PCT1)
    B.name("PricingLev", wp, "$C$11")
    put(wp, "A12", "Band upper bounds (LTV)")
    for j, bnd in enumerate(pricing.PERM_BANDS):
        put(wp, f"{COL(3+j)}12", bnd, f_in, PCT1)
    put(wp, "A13", "Band in use (0 = above all bands)")
    put(wp, "C13", "=IF(PricingLev>$F$12+1E-9,0,COUNTIF($C$12:$F$12,\"<\"&PricingLev-1E-9)+1)")
    B.name("Band", wp, "$C$13")
    header(wp, 15, ["Lender type", "Benchmark", "Spread band 1 (bps)", "Band 2", "Band 3", "Band 4", "Max LTV",
                    "Min DSCR", "Property types", "Benchmark rate", "Spread used", "Indicative rate", "Eligible?",
                    "Term", "Amortization", "Recourse", "Source (ASSUMPTION)"])
    p0 = 16
    for i, (nm, L) in enumerate(pricing.PERM_LENDERS.items()):
        rr = p0 + i
        put(wp, f"A{rr}", nm); put(wp, f"B{rr}", L["benchmark"], f_in)
        for j, sp in enumerate(L["spreads"]):
            put(wp, f"{COL(3+j)}{rr}", sp, f_in, NUM)
        put(wp, f"G{rr}", L["max_ltv"], f_in, PCT1); put(wp, f"H{rr}", L["min_dscr"], f_in, X)
        put(wp, f"I{rr}", ", ".join(L["property_types"]) if L["property_types"] else "All", f_in)
        put(wp, f"J{rr}", f"=VLOOKUP(B{rr},{bench_rng},3,FALSE)", fmt=PCT)
        put(wp, f"K{rr}", f'=IF(Band=0,"",IF(INDEX(C{rr}:F{rr},Band)="","",INDEX(C{rr}:F{rr},Band)))', fmt=NUM)
        put(wp, f"L{rr}", f'=IF(K{rr}="","n/a",J{rr}+K{rr}/10000)', fmt=PCT)
        put(wp, f"M{rr}", f'=IF(AND(K{rr}<>"",PricingLev<=G{rr}+1E-9,OR(I{rr}="All",ISNUMBER(SEARCH(PropType,I{rr})))),"Yes","No")')
        put(wp, f"N{rr}", L["term"]); put(wp, f"O{rr}", L["amort"]); put(wp, f"P{rr}", L["recourse"])
        put(wp, f"Q{rr}", L["source"], f_note)
    p1 = p0 + len(pricing.PERM_LENDERS) - 1
    rr = p1 + 2
    put(wp, f"A{rr}", "Selected lender type"); put(wp, f"C{rr}", "=LenderType", f_link)
    put(wp, f"A{rr+1}", "Indicative rate for selected lender")
    put(wp, f"C{rr+1}", f"=INDEX($L${p0}:$L${p1},MATCH(LenderType,$A${p0}:$A${p1},0))", fmt=PCT)
    put(wp, f"A{rr+2}", "RATE USED FOR SIZING (override if entered, else indicative)", f_bold)
    put(wp, f"C{rr+2}", f'=IF(ISNUMBER(RateOverride),RateOverride,C{rr+1})', f_bold, PCT, fill_key)
    B.name("RateUsed", wp, f"$C${rr+2}")
    nxt = rr + 4
    if dev:
        section(wp, nxt, "Construction lenders (bands are loan-to-cost)", 6)
        put(wp, f"A{nxt+1}", "LTC band upper bounds")
        for j, bnd in enumerate(pricing.CONSTRUCTION_BANDS):
            put(wp, f"{COL(3+j)}{nxt+1}", bnd, f_in, PCT1)
        put(wp, f"A{nxt+2}", "LTC band in use")
        bref = f"$C${nxt+1}:$F${nxt+1}"
        put(wp, f"C{nxt+2}", f'=IF(LTCMax>$F${nxt+1}+1E-9,0,COUNTIF({bref},"<"&LTCMax-1E-9)+1)')
        header(wp, nxt + 3, ["Lender type", "Benchmark", "Band 1", "Band 2", "Band 3", "Band 4", "Max LTC", "",
                             "", "Benchmark rate", "Spread used", "Indicative rate", "Eligible?", "", "", "Recourse", "Source (ASSUMPTION)"])
        for i, (nm, L) in enumerate(pricing.CONSTRUCTION_LENDERS.items()):
            q = nxt + 4 + i
            put(wp, f"A{q}", nm); put(wp, f"B{q}", L["benchmark"], f_in)
            for j, sp in enumerate(L["spreads"]):
                put(wp, f"{COL(3+j)}{q}", sp, f_in, NUM)
            put(wp, f"G{q}", L["max_ltc"], f_in, PCT1)
            put(wp, f"J{q}", f"=VLOOKUP(B{q},{bench_rng},3,FALSE)", fmt=PCT)
            band = f"$C${nxt+2}"
            put(wp, f"K{q}", f'=IF({band}=0,"",IF(INDEX(C{q}:F{q},{band})="","",INDEX(C{q}:F{q},{band})))', fmt=NUM)
            put(wp, f"L{q}", f'=IF(K{q}="","n/a",J{q}+K{q}/10000)', fmt=PCT)
            put(wp, f"M{q}", f'=IF(AND(K{q}<>"",LTCMax<=G{q}+1E-9),"Yes","No")')
            put(wp, f"P{q}", L["recourse"]); put(wp, f"Q{q}", L["source"], f_note)
        q = nxt + 4 + len(pricing.CONSTRUCTION_LENDERS) + 1
        put(wp, f"A{q}", "Construction rate used (SOFR input + deal spread)", f_bold)
        put(wp, f"C{q}", "=SOFRCons+ConsSpread/10000", f_bold, PCT, fill_key)
        B.name("ConsRate", wp, f"$C${q}")
        nxt = q + 2
    B.check(wp, nxt, "a sizing rate is available for the selected lender", "ISNUMBER(RateUsed)")
    widths(wp, {"A": 40, "B": 11, "C": 12, "D": 11, "E": 11, "F": 11, "G": 9, "H": 9, "I": 14, "J": 11,
                "K": 10, "L": 11, "M": 9, "N": 16, "O": 12, "P": 22, "Q": 70})

    # ================================================================== Loan Sizing
    wl = B.sheet("Loan Sizing", "Loan Sizing", "Loan = MIN(LTV, DSCR, debt yield). " +
                 ("Permanent takeout sized on stabilized NOI and value at the exit cap." if dev else ""))
    header(wl, 4, ["Item", "Value", "", "Note"])
    L_rows = [
        ("Value", "=Value", USD, f_link), ("NOI", "=NOI", USD, f_link), ("Rate used", "=RateUsed", PCT, f_link),
        ("Amortization (years)", "=AmortYrs", NUM, f_link),
        ("Loan constant (annual, amortizing)", "=PMT(RateUsed/12,AmortYrs*12,-1)*12", PCT, f_calc),
        ("Debt service factor used for DSCR sizing", '=IF(DSCRBasis="io",RateUsed,B9)', PCT, f_calc),
    ]
    for i, (lab, f, fm, ft) in enumerate(L_rows):
        put(wl, f"A{5+i}", lab); put(wl, f"B{5+i}", f, ft, fm)
    B.name("Constant", wl, "$B$9")
    header(wl, 12, ["Constraint", "Max loan", "Binding?", "Limit"])
    cons = [("LTV", "=MaxLTV*Value", "=MaxLTV", PCT1), ("DSCR", "=NOI/MinDSCR/B10", "=MinDSCR", X),
            ("Debt Yield", "=NOI/MinDY", "=MinDY", PCT1)]
    for i, (lab, f, lim, fm) in enumerate(cons):
        rr = 13 + i
        put(wl, f"A{rr}", lab); put(wl, f"B{rr}", f, fmt=USD)
        put(wl, f"C{rr}", f'=IF(B{rr}=$B$16,"BINDS","")', f_bold)
        put(wl, f"D{rr}", lim, f_link, fm)
    put(wl, "A16", "SIZED LOAN (minimum)", f_bold); put(wl, "B16", "=MIN(B13:B15)", f_bold, USD, fill_key)
    put(wl, "A17", "Binding constraint", f_bold); put(wl, "B17", "=INDEX(A13:A15,MATCH(B16,B13:B15,0))", f_bold)
    B.name("LoanSized", wl, "$B$16"); B.name("Binding", wl, "$B$17")
    header(wl, 19, ["Metric at sized loan", "Value", "", "Note"])
    met = [("LTV", "=LoanSized/Value", PCT1), ("Debt yield", "=NOI/LoanSized", PCT),
           ("Annual debt service, amortizing", "=LoanSized*Constant", USD),
           ("Annual debt service, interest-only", "=LoanSized*RateUsed", USD),
           ("DSCR, amortizing", "=NOI/B22", X), ("DSCR, interest-only", "=NOI/B23", X),
           ("IO minus amortizing DSCR", "=B25-B24", X),
           ("DSCR flag", '=IF(AND(B25>=MinDSCR,B24<MinDSCR),"Passes IO but FAILS amortizing",IF(B26>=0.15,"Material gap: IO vs amortizing","OK"))', None),
           ("Breakeven occupancy (opex + amortizing DS) / gross income", "=(Opex+B22)/(GPR+OtherInc)", PCT1)]
    for i, (lab, f, fm) in enumerate(met):
        put(wl, f"A{20+i}", lab); put(wl, f"B{20+i}", f, fmt=fm)
    for nm, row in [("LTV", 20), ("DebtYield", 21), ("DSAmort", 22), ("DSIO", 23), ("DSCRAmort", 24),
                    ("DSCRIO", 25), ("Breakeven", 28)]:
        B.name(nm, wl, f"$B${row}")
    put(wl, "D27", "Flags when the loan only works interest-only, or IO overstates coverage by 0.15x+", f_note)
    rr = 30
    section(wl, rr, "Requested loan test", 4)
    put(wl, f"A{rr+1}", "Requested loan (requested LTV x value)")
    put(wl, f"B{rr+1}", '=IF(ISNUMBER(ReqLTV),ReqLTV*Value,0)', fmt=USD)
    put(wl, f"A{rr+2}", "DSCR on sizing basis"); put(wl, f"B{rr+2}", f'=IF(B{rr+1}=0,0,NOI/(B{rr+1}*B10))', fmt=X)
    put(wl, f"A{rr+3}", "DSCR, amortizing"); put(wl, f"B{rr+3}", f'=IF(B{rr+1}=0,0,NOI/(B{rr+1}*Constant))', fmt=X)
    put(wl, f"A{rr+4}", "Debt yield"); put(wl, f"B{rr+4}", f'=IF(B{rr+1}=0,0,NOI/B{rr+1})', fmt=PCT)
    put(wl, f"A{rr+5}", "Result", f_bold)
    put(wl, f"B{rr+5}", f'=IF(B{rr+1}=0,"No request entered",IF(AND(ReqLTV<=MaxLTV+1E-9,B{rr+2}>=MinDSCR-1E-9,B{rr+4}>=MinDY-1E-9),"PASSES","FAILS"))', f_bold)
    put(wl, f"A{rr+6}", "If it fails: loan sizes down to")
    put(wl, f"B{rr+6}", f'=IF(B{rr+5}="FAILS",LoanSized,0)', fmt=USD)
    put(wl, f"C{rr+6}", f'=IF(B{rr+5}="FAILS",Binding&" binds","")')
    wl.conditional_formatting.add(f"B{rr+5}", FormulaRule(formula=[f'B{rr+5}="FAILS"'], fill=fill_bad))
    wl.conditional_formatting.add(f"B{rr+5}", FormulaRule(formula=[f'B{rr+5}="PASSES"'], fill=fill_ok))
    B.name("ReqLoan", wl, f"$B${rr+1}"); B.name("ReqDSCRAmort", wl, f"$B${rr+3}"); B.name("ReqDY", wl, f"$B${rr+4}")
    rr += 8
    if dev:
        section(wl, rr, "Construction loan and takeout test", 4)
        tk = [("Total project cost", "=TotalCost", USD), ("Construction loan (max LTC x cost)", "=ConsLoan", USD),
              ("Actual LTC", "=ConsLoan/TotalCost", PCT1), ("Origination fee", "=OrigPct*ConsLoan", USD),
              ("Interest reserve budgeted (financing less origination)", f"=Financing-B{rr+4}", USD),
              ("Interest reserve required (Construction tab)", "=ReqReserve", USD),
              ("Interest reserve shortfall", f"=MAX(0,B{rr+6}-B{rr+5})", USD),
              ("Construction loan payoff at takeout", "=MIN(ConsEndBal,ConsLoan)", USD),
              ("Permanent takeout loan (sized above)", "=LoanSized", USD),
              ("Takeout coverage (perm / payoff)", f"=B{rr+9}/B{rr+8}", X),
              ("Excess proceeds (perm - payoff)", f"=B{rr+9}-B{rr+8}", USD),
              ("Less: perm closing costs", "=-PermClosingPct*LoanSized", USD),
              ("Net refinancing proceeds", f"=B{rr+11}+B{rr+12}", USD),
              ("Takeout test", f'=IF(B{rr+9}>=B{rr+8},"PASS: perm covers payoff","FAIL: perm short by "&TEXT(B{rr+8}-B{rr+9},"$#,##0"))', None),
              ("Trended: value at exit cap", "=IF(NOITrended=0,0,NOITrended/ExitCap)", USD),
              ("Trended: perm loan (min of 3 constraints)",
               f"=IF(NOITrended=0,0,MIN(MaxLTV*B{rr+15},NOITrended/MinDSCR/B10,NOITrended/MinDY))", USD),
              ("Trended: takeout coverage", f"=IF(B{rr+16}=0,0,B{rr+16}/B{rr+8})", X)]
        for i, (lab, f, fm) in enumerate(tk):
            put(wl, f"A{rr+1+i}", lab); put(wl, f"B{rr+1+i}", f, fmt=fm)
        B.name("LTCActual", wl, f"$B${rr+3}"); B.name("ReserveShort", wl, f"$B${rr+7}")
        B.name("Payoff", wl, f"$B${rr+8}"); B.name("TakeoutCov", wl, f"$B${rr+10}")
        B.name("NetRefi", wl, f"$B${rr+13}")
        rr += len(tk) + 2
    B.check(wl, rr, "sized loan does not exceed any constraint", "AND(B16<=B13+0.5,B16<=B14+0.5,B16<=B15+0.5)")
    B.check(wl, rr + 1, "constant ties to Amortization tab payment", "ABS(Constant-AmortPay*12/LoanSized)<0.000001")
    widths(wl, {"A": 52, "B": 22, "C": 10, "D": 60})

    # ================================================================== Amortization
    wa = B.sheet("Amortization", "Amortization schedule: permanent loan", "Monthly, fully amortizing at the sizing rate")
    put(wa, "A3", "Monthly payment"); put(wa, "B3", "=PMT(RateUsed/12,AmortYrs*12,-LoanSized)", f_bold, USD)
    B.name("AmortPay", wa, "$B$3")
    header(wa, 5, ["Month", "Year", "Beginning balance", "Payment", "Interest", "Principal", "Ending balance"])
    n = int(perm["amort_years"] * 12)
    for m in range(1, n + 1):
        rr = 5 + m
        put(wa, f"A{rr}", m); put(wa, f"B{rr}", f"=ROUNDUP(A{rr}/12,0)")
        put(wa, f"C{rr}", "=LoanSized" if m == 1 else f"=G{rr-1}", fmt=USD)
        put(wa, f"D{rr}", "=AmortPay", fmt=USD)
        put(wa, f"E{rr}", f"=C{rr}*RateUsed/12", fmt=USD)
        put(wa, f"F{rr}", f"=D{rr}-E{rr}", fmt=USD)
        put(wa, f"G{rr}", f"=C{rr}-F{rr}", fmt=USD)
    last = 5 + n
    header(wa, 5, ["Year", "Payments", "Interest", "Principal", "Ending balance"], col=9)
    for y in range(1, int(perm["amort_years"]) + 1):
        rr = 5 + y
        put(wa, f"I{rr}", y)
        for col, src in [("J", "D"), ("K", "E"), ("L", "F")]:
            put(wa, f"{col}{rr}", f"=SUMIFS({src}$6:{src}${last},$B$6:$B${last},I{rr})", fmt=USD)
        put(wa, f"M{rr}", f"=INDEX($G$6:$G${last},MATCH(I{rr}*12,$A$6:$A${last},0))", fmt=USD)
    put(wa, "D3", "CHECK: principal repaid = loan, ending balance = 0", f_bold)
    c = put(wa, "G3", f'=IF(AND(ABS(SUM(F6:F{last})-LoanSized)<1,ABS(G{last})<1),"OK","CHECK")', f_bold)
    wa.conditional_formatting.add("G3", FormulaRule(formula=['G3="CHECK"'], fill=fill_bad))
    wa.conditional_formatting.add("G3", FormulaRule(formula=['G3="OK"'], fill=fill_ok))
    B.checks.append(("Amortization", "G3"))
    wa.freeze_panes = "A6"
    widths(wa, {"A": 8, "B": 6, "C": 16, "D": 13, "E": 13, "F": 13, "G": 16, "H": 3, "I": 6, "J": 14, "K": 14, "L": 14, "M": 16})

    # ================================================================== Construction (development only)
    if dev:
        wc = B.sheet("Construction", "Construction draws & interest reserve",
                     "Equity funds first, then the loan. Interest accrues on the loan balance; lease-up NOI pays interest first, the reserve covers the rest.")
        put(wc, "A3", "Rows were generated for the construction + lease-up months at build time. Regenerate from the app if timing changes.", f_note)
        header(wc, 5, ["Month", "Phase", "Cost weight", "Cost draw", "Equity left (beg)", "Equity draw", "Loan cost draw",
                       "Loan balance (beg)", "Interest", "Lease-up NOI", "NOI applied to interest", "Reserve draw", "Loan balance (end)"])
        C = deal["construction"]["construction_months"]
        T = C + deal["construction"]["leaseup_months"]
        for m in range(T + 1):
            rr = 6 + m
            put(wc, f"A{rr}", m)
            put(wc, f"B{rr}", f'=IF(A{rr}=0,"Close",IF(A{rr}<=ConsMonths,"Construction","Lease-up"))')
            put(wc, f"C{rr}", f'=IF(OR(A{rr}=0,A{rr}>ConsMonths),0,IF(DrawCurve="s-curve",(1-COS(2*PI()*(A{rr}-0.5)/ConsMonths))/ConsMonths,1/ConsMonths))', fmt="0.0000")
            put(wc, f"D{rr}", f"=IF(A{rr}=0,Land+OrigPct*ConsLoan,(Hard+Soft)*C{rr})", fmt=USD)
            put(wc, f"E{rr}", "=TotalCost-ConsLoan" if m == 0 else f"=E{rr-1}-F{rr-1}", fmt=USD)
            put(wc, f"F{rr}", f"=MIN(D{rr},E{rr})", fmt=USD)
            put(wc, f"G{rr}", f"=D{rr}-F{rr}", fmt=USD)
            put(wc, f"H{rr}", 0 if m == 0 else f"=M{rr-1}", fmt=USD)
            put(wc, f"I{rr}", f"=IF(A{rr}=0,0,H{rr}*ConsRate/12)", fmt=USD)
            put(wc, f"J{rr}", f"=IF(AND(A{rr}>ConsMonths,NOIPaysInt=1),NOI/12*(A{rr}-ConsMonths)/LeaseupMonths,0)", fmt=USD)
            put(wc, f"K{rr}", f"=MIN(I{rr},J{rr})", fmt=USD)
            put(wc, f"L{rr}", f"=I{rr}-K{rr}", fmt=USD)
            put(wc, f"M{rr}", f"=H{rr}+G{rr}+L{rr}", fmt=USD)
        last_c = 6 + T
        tr = last_c + 2
        put(wc, f"A{tr}", "Totals", f_bold)
        for col in "DFGIKL":
            put(wc, f"{col}{tr}", f"=SUM({col}6:{col}{last_c})", f_bold, USD)
        put(wc, f"M{tr}", f"=M{last_c}", f_bold, USD)
        B.name("ReqReserve", wc, f"$L${tr}"); B.name("ConsEndBal", wc, f"$M${tr}")
        put(wc, f"A{tr+1}", "Required interest reserve = total reserve draws (col L). Ending balance = construction payoff before cap at commitment.", f_note)
        B.check(wc, tr + 3, "cost draws = land + hard + soft + origination",
                f"ABS(D{tr}-(Land+Hard+Soft+OrigPct*ConsLoan))<1")
        B.check(wc, tr + 4, "equity drawn = sponsor equity", f"ABS(F{tr}-(TotalCost-ConsLoan))<1")
        # Credit finding, not a model-integrity check: shown here and on the Credit Summary
        put(wc, f"A{tr+6}", "CREDIT FLAG: loan balance vs commitment", f_bold)
        fl = put(wc, f"B{tr+6}", f'=IF(M{tr}<=ConsLoan+1,"Within commitment","Exceeds commitment by "&TEXT(M{tr}-ConsLoan,"$#,##0")&": interest reserve short, sponsor must fund")', f_bold)
        wc.conditional_formatting.add(fl.coordinate, FormulaRule(formula=[f'LEFT({fl.coordinate},7)="Exceeds"'], fill=fill_bad))
        wc.freeze_panes = "B6"
        widths(wc, {"A": 7, "B": 13, "C": 10, **{k: 15 for k in "DEFGHIJKLM"}})

    # ================================================================== Stress Tests
    wt = B.sheet("Stress Tests", "Stress tests & risk scorecard",
                 "Shocks are inputs (blue). Loan held at the sized amount; rate shock = refinance-rate risk." +
                 (" Takeout loan re-sized under each shock." if dev else ""))
    cols = ["Scenario", "Rate shock", "Vacancy shock (pts)", "Rent shock", "Cost shock", "Cap shock",
            "NOI", "Rate", "DSCR, IO", "DSCR, amortizing", "Debt yield", "Breakeven occupancy", "Passes min DSCR?"]
    if dev:
        cols += ["Total cost", "Yield on cost", "Equity required", "Value", "Takeout loan", "Takeout coverage"]
    header(wt, 4, cols)
    sc = res["stress"]
    for i, s in enumerate(sc):
        rr = 5 + i
        put(wt, f"A{rr}", s["name"])
        for col, k, fm in [("B", "d_rate", PCT), ("C", "d_vac", PCT1), ("D", "d_rent", PCT1), ("E", "d_cost", PCT1), ("F", "d_cap", PCT)]:
            put(wt, f"{col}{rr}", s[k], f_in, fm)
        put(wt, f"G{rr}", f"=GPR*(1+D{rr})*(1-(VacPct+C{rr}))+OtherInc-Opex", fmt=USD)
        put(wt, f"H{rr}", f"=RateUsed+B{rr}", fmt=PCT)
        put(wt, f"I{rr}", f"=G{rr}/(LoanSized*H{rr})", fmt=X)
        put(wt, f"J{rr}", f"=G{rr}/(LoanSized*PMT(H{rr}/12,AmortYrs*12,-1)*12)", fmt=X)
        put(wt, f"K{rr}", f"=G{rr}/LoanSized", fmt=PCT)
        put(wt, f"L{rr}", f"=(Opex+LoanSized*PMT(H{rr}/12,AmortYrs*12,-1)*12)/(GPR*(1+D{rr})+OtherInc)", fmt=PCT1)
        put(wt, f"M{rr}", f'=IF(J{rr}>=MinDSCR,"Yes","NO")')
        if dev:
            put(wt, f"N{rr}", f"=Land+(Hard+Soft)*(1+E{rr})+Financing", fmt=USD)
            put(wt, f"O{rr}", f"=G{rr}/N{rr}", fmt=PCT)
            put(wt, f"P{rr}", f"=N{rr}-ConsLoan", fmt=USD)
            put(wt, f"Q{rr}", f"=G{rr}/(ExitCap+F{rr})", fmt=USD)
            put(wt, f"R{rr}", f'=MIN(MaxLTV*Q{rr},G{rr}/MinDSCR/IF(DSCRBasis="io",H{rr},PMT(H{rr}/12,AmortYrs*12,-1)*12),G{rr}/MinDY)', fmt=USD)
            put(wt, f"S{rr}", f"=R{rr}/Payoff", fmt=X)
    s_last = 4 + len(sc)
    wt.conditional_formatting.add(f"M5:M{s_last}", FormulaRule(formula=['M5="NO"'], fill=fill_bad))
    put(wt, f"A{s_last+1}", "Worst case", f_bold)
    put(wt, f"J{s_last+1}", f"=MIN(J5:J{s_last})", f_bold, X)
    put(wt, f"L{s_last+1}", f"=MAX(L5:L{s_last})", f_bold, PCT1)
    B.name("WorstDSCR", wt, f"$J${s_last+1}")
    if dev:
        put(wt, f"S{s_last+1}", f"=MIN(S5:S{s_last})", f_bold, X)
        B.name("WorstCov", wt, f"$S${s_last+1}")
    # scorecard
    k0 = s_last + 4
    section(wt, k0 - 1, "Weighted risk scorecard (1 = high risk, 5 = low risk). Thresholds and weights are inputs.", 11)
    header(wt, k0, ["Factor", "Value", "Weight", "Better if", "Score 5 at", "Score 4 at", "Score 3 at", "Score 2 at", "Score", "Weighted"])
    from .engine import SCORECARD, GRADES
    vals = {"DSCR (amortizing)": ("=DSCRAmort", X), "LTV / LTC": ("=LTCActual" if dev else "=LTV", PCT1),
            "Debt yield": ("=DebtYield", PCT), "Breakeven occupancy": ("=Breakeven", PCT1),
            "Worst stress DSCR": ("=WorstDSCR", X), "Takeout coverage (worst stress)": ("=WorstCov" if dev else 0, X),
            "Credit score": ("=CreditScore", NUM),
            "Liquidity / loan": (f"=Liquidity/{'ConsLoan' if dev else 'LoanSized'}", PCT1),
            "Net worth / loan": (f"=NetWorth/{'ConsLoan' if dev else 'LoanSized'}", X),
            "Experience (deals completed)": ("=DealsDone", NUM)}
    for i, (nm, wp_, wd_, direction, t) in enumerate(SCORECARD):
        rr = k0 + 1 + i
        f, fm = vals[nm]
        put(wt, f"A{rr}", nm); put(wt, f"B{rr}", f, f_link, fm)
        put(wt, f"C{rr}", wd_ if dev else wp_, f_in, PCT1); put(wt, f"D{rr}", direction, f_in)
        for j, thr in enumerate(t):
            put(wt, f"{COL(5+j)}{rr}", thr, f_in, fm)
        cmp = ">=" if direction == "higher" else "<="
        put(wt, f"I{rr}", f'=IF(D{rr}="higher",IF(B{rr}>=E{rr},5,IF(B{rr}>=F{rr},4,IF(B{rr}>=G{rr},3,IF(B{rr}>=H{rr},2,1)))),'
                          f'IF(B{rr}<=E{rr},5,IF(B{rr}<=F{rr},4,IF(B{rr}<=G{rr},3,IF(B{rr}<=H{rr},2,1)))))')
        put(wt, f"J{rr}", f"=C{rr}*I{rr}", fmt="0.00")
    k1 = k0 + len(SCORECARD)
    put(wt, f"A{k1+1}", "RISK SCORE (weighted, 1-5)", f_bold)
    put(wt, f"J{k1+1}", f"=SUM(J{k0+1}:J{k1})", f_bold, "0.00", fill_key)
    put(wt, f"A{k1+2}", "Grade", f_bold)
    g = f"J{k1+1}"
    put(wt, f"J{k1+2}", "=" + "".join(f'IF({g}>={cut},"{lab}",' for cut, lab in GRADES[:-1]) + f'"{GRADES[-1][1]}"' + ")" * (len(GRADES) - 1), f_bold)
    B.name("RiskScore", wt, f"$J${k1+1}"); B.name("RiskGrade", wt, f"$J${k1+2}")
    B.check(wt, k1 + 4, "scorecard weights sum to 100%", f"ABS(SUM(C{k0+1}:C{k1})-1)<0.0001")
    B.check(wt, k1 + 5, "base-case DSCR ties to Loan Sizing", "ABS(J5-DSCRAmort)<0.000001")
    widths(wt, {"A": 40, **{COL(i): 13 for i in range(2, 20)}})
    wt.freeze_panes = "B5"

    # ================================================================== Debt Schedule
    wd = B.sheet("Debt Schedule", "Personal debt schedule with the new loan",
                 "Sample debts are placeholders: replace them with yours. Balances projected with FV at each debt's rate and payment.")
    header(wd, 4, ["Debt", "Balance", "Rate", "Monthly payment", "Annual payment", "Note"])
    debts = deal.get("personal_debts", [])
    for i, dbt in enumerate(debts):
        rr = 5 + i
        put(wd, f"A{rr}", dbt["name"], f_in); put(wd, f"B{rr}", dbt["balance"], f_in, USD)
        put(wd, f"C{rr}", dbt["rate"], f_in, PCT); put(wd, f"D{rr}", dbt["monthly_payment"], f_in, USD)
        put(wd, f"E{rr}", f"=D{rr}*12", fmt=USD)
    pr = 5 + len(debts)
    put(wd, f"A{pr}", "New property loan (permanent)"); put(wd, f"B{pr}", "=LoanSized", f_link, USD)
    put(wd, f"C{pr}", "=RateUsed", f_link, PCT); put(wd, f"D{pr}", "=AmortPay", f_link, USD)
    put(wd, f"E{pr}", f"=D{pr}*12", fmt=USD); put(wd, f"F{pr}", "Serviced by property NOI", f_note)
    tr = pr + 1
    put(wd, f"A{tr}", "Total", f_bold)
    for col in "BDE":
        put(wd, f"{col}{tr}", f"=SUM({col}5:{col}{pr})", f_bold, USD)
    m0 = tr + 2
    mets = [("Annual personal income", "=AnnualIncome", USD),
            ("Personal debt payments (annual, excl. property loan)", f"=SUM(E5:E{pr-1})", USD),
            ("Personal DTI (payments / income)", f"=IF(B{m0}=0,0,B{m0+1}/B{m0})", PCT1),
            ("Property debt service (annual)", f"=E{pr}", USD),
            ("Global DSCR ((income + NOI) / (personal + property debt service))", f"=(B{m0}+NOI)/(B{m0+1}+B{m0+3})", X),
            ("Sponsor equity required at closing (equity x sponsor share)", "=EquityReq*SponsorEqShare", USD),
            ("Post-closing liquidity", f"=Liquidity-B{m0+5}", USD),
            ("Months of property debt service covered by liquidity", f"=IF(AmortPay=0,0,B{m0+6}/AmortPay)", "0.0"),
            ("Contingent liability (guaranty share x " + ("construction" if dev else "property") + " loan)",
             f"=RecoursePct*{'ConsLoan' if dev else 'LoanSized'}", USD)]
    for i, (lab, f, fm) in enumerate(mets):
        put(wd, f"A{m0+i}", lab); put(wd, f"B{m0+i}", f, f_link if f.startswith("=Annual") or f.startswith("=Equity") else f_calc, fm)
    B.name("GlobalDSCR", wd, f"$B${m0+4}"); B.name("PostLiq", wd, f"$B${m0+6}")
    j0 = m0 + len(mets) + 2
    header(wd, j0, ["Year"] + [d_["name"] for d_ in debts] + ["New property loan", "Total"])
    for y in range(0, 11):
        rr = j0 + 1 + y
        put(wd, f"A{rr}", y)
        for i in range(len(debts) + 1):
            src = 5 + i
            put(wd, f"{COL(2+i)}{rr}", f"=MAX(0,FV($C${src}/12,12*$A{rr},$D${src},-$B${src}))", fmt=USD)
        put(wd, f"{COL(3+len(debts))}{rr}", f"=SUM(B{rr}:{COL(2+len(debts))}{rr})", f_bold, USD)
    B.check(wd, j0 + 13, "year-0 projection = current balances", f"ABS({COL(3+len(debts))}{j0+1}-B{tr})<1")
    widths(wd, {"A": 58, "B": 16, "C": 16, "D": 16, "E": 16, "F": 22, "G": 16})

    # ================================================================== Credit Summary
    ws = B.sheet("Credit Summary", "Credit Summary", "One-page view for the credit officer. All values link to the model.")
    wb.move_sheet("Credit Summary", offset=-(len(wb.sheetnames) - 1))
    put(ws, "A3", "=DealName", f_link)
    put(ws, "A4", '=Path&" | "&PropType&" | "&Market&" | "&TEXT(Units,"#,##0")&" units"', f_link)
    lines = [("LOAN", None, None)]
    if dev:
        lines += [("Construction loan", "=ConsLoan", USD), ("Loan-to-cost", "=LTCActual", PCT1),
                  ("Construction rate (SOFR + spread)", "=ConsRate", PCT), ("Total project cost", "=TotalCost", USD),
                  ("Yield on cost (untrended / trended)", '=TEXT(YoC,"0.00%")&" / "&TEXT(YoCTrended,"0.00%")', None),
                  ("Interest reserve shortfall", "=ReserveShort", USD),
                  ("Permanent takeout loan", "=LoanSized", USD), ("Takeout coverage (perm / payoff)", "=TakeoutCov", X),
                  ("Net refinancing proceeds", "=NetRefi", USD)]
    else:
        lines += [("Sized loan", "=LoanSized", USD), ("Requested loan", "=ReqLoan", USD)]
    lines += [("Binding constraint", "=Binding", None), ("Rate used / lender type", '=TEXT(RateUsed,"0.00%")&" / "&LenderType', None),
              ("LTV (perm)", "=LTV", PCT1), ("Debt yield", "=DebtYield", PCT),
              ("DSCR, amortizing / interest-only", '=TEXT(DSCRAmort,"0.00x")&" / "&TEXT(DSCRIO,"0.00x")', None),
              ("DSCR flag", "='Loan Sizing'!B27", None), ("Breakeven occupancy", "=Breakeven", PCT1),
              ("RISK", None, None), ("Risk score / grade", '=TEXT(RiskScore,"0.00")&" / "&RiskGrade', None),
              ("Worst stress DSCR (amortizing)", "=WorstDSCR", X)]
    if dev:
        lines += [("Worst stress takeout coverage", "=WorstCov", X)]
    lines += [("BORROWER", None, None), ("Credit score", "=CreditScore", NUM), ("Liquidity / post-closing", '=TEXT(Liquidity,"$#,##0")&" / "&TEXT(PostLiq,"$#,##0")', None),
              ("Net worth", "=NetWorth", USD), ("Global DSCR", "=GlobalDSCR", X)]
    rr = 6
    for lab, f, fm in lines:
        if f is None:
            section(ws, rr, lab, 4)
        else:
            put(ws, f"A{rr}", lab); put(ws, f"B{rr}", f, f_link, fm)
            ws[f"B{rr}"].alignment = Alignment(horizontal="right")
        rr += 1
    # model integrity
    rr += 1
    section(ws, rr, "MODEL CHECKS", 4)
    chk_formula = "+".join(f"('{s}'!{c}=\"CHECK\")" for s, c in B.checks)
    put(ws, f"A{rr+1}", "Failed checks across all tabs"); put(ws, f"B{rr+1}", f"={chk_formula}", f_bold, NUM)
    put(ws, f"A{rr+2}", "Model status")
    c = put(ws, f"B{rr+2}", f'=IF(B{rr+1}=0,"OK: all tabs tie","CHECK: see tabs")', f_bold)
    ws.conditional_formatting.add(c.coordinate, FormulaRule(formula=[f'B{rr+1}>0'], fill=fill_bad))
    ws.conditional_formatting.add(c.coordinate, FormulaRule(formula=[f'B{rr+1}=0'], fill=fill_ok))
    # validation vs source model
    rr += 4
    exp = deal.get("expected", [])
    if exp:
        section(ws, rr, "RECONCILIATION TO SOURCE / TEST MODEL (model recomputes; differences are flagged, never forced)", 6)
        header(ws, rr + 1, ["Item", "Source figure", "This model", "Difference", "Status", "Source"])
        refs = {"total_cost": "TotalCost", "construction_loan": "ConsLoan", "equity": "EquityReq", "ltc": "LTCActual",
                "yoc_untrended": "YoC", "yoc_trended": "YoCTrended", "loan": "LoanSized", "binding": "Binding",
                "constant": "Constant", "dscr_io": "DSCRIO", "dscr_amort": "DSCRAmort", "net_refi_proceeds": "NetRefi",
                "requested_loan": "ReqLoan", "requested_dscr_amort": "ReqDSCRAmort", "requested_debt_yield": "ReqDY",
                "ltv": "LTV", "debt_yield": "DebtYield", "takeout_coverage": "TakeoutCov"}
        for i, e in enumerate(exp):
            q = rr + 2 + i
            is_txt = isinstance(e["value"], str)
            fm = USD if (not is_txt and abs(e["value"]) >= 1000) else (None if is_txt else ("0.0000" if abs(e["value"]) < 1 else "0.00"))
            put(ws, f"A{q}", e["label"]); put(ws, f"B{q}", e["value"], f_in, fm)
            put(ws, f"C{q}", f"={refs[e['key']]}", f_link, fm)
            if is_txt:
                put(ws, f"E{q}", f'=IF(C{q}=B{q},"Ties","MISMATCH")', f_bold)
            else:
                put(ws, f"D{q}", f"=C{q}-B{q}", fmt=fm)
                tol = e.get("tol", abs(e["value"]) * 0.001)
                put(ws, f"E{q}", f'=IF(ABS(D{q})<={tol},"Ties",IF(ABS(D{q})<=ABS(B{q})*0.01,"Minor mismatch (<1%)","MISMATCH"))', f_bold)
            put(ws, f"F{q}", e.get("source", ""), f_note)
            ws.conditional_formatting.add(f"E{q}", FormulaRule(formula=[f'E{q}="MISMATCH"'], fill=fill_bad))
            ws.conditional_formatting.add(f"E{q}", FormulaRule(formula=[f'E{q}="Ties"'], fill=fill_ok))
    widths(ws, {"A": 46, "B": 36, "C": 18, "D": 16, "E": 22, "F": 22})
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    for name in ["Loan Sizing", "Pricing", "Stress Tests", "Construction", "Debt Schedule", "Amortization"]:
        if name in wb.sheetnames:   # print wide tabs landscape, one page wide
            w_ = wb[name]
            w_.page_setup.orientation = "landscape"
            w_.page_setup.fitToWidth, w_.page_setup.fitToHeight = 1, 0
            w_.sheet_properties.pageSetUpPr.fitToPage = True

    order = ["Credit Summary", "Inputs", "NOI", "Sources & Uses", "Loan Sizing", "Pricing", "Amortization"] + \
            (["Construction"] if dev else []) + ["Stress Tests", "Debt Schedule"]
    wb._sheets = [wb[s] for s in order]
    # openpyxl writes formulas without cached results, so tell Excel to calculate everything on open
    wb.calculation.fullCalcOnLoad = True
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
