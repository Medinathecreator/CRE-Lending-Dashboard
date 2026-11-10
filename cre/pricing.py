"""
Benchmarks (FRED) and the lender spread grid.

Every spread below is an ASSUMPTION taken from public market commentary, not a
lender quote. Sources are listed next to each lender type and shown in the app
and the Excel Pricing tab. Update them when you have fresher data.
"""
from __future__ import annotations

import csv
import io
import os
from datetime import date

# Fallback benchmark levels used when FRED can't be reached (e.g. offline).
# Treasury par yields as of Nov 6, 2026 (sofrrate.com); SOFR as of Nov 9, 2026
# (FRED, via multifamily-usa.com/rates).
FALLBACK_BENCHMARKS = {
    "SOFR": {"value": 0.0388, "as_of": "2026-11-09", "fred_id": "SOFR", "source": "fallback"},
    "UST2": {"value": 0.0451, "as_of": "2026-11-06", "fred_id": "DGS2", "source": "fallback"},
    "UST5": {"value": 0.0464, "as_of": "2026-11-06", "fred_id": "DGS5", "source": "fallback"},
    "UST10": {"value": 0.0486, "as_of": "2026-11-06", "fred_id": "DGS10", "source": "fallback"},
}

# Leverage bands: upper bound of LTV (perm) or LTC (construction) for each column of spreads
PERM_BANDS = [0.55, 0.65, 0.75, 0.80]
CONSTRUCTION_BANDS = [0.60, 0.65, 0.75, 0.85]

# Spreads in bps per band. None = lender doesn't lend at that leverage.
PERM_LENDERS = {
    "Agency (Fannie/Freddie)": {
        "benchmark": "UST10", "spreads": [105, 125, 145, 165], "max_ltv": 0.80, "min_dscr": 1.25,
        "property_types": ["Multifamily"], "term": "5-10 yr fixed", "amort": "30 yr", "recourse": "Non-recourse",
        "source": "Multi-Family USA rates page (spreads as of Nov 1, 2026): Fannie DUS / Freddie Optigo +105-165 bps, up to 80% LTV, 1.25x DCR",
    },
    "Bank balance sheet": {
        "benchmark": "UST5", "spreads": [150, 175, 200, None], "max_ltv": 0.70, "min_dscr": 1.25,
        "property_types": None, "term": "5 yr fixed", "amort": "25-30 yr", "recourse": "Partial/full recourse",
        "source": "Multi-Family USA (Nov 2026): bank 5-yr fixed +150-225 bps, up to ~70% LTV; INSIGNIA (Sep 2026): banks 5.50-6.50%",
    },
    "CMBS conduit": {
        "benchmark": "UST10", "spreads": [175, 190, 210, None], "max_ltv": 0.70, "min_dscr": 1.25,
        "property_types": None, "term": "10 yr fixed", "amort": "30 yr or IO", "recourse": "Non-recourse",
        "source": "Multi-Family USA (Nov 2026): CMBS 10-yr +175-225 bps, up to ~70% LTV, debt yield ~9-10%",
    },
    "Debt fund (bridge)": {
        "benchmark": "SOFR", "spreads": [275, 325, 400, 500], "max_ltv": 0.80, "min_dscr": 1.00,
        "property_types": None, "term": "2-3 yr floating + extensions", "amort": "IO", "recourse": "Non-recourse",
        "source": "Multi-Family USA (Nov 2026): bridge SOFR +275-500 bps; INSIGNIA (Sep 2026): debt funds SOFR +275-375",
    },
    "DSCR loan (investor)": {
        "benchmark": "UST2", "spreads": [150, 200, 250, 300], "max_ltv": 0.80, "min_dscr": 1.00,
        "property_types": None, "term": "30 yr (fixed/ARM)", "amort": "30 yr", "recourse": "Usually non-recourse",
        "source": "Ridge Street Capital (Nov 6, 2026): DSCR rate = 2-yr UST + premium, typical 6.0-7.99%; The Close (Aug 2026): max 80% LTV",
    },
}

CONSTRUCTION_LENDERS = {
    "Bank construction": {
        "benchmark": "SOFR", "spreads": [275, 300, None, None], "max_ltc": 0.65,
        "recourse": "Full/partial recourse, burns off at stabilization",
        "source": "Core Insights Review (Jul 2026): banks SOFR +275-400 bps at 60-65% LTC; Multi-Housing News: bank quotes low 200s",
    },
    "Debt fund construction": {
        "benchmark": "SOFR", "spreads": [400, 425, 475, 550], "max_ltc": 0.85,
        "recourse": "Often non-recourse",
        "source": "Core Insights Review (Jul 2026): debt funds SOFR +400-650 bps at 65-85% LTC",
    },
}


def band_index(leverage, bands):
    for i, top in enumerate(bands):
        if leverage <= top + 1e-9:
            return i
    return None


def fetch_benchmarks(timeout=10):
    """Latest values from FRED. Uses the FRED API when FRED_API_KEY is set (Replit Secrets),
    otherwise FRED's public CSV download. Falls back to FALLBACK_BENCHMARKS per series."""
    import requests

    key = os.environ.get("FRED_API_KEY")
    out = {}
    for name, fb in FALLBACK_BENCHMARKS.items():
        sid = fb["fred_id"]
        try:
            if key:
                r = requests.get("https://api.stlouisfed.org/fred/series/observations",
                                 params={"series_id": sid, "api_key": key, "file_type": "json",
                                         "sort_order": "desc", "limit": 10}, timeout=timeout)
                r.raise_for_status()
                obs = [o for o in r.json()["observations"] if o["value"] not in (".", "")]
                val, dt = float(obs[0]["value"]) / 100, obs[0]["date"]
                src = "FRED API"
            else:
                r = requests.get(f"https://fred.stlouisfed.org/graph/fredgraph.csv?id={sid}", timeout=timeout)
                r.raise_for_status()
                rows = [row for row in csv.reader(io.StringIO(r.text))][1:]
                rows = [row for row in rows if len(row) == 2 and row[1] not in (".", "")]
                dt, val = rows[-1][0], float(rows[-1][1]) / 100
                src = "FRED CSV"
            out[name] = {"value": val, "as_of": dt, "fred_id": sid, "source": src}
        except Exception:
            out[name] = dict(fb)
    return out


def indicative_rates(benchmarks, leverage, property_type, path="permanent"):
    """Indicative all-in rate for every lender type at the deal's leverage."""
    rows = []
    lenders = CONSTRUCTION_LENDERS if path == "construction" else PERM_LENDERS
    bands = CONSTRUCTION_BANDS if path == "construction" else PERM_BANDS
    i = band_index(leverage, bands)
    for name, L in lenders.items():
        bm = benchmarks[L["benchmark"]]["value"]
        spread = L["spreads"][i] if i is not None else None
        max_lev = L.get("max_ltv", L.get("max_ltc"))
        reasons = []
        if spread is None or leverage > max_lev + 1e-9:
            reasons.append(f"leverage above {max_lev:.0%} max")
        types = L.get("property_types")
        if types and property_type not in types:
            reasons.append(f"{'/'.join(types)} only")
        rows.append({
            "lender_type": name, "benchmark": L["benchmark"], "benchmark_rate": bm,
            "spread_bps": spread, "rate": bm + spread / 1e4 if spread is not None else None,
            "eligible": not reasons, "note": "; ".join(reasons),
            "term": L.get("term", ""), "amort": L.get("amort", ""), "recourse": L.get("recourse", ""),
            "source": L["source"],
        })
    return rows


if __name__ == "__main__":
    for r in indicative_rates(FALLBACK_BENCHMARKS, 0.65, "Multifamily"):
        print(r["lender_type"], r["rate"], r["eligible"], r["note"])
    print(date.today())
