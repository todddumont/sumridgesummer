#!/usr/bin/env python3
r"""
demand_technical_engine.py
SMRD -- Credit Demand Technical Engine (TRACE + ICE BofA constituent reconstruction)




Builds a proxy for the "credit demand technical" framework used in sell-side
Global Credit Trader-style research, using only data SMRD actually has:
  - FINRA TRACE 30-day rolling customer trade tape
  - ICE BofA index constituent HISTORY (H0A0 HY, optionally C0A0 IG)




WHAT THIS REPLICATES (from TRACE + constituent history alone):
  - Net dealer lift / net customer demand, weekly, by index
  - Turnover ratio (customer volume / amt outstanding), weekly, by index
  - Tenor-bucket demand preference (short/intermediate/long), by index
  - Rating-cohort net demand, if a rating column is present
  - Coupon income vs net par-outstanding change ("coupon reinvestment vs
    net supply" proxy), across constituent-history snapshot dates. For a
    HY-only index this change also nets out fallen-angel/rising-star
    migration, same as GS's "net rating migration + organic net issuance"
    decomposition -- we can't split the two without more data, but the
    combined delta is still a real technical-supply signal.
  - A composite weekly "demand technical" reading per index




WHAT THIS DOES NOT REPLICATE (SMRD has no feed for these):
  - EPFR fund flows (ETF/mutual fund, active/passive, by duration bucket)
  - US Treasury TIC foreign-purchase data
  - New issue concession estimates
  - DTCC CDX/iTraxx net notional positioning




Because of the above, treat this as a narrower, TRACE-anchored proxy, not a
one-for-one replication of a full sell-side demand-technical model. The
composite score below = 2-3 inputs, not the 8 GS uses in its weekly
sentiment indicator.




Because the rolling TRACE window is only ~30 calendar days (~4-5 weekly
buckets) by default -- or up to 18 months if you point --trace-glob at a
longer history and leave --trace-months-back at its default -- any
percentile-style read here is computed IN-SAMPLE against whatever weeks
were actually loaded, not against arbitrary multi-year history. Treat the
composite reading as a directional read, not a statistically robust score
-- this is flagged explicitly in the output JSON via "sample_weeks".
--trace-months-back (default 18) caps how far back trades are kept
regardless of what's in the glob'd files; pass 0 to disable the cap.




DATA ASSUMPTION FLAGGED FOR REVIEW -- CONTRAPARTYTYPE:
SMRD's TRACE dump uses a 4-value CONTRAPARTYTYPE (C, D, A, T), not the
standard 2-value C/D. This engine treats ONLY "C" as customer flow (net
dealer lift = C-side trades); "D" (confirmed interdealer), "A", and "T"
are all excluded. If "A" turns out to mean agency-executed customer orders
rather than affiliate/other, that's real customer flow being wrongly
excluded -- confirm the code meaning with ops before trusting net_lift_mm
for anything more than a directional read. Row counts by CONTRAPARTYTYPE
are printed to console on every run so this is easy to sanity-check.




FILE FORMAT NOTES:
  - ICE H0A0/C0A0 files have a 2-3 line preamble before the real header
    (e.g. "Using July 2026 Universe" / "Report ICE_H0A0 RunDate ..."). The
    loader auto-detects the header row by scanning for "cusip" as a column
    token, so this doesn't need a fixed --skiprows.
  - ICE "Face Value LOC" = amount outstanding in $mm already (used as
    amt_out_mm directly, no conversion).
  - ICE CUSIPs are 8-char (no check digit); TRACE CUSIPs are typically
    9-char. Both are truncated to the first 8 chars before merging.
  - Each ICE_H0A0_YYYYMMDD.csv is ONE dated snapshot with its own
    "As of Date" column -- pass a glob (not a single file) to
    --constituents so multiple dates load as real history for the
    coupon-vs-supply panel.




USAGE
-----
python demand_technical_engine.py ^
    --trace-glob "P:\30d trace\*.csv" ^
    --constituents H0A0="P:\jmorris\ICE H0A0 Historical Index Data\ICE_H0A0_*.csv" ^
    --out demand_technical.json




If your file headers differ from the defaults below, edit the COLUMN MAP
constants -- don't fight the parser with CLI flags for every column.
"""




import argparse
import glob as globmod
import json
import re
import sys
import time
from datetime import datetime, timezone
from pathlib import Path




import numpy as np
import pandas as pd




_START_TIME = time.time()








def log(msg):
    """Print with an elapsed-time prefix and force an immediate flush -- without this,
    Python buffers stdout when it's not attached to an interactive terminal (common on
    Windows when run via a shortcut/batch file), so all output appears to arrive at once
    at the very end even though the script is progressing normally."""
    print(f"[{time.time() - _START_TIME:6.1f}s] {msg}", flush=True)




_FILENAME_DATE_RE = re.compile(r"(20\d{2})[-_]?(\d{2})[-_]?(\d{2})")








def _extract_file_date(path):
    """Pull a date out of the filename (YYYYMMDD, YYYY-MM-DD, YYYY_MM_DD) so old
    files can be skipped WITHOUT opening/parsing them. Falls back to file mtime
    if no date pattern is found in the name."""
    m = _FILENAME_DATE_RE.search(Path(path).name)
    if m:
        y, mo, d = m.groups()
        try:
            return pd.Timestamp(year=int(y), month=int(mo), day=int(d))
        except ValueError:
            pass
    try:
        return pd.Timestamp(Path(path).stat().st_mtime, unit="s").normalize()
    except OSError:
        return None




# ---------------------------------------------------------------------------
# COLUMN MAPS -- edit these if your TRACE dump / ICE constituent file headers
# differ. Keys are canonical names used internally; values are the raw
# column name(s) to search for (first match wins, case-sensitive).
# ---------------------------------------------------------------------------
TRACE_COLS = {
    "cusip":       ["cusip", "CUSIP", "cusip_id"],
    "trade_date":  ["trade_date", "trd_exctn_dt", "TradeDate", "EXECUTIONDATE"],
    "rpt_side_cd": ["rpt_side_cd", "RptSideCd", "side", "REPORTINGPARTYSIDE"],   # B = dealer bought, S = dealer sold, D = interdealer (dropped)
    "contra_type": ["contra_type", "cntra_mp_id", "ContraPartyType", "CONTRAPARTYTYPE"],  # C = customer (see docstring re: A/T)
    "qty":         ["qty", "entrd_vol_qt", "quantity", "par_value", "QUANTITY"],
    "price":       ["price", "rptd_pr", "Price", "PRICE"],
}




# Optional TRACE quality/scope filters -- applied only if the column exists.
TRACE_OPTIONAL_COLS = {
    "subproduct_type": ["SUBPRODUCTTYPE", "sub_prdct"],   # kept only where == "CORP"
    "status":          ["STATUS"],                        # kept only where blank/NaN
    "asof_indicator":  ["ASOFINDICATOR", "asof_cd"],       # kept only where blank/NaN (drops as-of/reversal prints)
}




CONSTIT_COLS = {
    "cusip":         ["cusip", "CUSIP", "Cusip"],
    "as_of_date":    ["as_of_date", "AsOfDate", "index_date", "date", "As of Date"],
    "issuer":        ["issuer", "Issuer", "issuer_name", "Description"],   # ICE has no clean issuer field; bond description used as fallback
    "ticker":        ["ticker", "Ticker"],
    "coupon":        ["coupon", "Coupon", "cpn", "Par Wtd Coupon"],
    "maturity_date": ["maturity_date", "Maturity", "maturity", "Maturity Date"],
    "amt_out_mm":    ["amt_out_mm", "AmtOutstanding", "amount_outstanding_mm", "par_amt_mm", "Face Value LOC"],
    "duration":      ["duration", "Duration", "eff_dur", "oad", "Spread Duration"],
    "rating":        ["rating", "Rating", "composite_rating"],
}








def _find_header_row(path, key_token="cusip", max_scan=10):
    """ICE files have a 2-3 line preamble ('Using <month> Universe', 'Report ... RunDate ...')
    before the real header row. Scan the first few lines for a row containing a 'cusip'-like
    token and use that as the header row. Falls back to row 0 (no preamble) if not found."""
    try:
        with open(path, "r", encoding="utf-8", errors="replace") as f:
            for i in range(max_scan):
                line = f.readline()
                if not line:
                    break
                tokens = [t.strip().strip('"').lower() for t in line.split(",")]
                if key_token in tokens:
                    return i
    except OSError:
        pass
    return 0








def _pick_col(df_cols, candidates):
    for c in candidates:
        if c in df_cols:
            return c
    return None








def _normalize(df, colmap, label):
    out = {}
    missing = []
    for canon, candidates in colmap.items():
        found = _pick_col(df.columns, candidates)
        if found is not None:
            out[canon] = df[found]
        else:
            missing.append(canon)
    norm = pd.DataFrame(out)
    if missing:
        print(f"[{label}] WARNING -- columns not found under any alias: {missing}", file=sys.stderr, flush=True)
    return norm








def load_trace(trace_glob, qty_multiplier, months_back):
    files = sorted(globmod.glob(trace_glob))
    log(f"[TRACE] glob {trace_glob!r} matched {len(files)} file(s)")
    if not files:
        raise FileNotFoundError(f"No files matched --trace-glob {trace_glob!r}")




    cutoff = None
    if months_back is not None:
        cutoff = pd.Timestamp.now().normalize() - pd.DateOffset(months=int(months_back))
        kept, skipped = [], 0
        for f in files:
            fdate = _extract_file_date(f)
            if fdate is not None and fdate < cutoff:
                skipped += 1
                continue
            kept.append(f)
        log(f"[TRACE] file-level date filter (cutoff {cutoff.strftime('%Y-%m-%d')}): "
            f"skipped {skipped} file(s) without opening them, {len(kept)}/{len(files)} kept")
        files = kept
        if not files:
            raise FileNotFoundError(
                f"All files matched by --trace-glob are older than the {months_back}-month cutoff "
                f"({cutoff.strftime('%Y-%m-%d')}). Check the glob path or --trace-months-back."
            )




    frames = []
    n_files = len(files)
    running_rows = 0
    for i, f in enumerate(files, 1):
        header_row = _find_header_row(f)
        raw = pd.read_csv(f, skiprows=header_row, low_memory=False)
        running_rows += len(raw)
        norm = _normalize(raw, TRACE_COLS, f"TRACE:{Path(f).name}")
        for canon, candidates in TRACE_OPTIONAL_COLS.items():
            found = _pick_col(raw.columns, candidates)
            if found is not None:
                norm[canon] = raw[found]
        frames.append(norm)
        log(f"[TRACE] read {i}/{n_files}: {Path(f).name} ({len(raw):,} rows, {running_rows:,} total so far)")
    trace = pd.concat(frames, ignore_index=True)




    trace["cusip"] = trace["cusip"].astype(str).str.strip().str.upper().str[:8]
    trace["trade_date"] = pd.to_datetime(trace["trade_date"], errors="coerce")
    trace["contra_type"] = trace["contra_type"].astype(str).str.strip().str.upper()
    trace["rpt_side_cd"] = trace["rpt_side_cd"].astype(str).str.strip().str.upper().str[0]
    trace["qty"] = pd.to_numeric(trace["qty"], errors="coerce") * qty_multiplier




    before = len(trace)
    trace = trace.dropna(subset=["cusip", "trade_date", "qty"])




    if cutoff is not None:
        n0 = len(trace)
        trace = trace[trace["trade_date"] >= cutoff]
        log(f"[TRACE] row-level lookback filter (cutoff {cutoff.strftime('%Y-%m-%d')}, safety net "
            f"in case a kept file spans multiple dates): {n0:,} -> {len(trace):,}")




    log(f"[TRACE] contra_type breakdown pre-filter: {trace['contra_type'].value_counts().to_dict()}")
    trace = trace[trace["contra_type"] == "C"]  # customer trades only -- see docstring re: A/T codes




    if "subproduct_type" in trace:
        n0 = len(trace)
        trace = trace[trace["subproduct_type"].astype(str).str.upper() == "CORP"]
        log(f"[TRACE] subproduct_type filter (CORP only): {n0:,} -> {len(trace):,}")
    if "status" in trace:
        n0 = len(trace)
        trace = trace[trace["status"].isna() | (trace["status"].astype(str).str.strip() == "")]
        log(f"[TRACE] status filter (blank only): {n0:,} -> {len(trace):,}")
    if "asof_indicator" in trace:
        n0 = len(trace)
        trace = trace[trace["asof_indicator"].isna() | (trace["asof_indicator"].astype(str).str.strip() == "")]
        log(f"[TRACE] asof_indicator filter (blank only, drops as-of/reversal prints): {n0:,} -> {len(trace):,}")




    log(f"[TRACE] loaded {before:,} rows across {len(files)} file(s) -> {len(trace):,} usable customer-side rows")
    return trace








def load_constituents(pairs):
    """pairs: dict of index_id -> glob pattern (e.g. 'ICE_H0A0_*.csv'). Each matched file is
    treated as one dated snapshot; its 'As of Date' column supplies the snapshot date."""
    frames = []
    for index_id, pattern in pairs.items():
        files = sorted(globmod.glob(pattern))
        log(f"[CONSTIT] glob {index_id}={pattern!r} matched {len(files)} file(s)")
        if not files:
            raise FileNotFoundError(f"No files matched --constituents {index_id}={pattern!r}")
        for i, f in enumerate(files, 1):
            header_row = _find_header_row(f)
            raw = pd.read_csv(f, skiprows=header_row, low_memory=False)
            norm = _normalize(raw, CONSTIT_COLS, f"CONSTIT:{index_id}:{Path(f).name}")
            norm["index_id"] = index_id
            frames.append(norm)
            log(f"[CONSTIT] {index_id} read {i}/{len(files)}: {Path(f).name} ({len(raw):,} rows)")
    constit = pd.concat(frames, ignore_index=True)




    constit["cusip"] = constit["cusip"].astype(str).str.strip().str.upper().str[:8]
    constit["as_of_date"] = pd.to_datetime(constit["as_of_date"], errors="coerce")
    constit["amt_out_mm"] = pd.to_numeric(constit["amt_out_mm"], errors="coerce")
    if "coupon" in constit:
        constit["coupon"] = pd.to_numeric(constit["coupon"], errors="coerce")
    if "maturity_date" in constit:
        constit["maturity_date"] = pd.to_datetime(constit["maturity_date"], errors="coerce")
    if "duration" in constit:
        constit["duration"] = pd.to_numeric(constit["duration"], errors="coerce")




    before = len(constit)
    constit = constit.dropna(subset=["cusip", "as_of_date", "amt_out_mm"])
    log(f"[CONSTIT] loaded {before:,} rows -> {len(constit):,} usable rows across {list(pairs.keys())}")
    return constit








def latest_snapshot(constit):
    return (
        constit.sort_values("as_of_date")
        .groupby(["index_id", "cusip"])
        .tail(1)
        .reset_index(drop=True)
    )








def tenor_bucket(row, short_max, interm_max):
    if pd.notna(row.get("duration")):
        yrs = row["duration"]
    elif pd.notna(row.get("maturity_date")) and pd.notna(row.get("as_of_date")):
        yrs = (row["maturity_date"] - row["as_of_date"]).days / 365.25
    else:
        return None
    if yrs < short_max:
        return "Short"
    elif yrs < interm_max:
        return "Intermediate"
    return "Long"








def parse_constituents_arg(pairs_list):
    out = {}
    for item in pairs_list:
        if "=" not in item:
            raise ValueError(f"--constituents entries must be INDEX_ID=path.csv, got {item!r}")
        k, v = item.split("=", 1)
        out[k.strip().upper()] = v.strip()
    return out








def clean(df):
    if df is None or df.empty:
        return []
    return df.astype(object).where(pd.notnull(df), None).to_dict("records")








def main():
    ap = argparse.ArgumentParser(description="SMRD demand technical engine (TRACE + ICE constituent proxy)")
    ap.add_argument("--trace-glob", required=True, help='e.g. "P:\\30d trace\\*.csv"')
    ap.add_argument("--constituents", nargs="+", default=[
                         r'H0A0=P:\jmorris\ICE H0A0 Historical Index Data\ICE_H0A0_*.csv',
                         r'C0A0=P:\jmorris\ICE C0A0 Historical Index Data\ICE_C0A0_*.csv',
                     ],
                     help='INDEX_ID=glob pairs, e.g. H0A0="ICE_H0A0_*.csv" C0A0="ICE_C0A0_*.csv" '
                          '-- each matched file is one dated snapshot. Defaults to both the H0A0 and '
                          'C0A0 folders under P:\\jmorris\\ if not specified -- pass this flag to '
                          'override (e.g. to run just one index).')
    ap.add_argument("--out", default="demand_technical.json")
    ap.add_argument("--short-max", type=float, default=3.0, help="years; below this = Short bucket")
    ap.add_argument("--interm-max", type=float, default=7.0, help="years; below this (and above short-max) = Intermediate")
    ap.add_argument("--week-freq", default="W-WED",
                     help="pandas period alias for weekly bucketing (default W-WED, matches TRACE Wed-to-Wed convention)")
    ap.add_argument("--qty-multiplier", type=float, default=1.0,
                     help="multiply raw TRACE qty by this to get $ par (e.g. 1000 if your dump reports qty in $000s)")
    ap.add_argument("--trace-months-back", type=float, default=18.0,
                     help="only keep TRACE trades within this many months of run time (default 18; use 0 to disable)")
    args = ap.parse_args()
    log(f"Starting run: trace_glob={args.trace_glob!r} constituents={args.constituents} months_back={args.trace_months_back}")




    pairs = parse_constituents_arg(args.constituents)
    trace = load_trace(args.trace_glob, args.qty_multiplier, args.trace_months_back if args.trace_months_back > 0 else None)
    constit = load_constituents(pairs)




    snap = latest_snapshot(constit)
    snap["tenor"] = snap.apply(lambda r: tenor_bucket(r, args.short_max, args.interm_max), axis=1)
    if "rating" in snap:
        # collapse notch-level ratings (B1, BB2, ...) to letter buckets (B, BB, ...) for cohort grouping
        snap["rating"] = snap["rating"].astype(str).str.replace(r"\d+$", "", regex=True).replace({"nan": None})




    trace_before = len(trace)
    trace = trace.merge(snap[["cusip", "index_id", "amt_out_mm", "tenor", "rating"]], on="cusip", how="inner")
    match_rate = len(trace) / trace_before if trace_before else 0.0
    log(f"[MATCH] {len(trace):,}/{trace_before:,} customer trades matched to index constituents ({match_rate:.1%})")
    if trace.empty:
        print("ERROR: no TRACE rows matched any constituent CUSIP -- check COLUMN MAP / cusip formatting / date overlap.",
              file=sys.stderr, flush=True)
        sys.exit(1)




    trace["buy_mm"] = np.where(trace["rpt_side_cd"] == "S", trace["qty"], 0.0) / 1e6
    trace["sell_mm"] = np.where(trace["rpt_side_cd"] == "B", trace["qty"], 0.0) / 1e6
    trace["net_signed_mm"] = trace["buy_mm"] - trace["sell_mm"]
    trace["week_ending"] = trace["trade_date"].dt.to_period(args.week_freq).apply(lambda p: p.end_time.normalize())




    weekly = trace.groupby(["index_id", "week_ending"]).agg(
        cust_buy_mm=("buy_mm", "sum"),
        cust_sell_mm=("sell_mm", "sum"),
    ).reset_index()
    weekly["net_lift_mm"] = weekly["cust_buy_mm"] - weekly["cust_sell_mm"]
    weekly["total_vol_mm"] = weekly["cust_buy_mm"] + weekly["cust_sell_mm"]




    universe = snap.groupby("index_id")["amt_out_mm"].sum().rename("universe_amt_out_mm").reset_index()
    weekly = weekly.merge(universe, on="index_id")
    weekly["turnover_pct"] = weekly["total_vol_mm"] / weekly["universe_amt_out_mm"] * 100
    weekly = weekly.sort_values(["index_id", "week_ending"])




    tenor_agg = trace.dropna(subset=["tenor"]).groupby(["index_id", "tenor"]).agg(
        volume_mm=("qty", lambda s: s.sum() / 1e6),
        net_lift_mm=("net_signed_mm", "sum"),
    ).reset_index()
    if not tenor_agg.empty:
        tenor_agg["volume_share_pct"] = tenor_agg.groupby("index_id")["volume_mm"].transform(lambda s: s / s.sum() * 100)




    rating_agg = pd.DataFrame()
    if trace["rating"].notna().mean() > 0.5:
        rating_agg = trace.dropna(subset=["rating"]).groupby(["index_id", "rating"]).agg(
            volume_mm=("qty", lambda s: s.sum() / 1e6),
            net_lift_mm=("net_signed_mm", "sum"),
        ).reset_index()




    snap_dates = (
        constit.groupby(["index_id", "as_of_date"])["amt_out_mm"].sum()
        .rename("total_amt_out_mm").reset_index()
        .sort_values(["index_id", "as_of_date"])
    )
    coup = constit.copy()
    coup["coupon_income_mm"] = coup["amt_out_mm"] * coup.get("coupon", pd.Series(dtype=float)).fillna(0) / 100.0
    coup_agg = coup.groupby(["index_id", "as_of_date"])["coupon_income_mm"].sum().reset_index()
    cvs = snap_dates.merge(coup_agg, on=["index_id", "as_of_date"])
    cvs["net_supply_mm"] = cvs.groupby("index_id")["total_amt_out_mm"].diff()
    cvs["coverage_ratio"] = np.where(cvs["net_supply_mm"] > 0, cvs["coupon_income_mm"] / cvs["net_supply_mm"], np.nan)




    def pct_rank(series, value):
        return float((series <= value).mean() * 100)




    # trace["week_ending"] labels a week by its calendar end date regardless of how
    # much of that week is actually covered by the TRACE data on hand. Using a
    # partial trailing week as "latest" understates net_lift_mm/turnover_pct and
    # skews the percentile read. Drop any trailing week(s) whose week_ending falls
    # after the last date actually present in the TRACE tape, so "latest" always
    # means the most recent COMPLETE week.
    max_trade_date = trace["trade_date"].max()




    composite = []
    for index_id, g in weekly.groupby("index_id"):
        g = g.sort_values("week_ending")
        g_complete = g[g["week_ending"] <= max_trade_date]
        if g_complete.empty:
            g_complete = g  # only a partial week exists at all -- use it rather than drop the index
        g = g_complete
        latest = g.iloc[-1]
        lift_pct = pct_rank(g["net_lift_mm"], latest["net_lift_mm"])
        turn_pct = pct_rank(g["turnover_pct"], latest["turnover_pct"])
        cvs_i = cvs[cvs["index_id"] == index_id].dropna(subset=["coverage_ratio"])
        coverage_latest = float(cvs_i["coverage_ratio"].iloc[-1]) if len(cvs_i) else None
        composite_pct = float(np.mean([lift_pct, turn_pct]))
        label = "Supportive" if composite_pct >= 65 else ("Weak" if composite_pct <= 35 else "Neutral")
        composite.append({
            "index_id": index_id,
            "week_ending": latest["week_ending"].strftime("%Y-%m-%d"),
            "sample_weeks": int(len(g)),
            "net_lift_mm": round(float(latest["net_lift_mm"]), 1),
            "net_lift_percentile": round(lift_pct, 1),
            "turnover_pct": round(float(latest["turnover_pct"]), 2),
            "turnover_percentile": round(turn_pct, 1),
            "coupon_supply_coverage": round(coverage_latest, 2) if coverage_latest is not None else None,
            "composite_percentile": round(composite_pct, 1),
            "reading": label,
        })




    weekly_out = weekly.copy()
    weekly_out["week_ending"] = weekly_out["week_ending"].dt.strftime("%Y-%m-%d")
    cvs_out = cvs.copy()
    cvs_out["as_of_date"] = cvs_out["as_of_date"].dt.strftime("%Y-%m-%d")




    out = {
        "meta": {
            "generated_at": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
            "indices": list(pairs.keys()),
            "trace_date_range": [
                trace["trade_date"].min().strftime("%Y-%m-%d"),
                trace["trade_date"].max().strftime("%Y-%m-%d"),
            ],
            "trace_rows_matched": int(len(trace)),
            "match_rate_pct": round(match_rate * 100, 1),
            "week_ending_convention": args.week_freq,
            "trace_months_back": args.trace_months_back if args.trace_months_back > 0 else None,
            "tenor_breaks_years": {"short_max": args.short_max, "intermediate_max": args.interm_max},
            "methodology_note": (
                "TRACE customer-side net lift + ICE BofA constituent history only. Does not include "
                "EPFR fund flows, TIC foreign demand, new issue concessions, or DTCC CDX/iTraxx "
                "positioning -- SMRD has no feed for those. See engine header comment for detail."
            ),
        },
        "weekly": clean(weekly_out),
        "tenor_buckets": clean(tenor_agg),
        "rating_cohort": clean(rating_agg),
        "coupon_vs_supply": clean(cvs_out),
        "composite": composite,
    }




    Path(args.out).write_text(json.dumps(out, indent=2))
    log(f"Wrote {args.out} ({len(weekly)} weekly rows, {len(tenor_agg)} tenor rows, {len(cvs)} coupon/supply snapshots)")








if __name__ == "__main__":
    main()
