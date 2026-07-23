#!/usr/bin/env python3 

r""" 

trace_digest.py 

 

Standalone TRACE market-context digest for the HY desk / RV + Account Grader site. 

Runs independently of the account grader (own script, own schedule), but stays 

"in H0A0": it loads your ICE H0A0 index reference and tags/enriches every printing 

CUSIP with its index data, so the digest is scoped to your investable universe. 

 

TRACE is FINRA's anonymized print tape -- NO counterparty identity -- so this is 

market context only. Per CUSIP it captures liquidity (print count / active days / 

adjusted size), a trade-based G-spread (YIELD - BENCHMARKYIELD) and its drift, then 

attaches H0A0 fields: in-index flag, ticker, sector, rating, tenor, index OAS. 

 

Inputs (defaults; override with --trace / --hy): 

  - TRACE folder:  P:\30d trace              (daily csv/xlsx files, rolling ~30d) 

  - H0A0 ref:      P:\jmorris\ICE H0A0 Historical Index Data 

 

Output: 

  - account_grader_output\trace_digest.csv   (read by trace_monitor.html) 

 

This script reuses the grader's tested parsers, so it must sit in the SAME folder 

as account_grader_hy_account_profile_v2.py. 

""" 

 

import argparse 

import bisect 

import csv 

import os 

import re 

import sys 

import time 

from collections import Counter, defaultdict 

from datetime import date, timedelta 

 

try: 

    from account_grader_hy_account_profile_v2 import ( 

        num, excel_date, norm_cusip, key_header, header_index, get, 

        iter_xlsx_rows, read_ice_source, 

    ) 

except ImportError as e: 

    print( 

        "ERROR: could not import shared helpers from account_grader_hy_account_profile_v2.py.\n" 

        "       Keep trace_digest.py in the SAME folder as the account grader script.\n" 

        f"       ({e})", 

        file=sys.stderr, 

    ) 

    sys.exit(1) 

 

DEFAULT_TRACE = r"P:\30d trace" 

DEFAULT_HY = r"P:\jmorris\ICE H0A0 Historical Index Data" 

DEFAULT_OUT = "account_grader_output" 

 

# TRACE column names (matched case-insensitively; filler columns are ignored). 

TRACE_COL_DATE = "ENTRYDATE" 

TRACE_COL_CUSIP = "CUSIP" 

TRACE_COL_QTY = "QUANTITY" 

TRACE_COL_ADJQTY = "ADJUSTQUANTITY" 

TRACE_COL_PRICE = "PRICE" 

TRACE_COL_YIELD = "YIELD" 

TRACE_COL_BMK_YLD = "BENCHMARKYIELD" 

TRACE_COL_BMK = "BENCHMARK ISSUE" 

TRACE_COL_SIDE = "ReportingPartySide" 

 

# The side column's exact header varies by export (ReportingPartySide, ReportingPartSide, 

# "Reporting Party Side", RPTSIDE, ...). Match on a normalized form so spelling/spacing 

# differences don't silently blank out the whole client-flow read. 

SIDE_COL_CANDIDATES = ( 

    "REPORTINGPARTYSIDE", "REPORTINGPARTSIDE", "REPORTINGSIDE", 

    "RPTPARTYSIDE", "RPTSIDE", "PARTYSIDE", "SIDE", "BUYSELLIND", "BUYSELL", 

) 

 

 

def _norm_key(x): 

    return re.sub(r"[^A-Z0-9]", "", str(x or "").upper()) 

 

 

def find_side_col(idx): 

    """Return the actual header name for the dealer-side column, or '' if absent.""" 

    norm = {_norm_key(k): k for k in idx.keys()} 

    for cand in SIDE_COL_CANDIDATES: 

        if cand in norm: 

            return norm[cand] 

    # last resort: any header that ends in SIDE 

    for nk, orig in norm.items(): 

        if nk.endswith("SIDE"): 

            return orig 

    return "" 

 

 

def client_side(v): 

    """TRACE ReportingPartySide is the DEALER's side. Translate to the CLIENT's side: 

    dealer Bought -> client SOLD (supply); dealer Sold -> client BOUGHT (demand). 

    'D' = inter-dealer print (no client), intentionally ignored. 

    Returns 'buy' / 'sell' from the client's perspective, or None (skip).""" 

    x = str(v or "").strip().upper() 

    if not x: 

        return None 

    if x.startswith("D"): 

        return None     # inter-dealer, no client flow 

    if x.startswith("B"): 

        return "sell"   # dealer bought => client sold 

    if x.startswith("S"): 

        return "buy"    # dealer sold => client bought 

    return None 

 

 

def _median(vals): 

    vals = sorted(v for v in vals if v is not None) 

    n = len(vals) 

    if n == 0: 

        return None 

    mid = n // 2 

    if n % 2: 

        return vals[mid] 

    return (vals[mid - 1] + vals[mid]) / 2.0 

 

 

def new_trace_agg(): 

    return { 

        "prints": 0, 

        "dates": set(), 

        "volume": 0.0, 

        "prices": [], 

        "yields": [], 

        "spreads": [],          # (date, gspread) pairs 

        "last_date": "", 

        "last_price": None, 

        "last_yield": None, 

        "last_spread": None, 

        "benchmarks": Counter(), 

        "cbuy": 0, "csell": 0, "cbuy_vol": 0.0, "csell_vol": 0.0, "side_prints": 0, 

    } 

 

 

def date_from_filename(name): 

    """Best-effort trade date from a TRACE filename (YYYYMMDD, YYYY-MM-DD, MM-DD-YYYY, ...).""" 

    m = re.search(r"(20\d{2})[-_.]?(\d{2})[-_.]?(\d{2})", name)  # YYYYMMDD family 

    if m: 

        try: 

            return date(int(m.group(1)), int(m.group(2)), int(m.group(3))) 

        except ValueError: 

            pass 

    m = re.search(r"(\d{2})[-_.](\d{2})[-_.](20\d{2})", name)     # MM-DD-YYYY family 

    if m: 

        try: 

            return date(int(m.group(3)), int(m.group(1)), int(m.group(2))) 

        except ValueError: 

            pass 

    return None 

 

 

def select_recent_trace_files(folder, days): 

    """Pick only the last `days` of TRACE files so old history in the folder is ignored. 

 

    Selection is by date parsed from the filename, so files outside the window are 

    never opened (fast, and it fixes folders that quietly hold years of daily files). 

    If filenames carry no parseable date, falls back to the newest `days` files by 

    name order. Returns (files, cutoff_iso, newest_iso, mode). 

    """ 

    all_files = [f for f in os.listdir(folder) if f.lower().endswith((".csv", ".xlsx"))] 

    dated = [(f, date_from_filename(f)) for f in all_files] 

    have = [d for _, d in dated if d] 

    if have: 

        newest = max(have) 

        cutoff = newest - timedelta(days=days) 

        kept = sorted(f for f, d in dated if d and d >= cutoff) 

        kept += sorted(f for f, d in dated if not d)  # keep undated; row-level cutoff still trims them 

        return kept, cutoff.isoformat(), newest.isoformat(), "filename date" 

    return sorted(all_files)[-days:], "", "", "name order (no dates found in filenames)" 

 

 

def process_trace_folder(folder, days=35, sample_rows=0): 

    files, cutoff_iso, newest_iso, mode = select_recent_trace_files(folder, days) 

    total_available = len([f for f in os.listdir(folder) if f.lower().endswith((".csv", ".xlsx"))]) 

    print(f"  window: last {days} days by {mode}" 

          + (f" ({cutoff_iso} .. {newest_iso})" if cutoff_iso else "") 

          + f"; using {len(files)} of {total_available} file(s) in folder") 

 

    agg = defaultdict(new_trace_agg) 

    side_col_used = [""] 

    side_seen = Counter() 

    total_rows = 0 

    used_rows = 0 

    windowed_out = 0 

    sample_shown = 0 

    note = "" 

 

    for f in files: 

        path = os.path.join(folder, f) 

        print(f"  {f}") 

        gen = iter_xlsx_rows(path, sample_rows) 

        try: 

            header = next(gen) 

        except StopIteration: 

            continue 

        idx = header_index(header) 

        side_col = find_side_col(idx) 

        if side_col and side_col != side_col_used[0]: 

            side_col_used[0] = side_col 

            print(f"    side column: {side_col!r} -> client flow enabled") 

        elif not side_col and not side_col_used[0]: 

            print("    !! no dealer-side column found; client buy/sell will be blank.") 

            print(f"       headers seen: {list(idx.keys())[:14]}") 

        if key_header(TRACE_COL_DATE) not in idx or key_header(TRACE_COL_CUSIP) not in idx: 

            note = f"{f}: could not find ENTRYDATE/CUSIP columns (first headers seen: {header[:8]})" 

            print("    NOTE:", note) 

            continue 

 

        for row in gen: 

            total_rows += 1 

            d = excel_date(get(row, idx, TRACE_COL_DATE)) 

            if cutoff_iso and (not d or d < cutoff_iso): 

                windowed_out += 1 

                continue 

            cusip = norm_cusip(get(row, idx, TRACE_COL_CUSIP)) 

            if len(cusip) != 9: 

                continue 

            price = num(get(row, idx, TRACE_COL_PRICE)) 

            yld = num(get(row, idx, TRACE_COL_YIELD)) 

            bmk_yld = num(get(row, idx, TRACE_COL_BMK_YLD)) 

            qty = num(get(row, idx, TRACE_COL_ADJQTY)) 

            if qty is None: 

                qty = num(get(row, idx, TRACE_COL_QTY)) 

            gspread = (yld - bmk_yld) if (yld is not None and bmk_yld is not None) else None 

            raw_side = get(row, idx, side_col) if side_col else "" 

            side_seen[str(raw_side).strip().upper()[:6] or "(blank)"] += 1 

            cside = client_side(raw_side) 

 

            a = agg[cusip] 

            a["prints"] += 1 

            if cside is not None: 

                a["side_prints"] += 1 

                v = abs(qty) if qty is not None else 0.0 

                if cside == "buy": 

                    a["cbuy"] += 1; a["cbuy_vol"] += v 

                else: 

                    a["csell"] += 1; a["csell_vol"] += v 

            if d: 

                a["dates"].add(d) 

            if qty is not None: 

                a["volume"] += abs(qty) 

            if price is not None: 

                a["prices"].append(price) 

            if yld is not None: 

                a["yields"].append(yld) 

            if gspread is not None: 

                a["spreads"].append((d, gspread)) 

            bmk = str(get(row, idx, TRACE_COL_BMK) or "").strip() 

            if bmk: 

                a["benchmarks"][bmk] += 1 

            if d and d >= (a["last_date"] or ""): 

                a["last_date"] = d 

                if price is not None: 

                    a["last_price"] = price 

                if yld is not None: 

                    a["last_yield"] = yld 

                if gspread is not None: 

                    a["last_spread"] = gspread 

            used_rows += 1 

 

            if sample_shown < 3: 

                print(f"    TRACE sample: date={get(row, idx, TRACE_COL_DATE)!r} cusip={cusip} " 

                      f"px={price} yld={yld} gspread={gspread} " 

                      f"dealer_side={raw_side!r} -> client={cside}") 

                sample_shown += 1 

 

    if side_seen: 

        top = ", ".join(f"{k}={v:,}" for k, v in side_seen.most_common(6)) 

        print(f"    dealer-side values: {top}") 

    quality = { 

        "trace_side_column": side_col_used[0], 

        "trace_files_available": total_available, 

        "trace_files_used": len(files), 

        "trace_window_start": cutoff_iso, 

        "trace_window_end": newest_iso, 

        "trace_rows_read": total_rows, 

        "trace_rows_outside_window": windowed_out, 

        "trace_rows_used": used_rows, 

        "trace_cusips": len(agg), 

        "trace_note": note, 

    } 

    return agg, quality 

 

def _flow_label(buy_vol, sell_vol, side_prints): 

    """Net client-flow read for the window. Demand = net client buying (dealer selling).""" 

    if not side_prints: 

        return "" 

    tot = buy_vol + sell_vol 

    if tot <= 0: 

        return "Balanced" 

    ratio = buy_vol / tot 

    if ratio >= 0.60: 

        return "Client buying"     # demand 

    if ratio <= 0.40: 

        return "Client selling"    # supply 

    return "Balanced" 

 

 

def build_trace_digest_rows(agg, ref, in_index_only=False): 

    prints_sorted = sorted(a["prints"] for a in agg.values()) or [0] 

    n = len(prints_sorted) 

 

    def liquidity_score(v): 

        return round(bisect.bisect_right(prints_sorted, v) / n * 10.0, 2) 

 

    def r4(v): 

        return round(v, 4) if v is not None else "" 

 

    rows = [] 

    in_index_count = 0 

    for cusip, a in agg.items(): 

        info = ref.get(cusip) 

        if in_index_only and not info: 

            continue 

        if info: 

            in_index_count += 1 

 

        chron = [g for _, g in sorted(a["spreads"], key=lambda x: (x[0] or ""))] 

        drift = None 

        drift_label = "" 

        if len(chron) >= 6: 

            k = max(1, len(chron) // 3) 

            older = _median(chron[:k]) 

            recent = _median(chron[-k:]) 

            if older is not None and recent is not None: 

                drift = recent - older 

                drift_label = "Tightening" if drift <= -0.05 else "Widening" if drift >= 0.05 else "Flat" 

 

        rows.append({ 

            "cusip": cusip, 

            # ---- H0A0 index enrichment ---- 

            "in_h0a0": 1 if info else 0, 

            "ticker": (info or {}).get("ticker", ""), 

            "sector_l3": (info or {}).get("sector_l3", ""), 

            "rating_bucket": (info or {}).get("rating_bucket", ""), 

            "tenor_bucket": (info or {}).get("tenor_bucket", ""), 

            "index_oas": (info or {}).get("oas", ""), 

            # ---- TRACE market context ---- 

            "trace_prints": a["prints"], 

            "trace_active_days": len(a["dates"]), 

            "trace_volume": round(a["volume"], 0), 

            "trace_liquidity_score": liquidity_score(a["prints"]), 

            "trace_last_date": a["last_date"], 

            "trace_last_price": r4(a["last_price"]), 

            "trace_last_yield": r4(a["last_yield"]), 

            "trace_last_gspread": r4(a["last_spread"]), 

            "trace_median_price": r4(_median(a["prices"])), 

            "trace_median_yield": r4(_median(a["yields"])), 

            "trace_median_gspread": r4(_median([g for _, g in a["spreads"]])), 

            "trace_gspread_drift": r4(drift), 

            "trace_drift_label": drift_label, 

            # ---- client flow (from dealer-side ReportingPartySide) ---- 

            "trace_client_buys": a["cbuy"], 

            "trace_client_sells": a["csell"], 

            "trace_client_buy_vol": round(a["cbuy_vol"], 0), 

            "trace_client_sell_vol": round(a["csell_vol"], 0), 

            "trace_client_net_vol": round(a["cbuy_vol"] - a["csell_vol"], 0), 

            "trace_flow_label": _flow_label(a["cbuy_vol"], a["csell_vol"], a["side_prints"]), 

            "trace_benchmark": a["benchmarks"].most_common(1)[0][0] if a["benchmarks"] else "", 

        }) 

 

    # in-index names first, then by activity 

    rows.sort(key=lambda r: (-r["in_h0a0"], -r["trace_prints"], r["cusip"])) 

    return rows, in_index_count 

 

 

DIGEST_COLS = [ 

    "cusip", "in_h0a0", "ticker", "sector_l3", "rating_bucket", "tenor_bucket", "index_oas", 

    "trace_prints", "trace_active_days", "trace_volume", "trace_liquidity_score", 

    "trace_last_date", "trace_last_price", "trace_last_yield", "trace_last_gspread", 

    "trace_median_price", "trace_median_yield", "trace_median_gspread", 

    "trace_gspread_drift", "trace_drift_label", 

    "trace_client_buys", "trace_client_sells", "trace_client_buy_vol", "trace_client_sell_vol", 

    "trace_client_net_vol", "trace_flow_label", 

    "trace_benchmark", 

] 

 

 

def write_trace_digest(path, rows): 

    with open(path, "w", newline="", encoding="utf-8") as f: 

        w = csv.writer(f) 

        w.writerow(DIGEST_COLS) 

        for r in rows: 

            w.writerow([r.get(c, "") for c in DIGEST_COLS]) 

 

 

def main(): 

    parser = argparse.ArgumentParser(description="Build a per-CUSIP TRACE market-context digest, scoped to the ICE H0A0 index.") 

    parser.add_argument("--trace", default=DEFAULT_TRACE, help="Folder of daily TRACE csv/xlsx files (rolling ~30d).") 

    parser.add_argument("--hy", default=DEFAULT_HY, help="ICE H0A0 csv file or folder (index reference).") 

    parser.add_argument("--out-dir", default=DEFAULT_OUT, help="Output directory (must match where the site reads CSVs).") 

    parser.add_argument("--days", type=int, default=35, help="Only use TRACE files/rows within the last N days (default 35 = ~a month).") 

    parser.add_argument("--sample-rows", type=int, default=0, help="Optional row limit per TRACE file for testing. 0 = full file.") 

    parser.add_argument("--in-index-only", action="store_true", help="Drop printing CUSIPs that are not in the H0A0 index.") 

    args = parser.parse_args() 

 

    started = time.time() 

    print("=" * 62) 

    print("  trace_digest.py  v2  ---  CLIENT FLOW ENABLED") 

    print("  writes: trace_client_buys / trace_client_sells / trace_flow_label") 

    print("=" * 62) 

    if not os.path.exists(args.trace): 

        print(f"ERROR: TRACE folder not found: {args.trace}", file=sys.stderr) 

        sys.exit(1) 

    os.makedirs(args.out_dir, exist_ok=True) 

 

    print("Loading ICE H0A0 index reference...") 

    ref = read_ice_source(args.hy, "HY") if args.hy and os.path.exists(args.hy) else {} 

    if not ref: 

        print("WARNING: no H0A0 reference loaded; digest will have empty index fields and in_h0a0=0.") 

    else: 

        print(f"Loaded {len(ref):,} H0A0 reference CUSIPs") 

 

    print("Processing TRACE folder...") 

    agg, quality = process_trace_folder(args.trace, args.days, args.sample_rows) 

    rows, in_index = build_trace_digest_rows(agg, ref, in_index_only=args.in_index_only) 

 

    out_path = os.path.join(args.out_dir, "trace_digest.csv") 

    write_trace_digest(out_path, rows) 

 

    print("\nDone.") 

    print(f"  window:           last {args.days} days" 

          + (f" ({quality['trace_window_start']} .. {quality['trace_window_end']})" if quality.get("trace_window_start") else "")) 

    print(f"  TRACE files:      {quality['trace_files_used']} used of {quality['trace_files_available']} in folder") 

    print(f"  prints parsed:    {quality['trace_rows_used']:,} in window" 

          + (f"  ({quality['trace_rows_outside_window']:,} rows skipped as older than window)" if quality.get("trace_rows_outside_window") else "")) 

    print(f"  CUSIPs printing:  {quality['trace_cusips']:,}") 

    if quality.get("trace_side_column"): 

        cb = sum(1 for r in rows if str(r.get("trace_client_buys") or 0) not in ("", "0")) 

        print(f"  client flow:      from column {quality['trace_side_column']!r}; {cb:,} CUSIP(s) with client buys") 

    else: 

        print("  client flow:      NOT AVAILABLE (no dealer-side column matched) <-- see headers above") 

    print(f"  in H0A0 index:    {in_index:,}  ({'index-only output' if args.in_index_only else 'all names kept, in_h0a0 flag set'})") 

    print(f"  digest written:   {out_path}  ({len(rows):,} rows)") 

    print(f"  runtime:          {round(time.time() - started, 2)}s") 

    if not quality["trace_rows_used"]: 

        print("  <-- 0 prints parsed: check the TRACE file format / column names (sample above)") 

 

 

if __name__ == "__main__": 

    try: 

        main() 

    except Exception as e: 

        print(f"ERROR: {e}", file=sys.stderr) 

        sys.exit(1) 
