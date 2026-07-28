#!/usr/bin/env python3
r"""
squeeze_screener.py

Standalone "Short Squeeze Screener" tool.

  Squeeze Score (0-100) = weighted blend of:
    - Crowding          (Days-to-Cover percentile, from markit_dtc.py)
    - Fee momentum      (change in INDICATIVEFEE over --trend-days)
    - Utilization trend (change in UTILISATIONBYQUANTITY over --trend-days)
    - Lendable trend    (% shrink in LENDABLEQUANTITY over --trend-days;
                          shrinking supply scores high)
    - Flow direction    (TRACE client buy/sell read over the same window
                          used for the DTC calc)

This is deliberately a SEPARATE script, not a new pipeline stage. It does
not re-implement TRACE/ICE/Markit parsing -- it imports the already-built,
already-tested pieces:

  from marv_pipeline.py:  process_trace_folder(), read_ice_source(),
                           norm_cusip(), num(), safe_div(), excel_date(),
                           get(), _flow_label(), DEFAULT_TRACE, DEFAULT_HY
  from markit_dtc.py:     collect_markit_files(), read_all_markit_files(),
                           latest_row_per_cusip(), build_rows() (the DTC/
                           crowding calc itself), make_percentile_fn(),
                           MK_* column constants

Both marv_pipeline.py and markit_dtc.py must sit in the same folder as this
script (or pass --pipeline-dir).

Output:
  <out-dir>/squeeze_screener.csv
  <out-dir>/squeeze_screener.json   (what the Neocities leaderboard fetches)

Usage:
  python squeeze_screener.py
  python squeeze_screener.py --markit "N:\toddddumont\markit_data" --trend-days 10 --top-n 20

NOTE ON WEIGHTS: the five component weights are a starting point, not a
back-tested model -- there's no return data behind them yet. They're stored
in meta.weights in the JSON output specifically so they're visible and
adjustable (via --w-crowd etc.) rather than buried. Treat the ranked list as
a screen to investigate, not a signal to trade off directly, until you've
had a chance to eyeball a few weeks of it against what actually happened.

NOTE ON UNIT ASSUMPTIONS: same caveat as markit_dtc.py -- INDICATIVEFEE and
UTILISATIONBYQUANTITY are used as raw, unscaled deltas. A raw fee/util
change is still directionally meaningful (rising vs falling) even before
the unit scale is confirmed, which is why momentum is safe to compute now;
just don't read the *_raw magnitude as a calibrated basis-point number yet.
"""

from __future__ import annotations

import argparse
import csv
import json
import os
import sys
import time
from collections import defaultdict
from datetime import datetime, timedelta
from pathlib import Path
from typing import Dict, List, Optional, Tuple


def _import_local_modules(pipeline_dir: str):
    sys.path.insert(0, pipeline_dir)
    try:
        import marv_pipeline as mp  # type: ignore
    except ImportError as exc:
        print(f"ERROR: could not import marv_pipeline.py from {pipeline_dir!r}: {exc}", file=sys.stderr)
        print("       pass --pipeline-dir pointing at the folder containing both scripts", file=sys.stderr)
        sys.exit(1)
    try:
        import markit_dtc as mdtc  # type: ignore
    except ImportError as exc:
        print(f"ERROR: could not import markit_dtc.py from {pipeline_dir!r}: {exc}", file=sys.stderr)
        print("       squeeze_screener.py builds on top of markit_dtc.py and needs it in the same folder", file=sys.stderr)
        sys.exit(1)
    return mp, mdtc


# --------------------------------------------------------------------------
# Full per-CUSIP history (not collapsed to latest-only, unlike markit_dtc's
# latest_row_per_cusip) -- needed for fee/utilization/lendable momentum.
# --------------------------------------------------------------------------

def build_history_series(raw_rows: List[dict], mdtc, mp) -> Dict[str, List[Tuple]]:
    """cusip -> ascending list of (date, fee, util, lendable_qty, short_qty).
    Deduped by (cusip, date); last occurrence across files wins."""
    tmp: Dict[Tuple[str, str], Tuple] = {}
    for rr in raw_rows:
        row, idx = rr["__row__"], rr["__idx__"]
        cusip = mp.norm_cusip(mp.get(row, idx, mdtc.MK_CUSIP))
        if len(cusip) != 9:
            continue
        d = mp.excel_date(mp.get(row, idx, mdtc.MK_DATE))
        if not d:
            continue
        fee = mp.num(mp.get(row, idx, mdtc.MK_FEE))
        util = mp.num(mp.get(row, idx, mdtc.MK_UTIL_QTY))
        lendable = mp.num(mp.get(row, idx, mdtc.MK_LENDABLE_QTY))
        short_qty = mp.num(mp.get(row, idx, mdtc.MK_SHORT_QTY))
        tmp[(cusip, d)] = (d, fee, util, lendable, short_qty)

    series: Dict[str, List[Tuple]] = defaultdict(list)
    for (cusip, _d), vals in tmp.items():
        series[cusip].append(vals)
    for cusip in series:
        series[cusip].sort(key=lambda t: t[0])
    return series


def _value_on_or_before(points: List[Tuple[str, Optional[float]]], target_date: str) -> Optional[float]:
    """points: ascending (date, value). Returns the value at the latest
    date <= target_date, or None if no such point exists."""
    result = None
    for d, v in points:
        if d <= target_date:
            result = v
        else:
            break
    return result


def trend_change(series: List[Tuple], field_index: int, trend_days: int, mode: str) -> Optional[float]:
    """series: ascending list of (date, fee, util, lendable, short_qty).
    mode: 'diff' (latest - old) or 'pct' ((latest-old)/abs(old))."""
    if len(series) < 2:
        return None
    latest_date = series[-1][0]
    latest_val = series[-1][field_index]
    if latest_val is None:
        return None
    target = (datetime.fromisoformat(latest_date) - timedelta(days=trend_days)).date().isoformat()
    points = [(t[0], t[field_index]) for t in series[:-1]]
    old_val = _value_on_or_before(points, target)
    if old_val is None:
        return None
    if mode == "pct":
        if old_val == 0:
            return None
        return (latest_val - old_val) / abs(old_val)
    return latest_val - old_val


FLOW_SCORE_MAP = {"Client buying": 100.0, "Balanced": 50.0, "Client selling": 0.0}

DEFAULT_WEIGHTS = {
    "crowd": 0.30,
    "fee_momentum": 0.25,
    "util_trend": 0.20,
    "lendable_trend": 0.15,
    "flow": 0.10,
}

SQ_COLS = [
    "rank", "cusip", "ticker", "rating_bucket", "markit_date",
    "squeeze_score", "components_used", "insufficient_data_reason",
    "crowd_score", "dtc_pctile_h0a0", "days_to_cover",
    "fee_momentum_score", "fee_chg_raw", "fee_latest",
    "util_trend_score", "util_chg_raw", "util_latest",
    "lendable_trend_score", "lendable_chg_pct", "lendable_latest",
    "flow_score", "trace_flow_label",
    "dcbs_latest", "indicative_fee_note",
    "crowding_flag",
]


def build_squeeze_rows(dtc_rows: List[dict], history: Dict[str, List[Tuple]],
                        agg: dict, mp, trend_days: int, weights: dict) -> List[dict]:

    # ---- pass 1: raw components per in-index cusip ----
    raw: List[dict] = []
    for r in dtc_rows:
        if not r["in_h0a0"]:
            continue
        cusip = r["cusip"]
        series = history.get(cusip, [])

        fee_chg = trend_change(series, 1, trend_days, "diff")
        util_chg = trend_change(series, 2, trend_days, "diff")
        lendable_chg_pct = trend_change(series, 3, trend_days, "pct")

        fee_latest = series[-1][1] if series else None
        util_latest = series[-1][2] if series else None
        lendable_latest = series[-1][3] if series else None

        a = agg.get(cusip)
        flow_label = mp._flow_label(a["cbuy_vol"], a["csell_vol"], a["side_prints"]) if a else ""
        flow_score = FLOW_SCORE_MAP.get(flow_label)

        raw.append({
            "cusip": cusip,
            "ticker": r["ticker"],
            "rating_bucket": r["rating_bucket"],
            "markit_date": r["markit_date"],
            "dtc_pctile_h0a0": r["dtc_pctile_h0a0"],
            "days_to_cover": r["days_to_cover"],
            "crowding_flag": r["crowding_flag"],
            "dcbs_latest": r["dcbs"],
            "fee_chg_raw": round(fee_chg, 2) if fee_chg is not None else None,
            "fee_latest": fee_latest,
            "util_chg_raw": round(util_chg, 2) if util_chg is not None else None,
            "util_latest": util_latest,
            "lendable_chg_pct": round(lendable_chg_pct, 4) if lendable_chg_pct is not None else None,
            "lendable_latest": lendable_latest,
            "trace_flow_label": flow_label,
            "flow_score": flow_score,
        })

    # ---- pass 2: percentile-rank the momentum components across the universe ----
    fee_pct_fn = markit_dtc_percentile([r["fee_chg_raw"] for r in raw])
    util_pct_fn = markit_dtc_percentile([r["util_chg_raw"] for r in raw])
    # shrinking lendable is bullish -> rank on the negated value
    neg_lendable_pct_fn = markit_dtc_percentile(
        [-r["lendable_chg_pct"] if r["lendable_chg_pct"] is not None else None for r in raw]
    )

    rows = []
    for r in raw:
        crowd_score = r["dtc_pctile_h0a0"]
        fee_mom_score = fee_pct_fn(r["fee_chg_raw"])
        util_trend_score = util_pct_fn(r["util_chg_raw"])
        lendable_trend_score = neg_lendable_pct_fn(
            -r["lendable_chg_pct"] if r["lendable_chg_pct"] is not None else None
        )
        flow_score = r["flow_score"]

        subscores = {
            "crowd": crowd_score,
            "fee_momentum": fee_mom_score,
            "util_trend": util_trend_score,
            "lendable_trend": lendable_trend_score,
            "flow": flow_score,
        }
        available = {k: v for k, v in subscores.items() if v is not None}

        insufficient_reason = ""
        squeeze_score = None
        if crowd_score is None:
            insufficient_reason = "no TRACE-matched days-to-cover (excluded from ranking)"
        elif not available:
            insufficient_reason = "no components available"
        else:
            total_w = sum(weights[k] for k in available)
            squeeze_score = round(sum(weights[k] * available[k] for k in available) / total_w, 1)

        rows.append({
            "rank": None,
            "cusip": r["cusip"],
            "ticker": r["ticker"],
            "rating_bucket": r["rating_bucket"],
            "markit_date": r["markit_date"],
            "squeeze_score": squeeze_score,
            "components_used": ",".join(sorted(available.keys())),
            "insufficient_data_reason": insufficient_reason,
            "crowd_score": crowd_score,
            "dtc_pctile_h0a0": r["dtc_pctile_h0a0"],
            "days_to_cover": r["days_to_cover"],
            "fee_momentum_score": fee_mom_score,
            "fee_chg_raw": r["fee_chg_raw"],
            "fee_latest": r["fee_latest"],
            "util_trend_score": util_trend_score,
            "util_chg_raw": r["util_chg_raw"],
            "util_latest": r["util_latest"],
            "lendable_trend_score": lendable_trend_score,
            "lendable_chg_pct": r["lendable_chg_pct"],
            "lendable_latest": r["lendable_latest"],
            "flow_score": flow_score,
            "trace_flow_label": r["trace_flow_label"],
            "dcbs_latest": r["dcbs_latest"],
            "indicative_fee_note": "raw/unscaled -- direction only, see header note",
            "crowding_flag": r["crowding_flag"],
        })

    rows.sort(key=lambda x: (x["squeeze_score"] is None, -(x["squeeze_score"] or 0), x["cusip"]))
    rank = 0
    for r in rows:
        if r["squeeze_score"] is not None:
            rank += 1
            r["rank"] = rank
    return rows


def markit_dtc_percentile(values: List[Optional[float]]):
    """Standalone copy of markit_dtc.make_percentile_fn's bisect-rank logic
    so this module doesn't need to reach into markit_dtc's internals for a
    one-line helper; behavior is identical."""
    import bisect
    s = sorted(v for v in values if v is not None)
    n = len(s) or 1
    def pct(v):
        if v is None:
            return None
        return round(bisect.bisect_right(s, v) / n * 100.0, 1)
    return pct


def write_outputs(out_dir: Path, rows: List[dict], meta: dict) -> None:
    out_dir.mkdir(parents=True, exist_ok=True)

    csv_path = out_dir / "squeeze_screener.csv"
    with open(csv_path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(SQ_COLS)
        for r in rows:
            w.writerow([r.get(c, "") for c in SQ_COLS])

    json_path = out_dir / "squeeze_screener.json"
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump({"meta": meta, "rows": rows}, f, indent=2, default=str)

    print(f"  wrote: {csv_path}")
    print(f"  wrote: {json_path}")


def main(argv=None) -> int:
    p = argparse.ArgumentParser(description="Short squeeze screener (standalone, builds on markit_dtc.py).")
    p.add_argument("--markit", default=None, help="Path to Markit file/folder. Default: markit_dtc's DEFAULT_MARKIT.")
    p.add_argument("--pipeline-dir", default=str(Path(__file__).resolve().parent),
                   help="Folder containing marv_pipeline.py and markit_dtc.py.")
    p.add_argument("--trace", default=None, help="TRACE folder. Default: marv_pipeline's DEFAULT_TRACE.")
    p.add_argument("--hy", default=None, help="ICE H0A0 reference path. Default: marv_pipeline's DEFAULT_HY.")
    p.add_argument("--days", type=int, default=20, help="TRACE ADV window in calendar days (default 20).")
    p.add_argument("--trend-days", type=int, default=10, help="Lookback window for fee/util/lendable momentum (default 10 calendar days).")
    p.add_argument("--trace-sample-rows", type=int, default=0, help="Row cap per TRACE file, for testing.")
    p.add_argument("--top-n", type=int, default=20, help="Leaderboard size the Neocities page should default to showing (stored in meta; full list is still written).")
    p.add_argument("--out-dir", default="squeeze_screener_output", help="Output folder.")
    p.add_argument("--w-crowd", type=float, default=DEFAULT_WEIGHTS["crowd"])
    p.add_argument("--w-fee", type=float, default=DEFAULT_WEIGHTS["fee_momentum"])
    p.add_argument("--w-util", type=float, default=DEFAULT_WEIGHTS["util_trend"])
    p.add_argument("--w-lendable", type=float, default=DEFAULT_WEIGHTS["lendable_trend"])
    p.add_argument("--w-flow", type=float, default=DEFAULT_WEIGHTS["flow"])
    args = p.parse_args(argv)

    mp, mdtc = _import_local_modules(args.pipeline_dir)
    markit_path = args.markit or mdtc.DEFAULT_MARKIT
    trace_dir = args.trace or mp.DEFAULT_TRACE
    hy_path = args.hy or mp.DEFAULT_HY
    weights = {
        "crowd": args.w_crowd, "fee_momentum": args.w_fee, "util_trend": args.w_util,
        "lendable_trend": args.w_lendable, "flow": args.w_flow,
    }

    started = time.time()
    print("=" * 62)
    print("  squeeze_screener.py  --  Short Squeeze Screener")
    print("=" * 62)
    print(f"  weights: {weights}")

    if not os.path.exists(markit_path):
        print(f"ERROR: Markit path not found: {markit_path}", file=sys.stderr)
        return 2
    if not os.path.exists(trace_dir):
        print(f"ERROR: TRACE folder not found: {trace_dir}", file=sys.stderr)
        return 2

    print(f"Reading Markit data: {markit_path}")
    raw_rows = mdtc.read_all_markit_files(markit_path, mp)
    markit_latest = mdtc.latest_row_per_cusip(raw_rows, mp)
    history = build_history_series(raw_rows, mdtc, mp)
    dates_seen = sorted({t[0] for series in history.values() for t in series})
    print(f"  {len(raw_rows):,} row(s) total, {len(markit_latest):,} CUSIP(s), "
          f"history spans {dates_seen[0] if dates_seen else '?'} to {dates_seen[-1] if dates_seen else '?'} "
          f"({len(dates_seen)} distinct date(s))")

    print("Loading ICE H0A0 index reference...")
    ref = mp.read_ice_source(hy_path, "HY") if hy_path and os.path.exists(hy_path) else {}
    print(f"  {len(ref):,} H0A0 reference CUSIP(s) loaded" if ref else "  WARNING: no H0A0 reference loaded")

    print(f"Processing TRACE folder ({trace_dir}), last {args.days} day(s)...")
    agg, quality = mp.process_trace_folder(trace_dir, args.days, args.trace_sample_rows)
    print(f"  {quality['trace_files_used']} file(s) used, {quality['trace_rows_used']:,} print(s) in window")

    window_days = args.days
    if quality.get("trace_window_start") and quality.get("trace_window_end"):
        try:
            d0 = datetime.fromisoformat(quality["trace_window_start"])
            d1 = datetime.fromisoformat(quality["trace_window_end"])
            window_days = max(1, (d1 - d0).days + 1)
        except Exception:
            pass

    dtc_rows = mdtc.build_rows(markit_latest, agg, ref, mp, window_days)

    print(f"Computing squeeze components ({args.trend_days}-day momentum window)...")
    rows = build_squeeze_rows(dtc_rows, history, agg, mp, args.trend_days, weights)

    ranked = [r for r in rows if r["squeeze_score"] is not None]
    excluded = len(rows) - len(ranked)
    top = ranked[: args.top_n]

    if dates_seen and len(dates_seen) < 3:
        print(f"  NOTE: only {len(dates_seen)} distinct Markit date(s) available -- momentum components will mostly "
              f"read as 'insufficient history' until more weekly files build up. Crowding-only ranking still works.")

    meta = {
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "markit_source": os.path.basename(markit_path.rstrip("/\\")),
        "markit_history_dates": dates_seen,
        "trend_days": args.trend_days,
        "trace_window_days": window_days,
        "trace_window_start": quality.get("trace_window_start"),
        "trace_window_end": quality.get("trace_window_end"),
        "top_n": args.top_n,
        "ranked_count": len(ranked),
        "excluded_count": excluded,
        "weights": weights,
        "note": "Squeeze Score weights are a starting-point heuristic, not back-tested. "
                "Markit data typically lags TRACE by 2-3 days.",
    }

    out_dir = Path(args.out_dir)
    write_outputs(out_dir, rows, meta)

    print()
    print(f"  Ranked:    {len(ranked):,}")
    print(f"  Excluded:  {excluded:,} (no TRACE-matched days-to-cover)")
    print(f"  Top {min(args.top_n, len(top))}:")
    for r in top[:10]:
        print(f"    #{r['rank']:<3} {r['ticker'] or r['cusip']:<12} score={r['squeeze_score']:<6} "
              f"components=[{r['components_used']}]  flag={r['crowding_flag'] or '-'}")
    print(f"  runtime: {round(time.time() - started, 2)}s")
    if not ranked:
        print("  <-- 0 bonds ranked: check CUSIP overlap between Markit and TRACE (same as markit_dtc.py)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
