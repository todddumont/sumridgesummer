#!/usr/bin/env python3
"""
H0A0 Historical Relative Value Engine v2.0


Standalone analytics script for ICE/BAML H0A0 historical constituent files.


Purpose:
    Identify rich/cheap bonds and issuers using historical OAS, peer residuals,
    historical peer-adjusted residuals, Level-4 cohort valuation context,
    current shrunk peer residuals, issuer context, and own-OAS history.


Default data path:
    P:\\jmorris\\ICE H0A0 Historical Index Data


Outputs:
    output/bond_rv_latest.csv
    output/issuer_rv_latest.csv
    output/issuer_switch_candidates.csv
    output/bond_rv_other_latest.csv
    output/distressed_event_latest.csv
    output/data_quality_report.txt
    output/h0a0_relative_value_report.xlsx


Notes:
    - Does not upload or move raw ICE/BAML data.
    - Uses pandas/numpy/openpyxl only.
    - Saves a normalized-history cache after parsing so future runs are faster.
"""


from __future__ import annotations


import argparse
import csv
import json
import math
import os
import re
import sys
import time
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Tuple


import numpy as np
import pandas as pd
pd.set_option("future.no_silent_downcasting", True)


DEFAULT_HISTORY_DIR = r"P:\jmorris\ICE H0A0 Historical Index Data"
DEFAULT_OUTPUT_DIR = "output"
CACHE_DIR_NAME = "cache"
CACHE_FILE_NAME = "h0a0_normalized_history.pkl"
MANIFEST_FILE_NAME = "h0a0_cache_manifest.json"
DEFAULT_HISTORY_LOOKBACK_MONTHS = 12


SUPPORTED_EXTENSIONS = {".csv", ".txt", ".xlsx", ".xls", ".xlsm"}


# Canonical field aliases. The ICE/BAML file uses long names and occasional spacing variations.
COLUMN_ALIASES: Dict[str, List[str]] = {
    "cusip": ["Cusip", "CUSIP"],
    "ticker": ["Ticker", "Issuer Ticker"],
    "isin": ["ISIN number", "ISIN", "Isin"],
    "description": ["Description", "Security Description"],
    "rating": ["Rating", "Composite Rating"],
    "sector_l2": ["ML Industry Lvl 2", "Industry Level 2"],
    "sector_l3": ["ML Industry Lvl 3", "Industry Level 3"],
    "sector_l4": ["ML Industry Lvl 4", "Industry Level 4"],
    "iso_country": ["ISO Country", "Country"],
    "geo_region": ["Geo Region", "Region"],
    "price": ["Price", "Current Price"],
    "prev_price": ["Previous Price", "Prev Price"],
    "oas": ["OAS", "Option Adjusted Spread"],
    "spread_to_worst": ["Spread To Worst", "Spread to Worst", "STW"],
    "ytw": ["Yld to Worst", "Yield to Worst", "YTW"],
    "ytm": ["Yld to Mat", "Yield to Maturity", "YTM"],
    "years_to_worst": ["Yrs To Worst", "Years To Worst", "Years to Worst"],
    "maturity_wal": ["Maturity / WAL", "Maturity/WAL"],
    "maturity_date": ["Maturity Date", "Mty Date"],
    "coupon": ["Par Wtd Coupon", "Coupon"],
    "spread_duration": ["Effective Spread Duration", "Spread Duration", "Eff Spread Duration"],
    "duration_to_worst": ["Dur To Worst", "Duration To Worst"],
    "effective_duration": ["Effective Duration"],
    "index_weight": ["Mkt % Index Wght", "Market % Index Weight", "Index Weight"],
    "face_value": ["Face Value", "Par Amount"],
    "market_value": ["Full Market Value", "Market Value"],
    "mty_type": ["Mty Type", "Maturity Type"],
    "type": ["Type"],
    "oas_1d_chg": ["OAS 1-Day Change", "OAS 1 Day Change"],
    "oas_1w_chg": ["OAS 1-Week Change", "OAS 1 Week Change"],
    "oas_1m_chg": ["OAS 1-Mo Change", "OAS 1-Mth Change", "OAS 1 Month Change"],
    "excess_return_1d": ["Excess Rtn % 1-day", "Excess Return % 1-day"],
    "excess_return_1w": ["Excess Rtn % 1-week", "Excess Return % 1-week"],
    "excess_return_1m": ["Excess Rtn % 1-month", "Excess Return % 1-month"],
}


MODEL_VERSION = "v2.1 Full-Rating Scored Universe + Internal-OAS Weighted RV Score"
CORE_RATINGS_EXACT = {"BB1", "BB2", "BB3", "B1", "B2", "B3"}
CORE_RATING_BUCKETS = {"B", "BB"}

# v2.1: the scored universe is no longer restricted to BB1-B3. Crossover/split-rated
# entrants (BBB notches) and performing CCC credits are scored and written to the core
# output. Peer groups and cohorts are keyed on rating_bucket, so these names form their
# own comparison sets rather than contaminating BB/B medians.
SCORED_RATINGS_EXACT = {
    "BBB1", "BBB2", "BBB3",
    "BB1", "BB2", "BB3",
    "B1", "B2", "B3",
    "CCC1", "CCC2", "CCC3",
}
SCORED_RATING_BUCKETS = {"BBB", "BB", "B", "CCC_OR_BELOW"}

# Ratings that are never scored: actual/near default.
NEVER_SCORED_RATINGS = {"D", "C", "CC"}

# Universe mode. "all_rated" (default) scores every performing rated bond.
# "bb_b_only" restores the pre-v2.1 BB1-B3 behaviour.
CORE_UNIVERSE_MODE = "all_rated"

# Hard tradability gates for the scored universe under "all_rated".
# These exist to drop recovery-value/broken rows, not to enforce a rating band.
ALL_RATED_MIN_PRICE = 50.0
ALL_RATED_MAX_OAS = 2000.0

# Labelling thresholds. These tag a bond as stressed in the output; they no longer
# remove it from the scored book under "all_rated".
STRESS_LABEL_PRICE = 80.0
STRESS_LABEL_OAS = 800.0


def scored_ratings_exact() -> set:
    return CORE_RATINGS_EXACT if CORE_UNIVERSE_MODE == "bb_b_only" else SCORED_RATINGS_EXACT


def scored_rating_buckets() -> set:
    return CORE_RATING_BUCKETS if CORE_UNIVERSE_MODE == "bb_b_only" else SCORED_RATING_BUCKETS


def universe_min_price() -> float:
    return 80.0 if CORE_UNIVERSE_MODE == "bb_b_only" else ALL_RATED_MIN_PRICE


def universe_max_oas() -> float:
    return 800.0 if CORE_UNIVERSE_MODE == "bb_b_only" else ALL_RATED_MAX_OAS
TARGET_SECTOR_GROUPS = {"Retail", "Healthcare"}
TARGET_SECTOR_LABEL = "Full HY"


NUMERIC_COLUMNS = [
    "price",
    "prev_price",
    "oas",
    "spread_to_worst",
    "ytw",
    "ytm",
    "years_to_worst",
    "maturity_wal",
    "coupon",
    "spread_duration",
    "duration_to_worst",
    "effective_duration",
    "index_weight",
    "face_value",
    "market_value",
    "oas_1d_chg",
    "oas_1w_chg",
    "oas_1m_chg",
    "excess_return_1d",
    "excess_return_1w",
    "excess_return_1m",
]


OUTPUT_BOND_COLUMNS = [
    "as_of_date",
    "ticker",
    "description",
    "cusip",
    "isin",
    "bond_key",
    "bond_label",
    "coupon",
    "maturity_date",
    "rating",
    "rating_raw",
    "rating_bucket",
    "core_universe_status",
    "stress_label",
    "sector_l3",
    "sector_l4",
    "target_sector_group",
    "review_bucket",
    "review_priority_score",
    "lender_stress_status",
    "price",
    "oas",
    "ytw",
    "years_to_worst",
    "maturity_bucket",
    "spread_duration",
    "index_weight",
    "oas_pct_1y",
    "oas_pct_2y",
    "oas_pct_full",
    "oas_obs_1y",
    "oas_vs_1y_median",
    "oas_vs_1y_p75",
    "residual_oas_pct_1y",
    "residual_oas_pct_2y",
    "residual_oas_pct_full",
    "residual_oas_obs_1y",
    "residual_oas_vs_1y_median",
    "residual_oas_vs_1y_p75",
    "issuer_oas_pct_1y",
    "issuer_oas_pct_2y",
    "issuer_oas_pct_full",
    "issuer_residual_pct_1y",
    "issuer_residual_pct_2y",
    "issuer_residual_pct_full",
    "issuer_oas_vs_1y_median",
    "issuer_oas_vs_1y_p75",
    "peer_group_used",
    "peer_count",
    "peer_group_quality",
    "peer_median_stability_bp",
    "peer_median_stability_label",
    "peer_median_oas",
    "fair_oas_shrunk",
    "peer_oas_residual",
    "shrunk_peer_residual",
    "peer_oas_percentile_today",
    "cohort_group_used",
    "cohort_current_count",
    "cohort_median_oas",
    "cohort_oas_pct_1y",
    "cohort_oas_pct_full",
    "cohort_context_label",
    "relative_vs_absolute_label",
    "issuer_curve_fair_oas",
    "issuer_curve_residual",
    "issuer_curve_method",
    "issuer_curve_confidence",
    "issuer_curve_reason",
    "issuer_curve_r2",
    "issuer_curve_x_span",
    "issuer_curve_bond_count",
    "issuer_curve_model_bond_count",
    "curve_structure_bucket",
    "historical_peer_adjusted_residual_score",
    "cohort_context_score",
    "current_shrunk_peer_residual_score",
    "issuer_context_score",
    "own_oas_history_score",
    "momentum_score",
    "issuer_curve_overlay_points",
    "rv_score_base",
    "rv_score_method",
    "history_confidence",
    "model_support",
    "why_signal_may_be_wrong",
    "rv_score",
    "rv_signal",
    "flags",
    "rv_note",
]


ISSUER_OUTPUT_COLUMNS = [
    "as_of_date",
    "ticker",
    "issuer_bond_count",
    "issuer_index_weight",
    "issuer_median_oas",
    "issuer_weighted_oas",
    "issuer_median_price",
    "issuer_median_ytw",
    "issuer_rating_bucket",
    "issuer_core_universe_status",
    "issuer_sector_l3",
    "issuer_target_sector_group",
    "issuer_review_bucket",
    "issuer_review_priority_score",
    "issuer_peer_group_used",
    "issuer_peer_count",
    "issuer_peer_median_oas",
    "issuer_peer_oas_residual",
    "issuer_oas_pct_1y",
    "issuer_oas_pct_2y",
    "issuer_oas_pct_full",
    "issuer_residual_pct_1y",
    "issuer_residual_pct_2y",
    "issuer_residual_pct_full",
    "issuer_oas_vs_1y_median",
    "issuer_oas_vs_1y_p75",
    "issuer_distressed_event",
    "issuer_flags",
    "issuer_confidence",
    "issuer_score",
    "issuer_signal",
    "issuer_relative_vs_absolute_label",
    "cheap_bond_cusips",
    "rich_bond_cusips",
    "issuer_note",
]




def log(message: str) -> None:
    now = datetime.now().strftime("%H:%M:%S")
    print(f"[{now}] {message}", flush=True)




def clean_col_name(value: object) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value).strip())




def parse_date_from_filename(path: Path) -> Optional[pd.Timestamp]:
    match = re.search(r"(20\d{6})", path.name)
    if not match:
        return None
    try:
        return pd.to_datetime(match.group(1), format="%Y%m%d")
    except Exception:
        return None




def find_header_row_csv(path: Path, max_rows: int = 20) -> int:
    with path.open("r", encoding="utf-8-sig", errors="replace", newline="") as handle:
        reader = csv.reader(handle)


        for idx, row in enumerate(reader):


            if idx >= max_rows:
                break


            normalized = [clean_col_name(x).lower() for x in row]


            # Look for the real ICE table header instead of exact names.
            has_cusip = any("cusip" in c for c in normalized)
            has_ticker = any("ticker" in c for c in normalized)
            has_desc = any("description" in c for c in normalized)
            has_coupon = any("coupon" in c for c in normalized)


            # A real ICE header will always have CUSIP plus either
            # ticker/description and the coupon column.
            if has_cusip and has_coupon and (has_ticker or has_desc):
                return idx


    return 0




def read_raw_file(path: Path) -> pd.DataFrame:
    ext = path.suffix.lower()
    if ext in {".csv", ".txt"}:
        header_row = find_header_row_csv(path)
        return pd.read_csv(path, skiprows=header_row, dtype=str, low_memory=False, encoding="utf-8-sig")
    if ext in {".xlsx", ".xls", ".xlsm"}:
        # Read a small preview to find the header row.
        preview = pd.read_excel(path, header=None, nrows=20, dtype=str)
        header_row = 0
        for idx, row in preview.iterrows():
            normalized = {clean_col_name(x).lower() for x in row.tolist()}
            if "cusip" in normalized and "ticker" in normalized:
                header_row = idx
                break
        return pd.read_excel(path, skiprows=header_row, dtype=str)
    raise ValueError(f"Unsupported file extension: {path.suffix}")




def pick_column(raw_columns: Sequence[str], aliases: Sequence[str]) -> Optional[str]:
    clean_to_raw = {clean_col_name(col).lower(): col for col in raw_columns}
    for alias in aliases:
        hit = clean_to_raw.get(clean_col_name(alias).lower())
        if hit is not None:
            return hit
    return None




def normalize_one_file(path: Path) -> Tuple[pd.DataFrame, Dict[str, object]]:
    as_of_date = parse_date_from_filename(path)
    raw = read_raw_file(path)
    raw.columns = [clean_col_name(c) for c in raw.columns]


    data: Dict[str, pd.Series] = {}
    for canonical, aliases in COLUMN_ALIASES.items():
        raw_col = pick_column(raw.columns, aliases)
        if raw_col is None:
            data[canonical] = pd.Series([np.nan] * len(raw), index=raw.index)
        else:
            data[canonical] = raw[raw_col]


    df = pd.DataFrame(data)
    df.insert(0, "as_of_date", as_of_date)
    df.insert(1, "source_file", path.name)


    for col in NUMERIC_COLUMNS:
        if col in df.columns:
            df[col] = (
                df[col]
                .astype(str)
                .str.replace(",", "", regex=False)
                .str.replace("%", "", regex=False)
                .replace({"": np.nan, "nan": np.nan, "None": np.nan, "--": np.nan, "N/A": np.nan})
            )
            df[col] = pd.to_numeric(df[col], errors="coerce")


    for col in ["ticker", "rating", "sector_l2", "sector_l3", "sector_l4", "isin", "cusip", "description"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip().replace({"": np.nan, "nan": np.nan, "None": np.nan})


    if "maturity_date" in df.columns:
        df["maturity_date"] = pd.to_datetime(df["maturity_date"], errors="coerce")


    # Remove rows with no usable security identity and no market data.
    identity_present = df[["isin", "cusip", "ticker", "description"]].notna().any(axis=1)
    market_present = df[["price", "oas", "ytw"]].notna().any(axis=1)
    df = df.loc[identity_present & market_present].copy()


    info = {
        "file": path.name,
        "date": str(as_of_date.date()) if as_of_date is not None and not pd.isna(as_of_date) else None,
        "rows": int(len(df)),
    }
    return df, info




def list_history_files(history_dir: Path) -> List[Path]:
    files: List[Path] = []
    for path in history_dir.iterdir():
        if path.is_file() and path.suffix.lower() in SUPPORTED_EXTENSIONS:
            # Do not accidentally parse prior outputs.
            lower_name = path.name.lower()
            if lower_name.startswith("bond_rv_") or lower_name.startswith("issuer_rv_"):
                continue
            if "relative_value_report" in lower_name or "rv_output" in lower_name:
                continue
            files.append(path)
    files.sort(key=lambda p: (parse_date_from_filename(p) or pd.Timestamp.min, p.name))
    return files






def filter_history_files_to_lookback(
    files: Sequence[Path],
    requested_as_of_date: Optional[str] = None,
    history_lookback_months: int = DEFAULT_HISTORY_LOOKBACK_MONTHS,
) -> Tuple[List[Path], Optional[pd.Timestamp], Optional[pd.Timestamp]]:
    """Return only ICE history files inside the requested lookback window.


    The RV model uses only the last 12 months of ICE files by default.
    The anchor date is the requested as-of date when supplied; otherwise it is
    the latest date found in the ICE filenames. Files outside the window are not
    parsed, cached, or used in score percentiles/residual histories.


    Set --history-lookback-months 0 to deliberately use all files.
    """
    all_files = list(files)
    if history_lookback_months is None or int(history_lookback_months) <= 0:
        log(f"History lookback disabled: using all {len(all_files):,} supported file(s).")
        return all_files, None, None


    dated: List[Tuple[pd.Timestamp, Path]] = []
    undated: List[Path] = []
    for path in all_files:
        dt = parse_date_from_filename(path)
        if dt is None or pd.isna(dt):
            undated.append(path)
        else:
            dated.append((pd.Timestamp(dt), path))


    if not dated:
        log("WARNING: no parseable dates found in ICE filenames; using all supported history files.")
        return all_files, None, None


    if requested_as_of_date:
        anchor_date = pd.Timestamp(pd.to_datetime(requested_as_of_date)).normalize()
    else:
        anchor_date = max(dt for dt, _ in dated).normalize()


    cutoff_date = (anchor_date - pd.DateOffset(months=int(history_lookback_months))).normalize()
    filtered = [path for dt, path in dated if cutoff_date <= dt.normalize() <= anchor_date]
    filtered.sort(key=lambda p: (parse_date_from_filename(p) or pd.Timestamp.min, p.name))


    if not filtered:
        raise FileNotFoundError(
            f"No ICE history files found from {cutoff_date.date()} through {anchor_date.date()} "
            f"in the {history_lookback_months}-month lookback window."
        )


    log(
        f"History lookback: using {len(filtered):,} dated ICE file(s) from "
        f"{cutoff_date.date()} through {anchor_date.date()} "
        f"({history_lookback_months} months)."
    )
    if undated:
        log(f"History lookback: ignored {len(undated):,} undated supported file(s).")
    return filtered, anchor_date, cutoff_date


def build_manifest(
    files: Sequence[Path],
    history_lookback_months: int = DEFAULT_HISTORY_LOOKBACK_MONTHS,
    anchor_date: Optional[pd.Timestamp] = None,
    cutoff_date: Optional[pd.Timestamp] = None,
) -> Dict[str, object]:
    return {
        "version": 4,
        "history_lookback_months": int(history_lookback_months or 0),
        "anchor_date": str(pd.Timestamp(anchor_date).date()) if anchor_date is not None and not pd.isna(anchor_date) else None,
        "cutoff_date": str(pd.Timestamp(cutoff_date).date()) if cutoff_date is not None and not pd.isna(cutoff_date) else None,
        "files": [
            {"name": f.name, "size": f.stat().st_size, "mtime": int(f.stat().st_mtime)}
            for f in files
        ],
    }




def manifest_matches(existing: Dict[str, object], current: Dict[str, object]) -> bool:
    return existing == current




def load_or_parse_history(
    history_dir: Path,
    output_dir: Path,
    rebuild_cache: bool = False,
    max_files: Optional[int] = None,
    as_of_date: Optional[str] = None,
    history_lookback_months: int = DEFAULT_HISTORY_LOOKBACK_MONTHS,
) -> Tuple[pd.DataFrame, List[Dict[str, object]], bool]:
    files = list_history_files(history_dir)
    files, anchor_date, cutoff_date = filter_history_files_to_lookback(
        files,
        requested_as_of_date=as_of_date,
        history_lookback_months=history_lookback_months,
    )
    if max_files is not None and max_files > 0:
        files = files[-max_files:]
        log(f"Test mode: using latest {len(files)} file(s).")


    if not files:
        raise FileNotFoundError(f"No supported history files found in {history_dir}")


    cache_dir = output_dir / CACHE_DIR_NAME
    cache_dir.mkdir(parents=True, exist_ok=True)
    cache_file = cache_dir / CACHE_FILE_NAME
    manifest_file = cache_dir / MANIFEST_FILE_NAME
    current_manifest = build_manifest(
        files,
        history_lookback_months=history_lookback_months,
        anchor_date=anchor_date,
        cutoff_date=cutoff_date,
    )


    if not rebuild_cache and max_files is None and cache_file.exists() and manifest_file.exists():
        try:
            existing_manifest = json.loads(manifest_file.read_text(encoding="utf-8"))
            if manifest_matches(existing_manifest, current_manifest):
                log(f"Loading normalized history cache: {cache_file}")
                history = pd.read_pickle(cache_file)
                history = apply_current_model_rules(history)
                log(f"Loaded {len(history):,} normalized row(s) from cache and applied {MODEL_VERSION} universe rules.")
                return history, [], True
        except Exception as exc:
            log(f"Cache load skipped: {exc}")


    log(f"Parsing {len(files):,} historical file(s) from: {history_dir}")
    frames: List[pd.DataFrame] = []
    parse_report: List[Dict[str, object]] = []
    start = time.time()


    for idx, path in enumerate(files, start=1):
        try:
            frame, info = normalize_one_file(path)
            frames.append(frame)
            parse_report.append(info)
        except Exception as exc:
            parse_report.append({"file": path.name, "date": None, "rows": 0, "error": str(exc)})
            log(f"WARNING: failed to parse {path.name}: {exc}")
            continue


        if idx == 1 or idx % 25 == 0 or idx == len(files):
            latest_date = max([r.get("date") for r in parse_report if r.get("date")], default="unknown")
            log(f"Parsed {idx:,}/{len(files):,} files; latest file date parsed {latest_date} ({path.name})")


    if not frames:
        raise RuntimeError("No files parsed successfully.")


    history = pd.concat(frames, ignore_index=True)
    history = finalize_history(history)
    log(f"Normalized history has {len(history):,} row(s). Parse time: {(time.time() - start) / 60:.1f} minutes.")


    if max_files is None:
        log(f"Saving normalized history cache: {cache_file}")
        history.to_pickle(cache_file)
        manifest_file.write_text(json.dumps(current_manifest, indent=2), encoding="utf-8")


    return history, parse_report, False






def target_sector_group(df: pd.DataFrame) -> pd.Series:
    """Classify the project target sector group.


    This helper keeps a lightweight Core/Other sector grouping,
    but v1.6 is a full high-yield split monitor. Classification is driven by ICE/BAML
    industry fields, not broad issuer-name guessing. Description fallback is
    only used when industry fields are missing or unknown.
    """
    sector_l3 = df.get("sector_l3", pd.Series("", index=df.index)).fillna("").astype(str).str.upper()
    sector_l4 = df.get("sector_l4", pd.Series("", index=df.index)).fillna("").astype(str).str.upper()
    sector_l2 = df.get("sector_l2", pd.Series("", index=df.index)).fillna("").astype(str).str.upper()
    sector_text = sector_l2 + " " + sector_l3 + " " + sector_l4
    desc = df.get("description", pd.Series("", index=df.index)).fillna("").astype(str).str.upper()


    retail_mask = sector_text.str.contains(r"\bRETAIL\b", regex=True, na=False)
    healthcare_mask = sector_text.str.contains(r"HEALTH\s*CARE|HEALTHCARE|\bHEALTH\b", regex=True, na=False)


    unknown_sector = sector_text.str.contains(r"UNKNOWN|NAN|NONE", regex=True, na=False) | sector_text.str.strip().eq("")
    # Very limited fallback for rows where sector fields are unavailable.
    retail_mask = retail_mask | (unknown_sector & desc.str.contains(r"\bRETAIL\b", regex=True, na=False))
    healthcare_mask = healthcare_mask | (unknown_sector & desc.str.contains(r"HEALTH\s*CARE|HEALTHCARE|HOSPITAL|MEDICAL", regex=True, na=False))


    out = pd.Series("Other", index=df.index, dtype=object)
    out = out.mask(retail_mask, "Retail")
    out = out.mask(healthcare_mask, "Healthcare")
    return out




def is_target_sector(df: pd.DataFrame) -> pd.Series:
    group = df.get("target_sector_group")
    if group is None:
        group = target_sector_group(df)
    return group.astype(str).isin(TARGET_SECTOR_GROUPS)




def apply_current_model_rules(df: pd.DataFrame) -> pd.DataFrame:
    """Reapply v1.6 universe rules to cached normalized history.


    Existing caches may have been created by earlier model versions. This migration
    keeps the cache useful while ensuring the current full high-yield split monitor
    rules and flags are applied before scoring.
    """
    df = df.copy()
    if "as_of_date" in df.columns:
        df["as_of_date"] = pd.to_datetime(df["as_of_date"], errors="coerce")
    if "ticker" in df.columns:
        df["ticker"] = df["ticker"].fillna("UNKNOWN").astype(str).str.upper().str.strip()
    if "rating" in df.columns:
        if "rating_raw" not in df.columns:
            df["rating_raw"] = df["rating"].fillna("NR").astype(str).str.upper().str.strip()
        df["rating"] = normalize_rating_notation(df["rating_raw"])
    if "sector_l2" in df.columns:
        df["sector_l2"] = df["sector_l2"].fillna("UNKNOWN").astype(str).str.strip()
    if "sector_l3" in df.columns:
        df["sector_l3"] = df["sector_l3"].fillna(df.get("sector_l2", "UNKNOWN")).fillna("UNKNOWN").astype(str).str.strip()
    if "sector_l4" in df.columns:
        df["sector_l4"] = df["sector_l4"].fillna(df.get("sector_l3", "UNKNOWN")).fillna("UNKNOWN").astype(str).str.strip()
    if "bond_key" not in df.columns:
        df["bond_key"] = df.get("isin", pd.Series(np.nan, index=df.index)).where(df.get("isin", pd.Series(np.nan, index=df.index)).notna(), df.get("cusip", pd.Series(np.nan, index=df.index)))
    df["rating_bucket"] = df["rating"].map(rating_bucket).fillna("OTHER")
    df["rating_score"] = df["rating"].map(rating_score)
    if "maturity_bucket" not in df.columns or df["maturity_bucket"].isna().all():
        df["maturity_bucket"] = pd.cut(
            df["years_to_worst"],
            bins=[-np.inf, 2, 4, 6, 8, 10, np.inf],
            labels=["0-2", "2-4", "4-6", "6-8", "8-10", "10+"],
        ).astype(str).replace("nan", "UNKNOWN")
    if "curve_structure_bucket" not in df.columns:
        df["curve_structure_bucket"] = build_curve_structure_bucket(df)
    df["target_sector_group"] = target_sector_group(df)
    df["lender_stress_status"] = lender_stress_status(df)
    df["core_universe_status"] = core_universe_status(df)
    df["stress_label"] = stress_label(df)
    df["normal_rv_eligible"] = normal_rv_eligible(df)
    df["distressed_event"] = ~df["normal_rv_eligible"]
    df["flags"] = build_flags(df)
    if "rv_weight" not in df.columns:
        df["rv_weight"] = df.get("index_weight", pd.Series(np.nan, index=df.index)).copy()
        df.loc[df["rv_weight"].isna() | (df["rv_weight"] <= 0), "rv_weight"] = df.get("market_value", pd.Series(np.nan, index=df.index))
        df.loc[df["rv_weight"].isna() | (df["rv_weight"] <= 0), "rv_weight"] = 1.0
    return df


def finalize_history(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["as_of_date"] = pd.to_datetime(df["as_of_date"], errors="coerce")
    df = df.loc[df["as_of_date"].notna()].copy()


    df["ticker"] = df["ticker"].fillna("UNKNOWN").astype(str).str.upper().str.strip()
    df["rating_raw"] = df["rating"].fillna("NR").astype(str).str.upper().str.strip()
    df["rating"] = normalize_rating_notation(df["rating"])
    df["sector_l2"] = df["sector_l2"].fillna("UNKNOWN").astype(str).str.strip()
    df["sector_l3"] = df["sector_l3"].fillna(df["sector_l2"]).fillna("UNKNOWN").astype(str).str.strip()
    df["sector_l4"] = df["sector_l4"].fillna(df["sector_l3"]).fillna("UNKNOWN").astype(str).str.strip()


    df["bond_key"] = df["isin"].where(df["isin"].notna(), df["cusip"])
    missing_key = df["bond_key"].isna()
    df.loc[missing_key, "bond_key"] = (
        df.loc[missing_key, "ticker"].astype(str)
        + "|"
        + df.loc[missing_key, "description"].astype(str)
        + "|"
        + df.loc[missing_key, "maturity_date"].astype(str)
    )


    df["rating_bucket"] = df["rating"].map(rating_bucket).fillna("OTHER")
    df["rating_score"] = df["rating"].map(rating_score)
    df["maturity_bucket"] = pd.cut(
        df["years_to_worst"],
        bins=[-np.inf, 2, 4, 6, 8, 10, np.inf],
        labels=["0-2", "2-4", "4-6", "6-8", "8-10", "10+"],
    ).astype(str).replace("nan", "UNKNOWN")


    # Best-effort structure bucket used only for issuer-curve fitting.
    # ICE/H0A0 files do not always provide clean seniority/collateral fields, so this
    # uses conservative text classification from description/type fields. If the bucket
    # is unknown, the curve model can still fall back to an all-issuer curve with lower
    # confidence.
    df["curve_structure_bucket"] = build_curve_structure_bucket(df)
    df["target_sector_group"] = target_sector_group(df)
    df["lender_stress_status"] = lender_stress_status(df)


    df["core_universe_status"] = core_universe_status(df)
    df["stress_label"] = stress_label(df)
    df["normal_rv_eligible"] = normal_rv_eligible(df)
    df["distressed_event"] = ~df["normal_rv_eligible"]
    df["flags"] = build_flags(df)


    # Use index_weight if available, otherwise market value, otherwise equal weight.
    df["rv_weight"] = df["index_weight"].copy()
    df.loc[df["rv_weight"].isna() | (df["rv_weight"] <= 0), "rv_weight"] = df.loc[
        df["rv_weight"].isna() | (df["rv_weight"] <= 0), "market_value"
    ]
    df.loc[df["rv_weight"].isna() | (df["rv_weight"] <= 0), "rv_weight"] = 1.0


    return df.reset_index(drop=True)




def rating_bucket(rating: object) -> str:
    r = str(rating).upper().strip()
    # BBB must be tested before BB, otherwise crossover names collapse into the BB bucket.
    if r.startswith("BBB"):
        return "BBB"
    if r.startswith("AAA") or r.startswith("AA") or re.fullmatch(r"A[123+-]?", r):
        return "IG_ABOVE_BBB"
    if r.startswith("BB"):
        return "BB"
    if re.fullmatch(r"B[123+-]?", r) or r.startswith("B") and not r.startswith("BB"):
        return "B"
    if r.startswith("CCC") or r in {"CC", "C", "D"}:
        return "CCC_OR_BELOW"
    if r in {"CASH"}:
        return "CASH"
    return "OTHER"




def rating_score(rating: object) -> float:
    order = {
        "BBB1": -2,
        "BBB2": -1,
        "BBB3": 0,
        "BB1": 1,
        "BB2": 2,
        "BB3": 3,
        "B1": 4,
        "B2": 5,
        "B3": 6,
        "CCC1": 7,
        "CCC2": 8,
        "CCC3": 9,
        "CC": 10,
        "C": 11,
        "D": 12,
    }
    return float(order.get(str(rating).upper().strip(), np.nan))


_RATING_NOTATION_MAP = {
    # Moody's notation -> ICE composite notation
    "BA1": "BB1", "BA2": "BB2", "BA3": "BB3",
    "CAA1": "CCC1", "CAA2": "CCC2", "CAA3": "CCC3",
    "CA": "CC",
    # S&P / Fitch sign notation -> ICE composite notation
    "BAA1": "BBB1", "BAA2": "BBB2", "BAA3": "BBB3",
    "BBB+": "BBB1", "BBB": "BBB2", "BBB-": "BBB3",
    "BB+": "BB1", "BB": "BB2", "BB-": "BB3",
    "B+": "B1", "B": "B2", "B-": "B3",
    "CCC+": "CCC1", "CCC": "CCC2", "CCC-": "CCC3",
}


def normalize_rating_notation(rating: pd.Series) -> pd.Series:
    """Normalize rating strings to ICE composite notation (BB1..B3, CCC1..).

    ICE composite ratings are already BB1/BB2/... style, but new index entrants
    (fallen angels) and some export variants can arrive as agency notation
    ("BB+", "Ba1"), or carry watch/outlook markers ("BB1 *-", "B1u"). Without
    normalization those rows silently fail the exact CORE_RATINGS_EXACT match
    and drop out of the Core B/BB output even though they belong there.
    """
    r = rating.fillna("NR").astype(str).str.upper().str.strip()
    # Strip watch/outlook decorations: "*", "*+", "*-", "(P)", trailing "U"/"E",
    # and internal whitespace. Keep only rating-relevant characters first.
    r = r.str.replace(r"\s*\*[+-]?", "", regex=True)
    r = r.str.replace(r"\(P\)", "", regex=True)
    r = r.str.replace(r"\s+", "", regex=True)
    r = r.str.replace(r"(?<=[0-9+\-])[UE]$", "", regex=True)
    return r.map(lambda v: _RATING_NOTATION_MAP.get(v, v))






def lender_finance_mask(df: pd.DataFrame) -> pd.Series:
    """Identify lending / specialty-finance credits for a tighter core filter.


    These issuers can create exaggerated B/BB RV results when they are already in a
    stressed funding environment. The purpose is not to remove all financials from
    analysis; it is to keep stressed lenders out of the performing Core B/BB page
    and route them to the HY / Distressed Review page.
    """
    sector_l2 = df.get("sector_l2", pd.Series("", index=df.index)).fillna("").astype(str).str.upper()
    sector_l3 = df.get("sector_l3", pd.Series("", index=df.index)).fillna("").astype(str).str.upper()
    sector_l4 = df.get("sector_l4", pd.Series("", index=df.index)).fillna("").astype(str).str.upper()
    desc = df.get("description", pd.Series("", index=df.index)).fillna("").astype(str).str.upper()
    typ = df.get("type", pd.Series("", index=df.index)).fillna("").astype(str).str.upper()
    text = sector_l2 + " " + sector_l3 + " " + sector_l4 + " " + desc + " " + typ


    lender_terms = (
        r"FINANCIAL\s+SERVICES|CONSUMER\s+FINANCE|SPECIALTY\s+FINANCE|MORTGAGE|"
        r"LEND(?:ER|ING)?|LOAN|CREDIT|BANK|BROKER|CAPITAL\s+CORP|FUNDING|"
        r"FINANCE|FINANCING|RECEIVABLE|SERVICING|LEAS(?:E|ING)|AUTO\s+FINANCE|"
        r"BDC|BUSINESS\s+DEVELOPMENT|INSTALLMENT|PAYDAY|SUBPRIME"
    )
    return text.str.contains(lender_terms, regex=True, na=False)




def lender_stress_status(df: pd.DataFrame) -> pd.Series:
    """Classify lenders separately from the normal B/BB performing universe.


    Performing lenders can remain in Core B/BB. Stressed or distressed lenders are
    moved to the review page because their spread cheapness is often a funding,
    liquidity, asset-quality, or capital-markets risk signal rather than a clean
    mean-reversion RV setup.
    """
    is_lender = lender_finance_mask(df)
    rating = df.get("rating", pd.Series("", index=df.index)).fillna("").astype(str).str.upper().str.strip()
    rating_bucket_col = df.get("rating_bucket", pd.Series("OTHER", index=df.index)).astype(str).str.upper()
    oas = pd.to_numeric(df.get("oas", pd.Series(np.nan, index=df.index)), errors="coerce")
    price = pd.to_numeric(df.get("price", pd.Series(np.nan, index=df.index)), errors="coerce")
    ytw = pd.to_numeric(df.get("ytw", pd.Series(np.nan, index=df.index)), errors="coerce")


    out = pd.Series("Non-Lender", index=df.index, dtype=object)
    out = out.mask(is_lender, "Lender Performing")


    core_lender = is_lender & rating.isin(scored_ratings_exact()) & rating_bucket_col.isin(scored_rating_buckets())


    stressed = core_lender & (
        oas.ge(525)
        | price.lt(90)
        | ytw.ge(9.5)
        | (rating.eq("B3") & oas.ge(475))
    )
    distressed = core_lender & (
        oas.ge(650)
        | price.lt(85)
        | ytw.ge(12)
        | (rating.eq("B3") & oas.ge(575))
    )


    out = out.mask(stressed, "Stressed Lender Watch")
    out = out.mask(distressed, "Distressed Lender Review")
    return out


def core_universe_status(df: pd.DataFrame) -> pd.Series:
    """Classify bonds before RV scoring.


    v1.6 is a full high-yield split monitor with two separate analytical views:
      1. Core B/BB Performing RV: BB1-B3 bonds that are not stressed/distressed.
      2. HY / Distressed Review: CCC-and-below, stressed B/BB, insufficient-data,
         cash, and other non-core or distressed/event bonds.


    The main RV score is meant for the Core B/BB universe. The review universe is
    intentionally displayed on a different dashboard page with different columns and
    language so distressed / stressed credits are not treated as normal mean-reversion RV.
    """
    rating_bucket_col = df.get("rating_bucket", pd.Series("OTHER", index=df.index)).astype(str).str.upper()
    rating = df.get("rating", pd.Series("", index=df.index)).astype(str).str.upper()
    sector_l2 = df.get("sector_l2", pd.Series("", index=df.index)).astype(str).str.upper()
    oas = df.get("oas", pd.Series(np.nan, index=df.index))
    price = df.get("price", pd.Series(np.nan, index=df.index))
    ytw = df.get("ytw", pd.Series(np.nan, index=df.index))
    years = df.get("years_to_worst", pd.Series(np.nan, index=df.index))


    status = pd.Series("Core B/BB", index=df.index, dtype=object)
    exact_core = rating.isin(scored_ratings_exact()) & rating_bucket_col.isin(scored_rating_buckets())
    min_price = universe_min_price()
    max_oas = universe_max_oas()


    # Non-tradable / insufficient rows.
    status = status.mask((sector_l2 == "CASH") | (rating == "CASH"), "Cash/Excluded")
    status = status.mask(price.isna() | oas.isna(), "Insufficient Data")


    # Actual or near-default ratings are never scored.
    status = status.mask(rating.isin(NEVER_SCORED_RATINGS), "Distressed/Event")


    if CORE_UNIVERSE_MODE == "bb_b_only":
        status = status.mask(rating_bucket_col.eq("CCC_OR_BELOW"), "CCC / Distressed Review")
        status = status.mask((price < 70) | (oas >= 1000) | (ytw.fillna(0) >= 50) | (years.fillna(0) < 0.25), "Distressed/Event")
    else:
        # v2.1: only genuinely broken / recovery-value rows leave the scored book.
        # A CCC or crossover bond that still trades on spread stays in and is scored
        # against its own rating cohort.
        status = status.mask(
            (price < min_price) | (oas >= max_oas) | (ytw.fillna(0) >= 50) | (years.fillna(0) < 0.25),
            "Distressed/Event",
        )


    lender_status = df.get("lender_stress_status")
    if lender_status is None:
        lender_status = lender_stress_status(df)
    lender_status = lender_status.astype(str)


    if CORE_UNIVERSE_MODE == "bb_b_only":
        status = status.mask((status == "Core B/BB") & exact_core & lender_status.eq("Stressed Lender Watch"), "Stressed Lender Watch")
        status = status.mask((status == "Core B/BB") & exact_core & lender_status.eq("Distressed Lender Review"), "Distressed Lender Review")
        status = status.mask((status == "Core B/BB") & exact_core & ((price < STRESS_LABEL_PRICE) | (oas >= STRESS_LABEL_OAS)), "Stressed Watch")
    else:
        # Distressed lenders still route out; stressed lenders and stressed levels are
        # now labelled on the row (stress_label / flags) instead of dropping the bond.
        status = status.mask((status == "Core B/BB") & lender_status.eq("Distressed Lender Review"), "Distressed Lender Review")


    # Anything left without a scoreable rating notch goes to the review book.
    status = status.mask((status == "Core B/BB") & (~exact_core), "Other HY Review")
    return status


def stress_label(df: pd.DataFrame) -> pd.Series:
    """Non-excluding stress tag carried on scored rows.

    Under the v2.1 all-rated universe a wide or low-dollar-price bond is still scored,
    so the stress information moves from a filter into a column the dashboard can show.
    """
    price = pd.to_numeric(df.get("price", pd.Series(np.nan, index=df.index)), errors="coerce")
    oas = pd.to_numeric(df.get("oas", pd.Series(np.nan, index=df.index)), errors="coerce")
    lender_status = df.get("lender_stress_status", pd.Series("Non-Lender", index=df.index)).astype(str)


    out = pd.Series("Performing", index=df.index, dtype=object)
    out = out.mask(lender_status.eq("Stressed Lender Watch"), "Stressed Lender")
    out = out.mask((price < STRESS_LABEL_PRICE) | (oas >= STRESS_LABEL_OAS), "Stressed Levels")
    out = out.mask((price < 70) | (oas >= 1000), "Deeply Stressed")
    out = out.mask(price.isna() | oas.isna(), "Unknown")
    return out




def normal_rv_eligible(df: pd.DataFrame) -> pd.Series:
    rating = df["rating"].astype(str).str.upper()
    sector_l2 = df["sector_l2"].astype(str).str.upper()
    oas = df["oas"]
    price = df["price"]
    ytw = df["ytw"]
    years = df["years_to_worst"]


    rating_bucket_col = df.get("rating_bucket", pd.Series("OTHER", index=df.index)).astype(str).str.upper()
    core_status = df.get("core_universe_status", pd.Series("", index=df.index)).astype(str)
    lender_status = df.get("lender_stress_status", lender_stress_status(df)).astype(str)


    excluded_lender = (
        ["Stressed Lender Watch", "Distressed Lender Review"]
        if CORE_UNIVERSE_MODE == "bb_b_only"
        else ["Distressed Lender Review"]
    )


    return (
        (core_status == "Core B/BB")
        & (~lender_status.isin(excluded_lender))
        & rating.isin(scored_ratings_exact())
        & rating_bucket_col.isin(scored_rating_buckets())
        & (sector_l2 != "CASH")
        & (rating != "CASH")
        & (~rating.isin(NEVER_SCORED_RATINGS))
        & price.notna()
        & oas.notna()
        & (price >= universe_min_price())
        & (oas > 0)
        & (oas < universe_max_oas())
        & (ytw.fillna(0) < 50)
        & (years.fillna(0) >= 0.25)
    )


def build_flags(df: pd.DataFrame) -> pd.Series:
    flags: List[pd.Series] = []
    base = pd.Series([""] * len(df), index=df.index, dtype=object)


    def append_flag(condition: pd.Series, label: str) -> None:
        nonlocal base
        base = np.where(condition, pd.Series(base, index=df.index).astype(str) + label + ";", base)
        base = pd.Series(base, index=df.index, dtype=object)


    rating = df["rating"].astype(str).str.upper()
    sector_l2 = df["sector_l2"].astype(str).str.upper()
    append_flag((sector_l2 == "CASH") | (rating == "CASH"), "cash")
    append_flag(df["price"] < 70, "price_below_70")
    append_flag(df["price"] < 50, "price_below_50")
    append_flag(df["oas"] > 1000, "oas_above_1000")
    append_flag(df["oas"] > 2000, "oas_above_2000")
    append_flag(df["oas"] >= 9999, "oas_cap")
    append_flag(df["ytw"] >= 50, "ytw_cap")
    append_flag(rating.isin(["D", "C", "CC"]), "very_low_rating")
    append_flag(df["years_to_worst"] < 0.25, "short_to_worst")
    append_flag(df["oas"] <= 0, "negative_or_zero_oas")
    status = df.get("core_universe_status", pd.Series("", index=df.index)).astype(str)
    append_flag(status.eq("Non-Core Rating"), "non_core_rating")
    append_flag(status.eq("Stressed Watch"), "stressed_watch")
    stress = df.get("stress_label", pd.Series("", index=df.index)).astype(str)
    append_flag(stress.eq("Stressed Levels"), "stressed_levels")
    append_flag(stress.eq("Deeply Stressed"), "deeply_stressed")
    lender_status = df.get("lender_stress_status", pd.Series("Non-Lender", index=df.index)).astype(str)
    append_flag(lender_status.eq("Lender Performing"), "lender_finance")
    append_flag(lender_status.eq("Stressed Lender Watch"), "stressed_lender_watch")
    append_flag(lender_status.eq("Distressed Lender Review"), "distressed_lender_review")
    return base.str.strip(";").replace("", "none")




def build_curve_structure_bucket(df: pd.DataFrame) -> pd.Series:
    """Classify rough capital-structure bucket for issuer curve fitting.


    This is intentionally conservative and text-based. It prevents obvious secured,
    subordinated, and preferred bonds from being forced onto the same issuer curve
    when enough same-bucket bonds exist. Unknown buckets remain usable, but receive
    lower curve confidence if mixed with known structures.
    """
    desc = df.get("description", pd.Series("", index=df.index)).fillna("").astype(str)
    typ = df.get("type", pd.Series("", index=df.index)).fillna("").astype(str)
    mty = df.get("mty_type", pd.Series("", index=df.index)).fillna("").astype(str)
    text = (desc + " " + typ + " " + mty).str.upper()


    bucket = pd.Series("UNSPECIFIED", index=df.index, dtype=object)
    bucket = bucket.mask(text.str.contains(r"\b(?:1ST|FIRST)\s+LIEN\b|\b1L\b", regex=True, na=False), "FIRST_LIEN")
    bucket = bucket.mask(text.str.contains(r"\b(?:2ND|SECOND)\s+LIEN\b|\b2L\b", regex=True, na=False), "SECOND_LIEN")
    bucket = bucket.mask(text.str.contains(r"SECURED|\bSECD\b|\bSEC\b", regex=True, na=False), "SECURED")
    bucket = bucket.mask(text.str.contains(r"UNSECURED|UNSEC|SR\s+UNSEC|SENIOR\s+UNSEC", regex=True, na=False), "SENIOR_UNSECURED")
    bucket = bucket.mask(text.str.contains(r"SUBORDINATED|\bSUB\b|JR\s+SUB|JUNIOR", regex=True, na=False), "SUBORDINATED")
    bucket = bucket.mask(text.str.contains(r"PREFERRED|\bPFD\b|PREFERENCE", regex=True, na=False), "PREFERRED")
    bucket = bucket.mask(text.str.contains(r"CONVERT|CONV", regex=True, na=False), "CONVERTIBLE")
    return bucket




def latest_date(history: pd.DataFrame, requested: Optional[str] = None) -> pd.Timestamp:
    if requested:
        dt = pd.to_datetime(requested)
        if dt not in set(history["as_of_date"].dropna().unique()):
            available = history["as_of_date"].dropna().sort_values().unique()
            raise ValueError(f"Requested as-of date {requested} was not found. Latest available: {pd.Timestamp(available[-1]).date()}")
        return dt
    return pd.Timestamp(history["as_of_date"].max())




def percentile_of_current(history: pd.DataFrame, current: pd.DataFrame, value_col: str, key_col: str, days: Optional[int]) -> pd.DataFrame:
    cols = [key_col, value_col, "as_of_date"]
    hist = history.loc[history[value_col].notna(), cols].copy()
    cur = current[[key_col, value_col, "as_of_date"]].rename(columns={value_col: "current_value", "as_of_date": "current_date"})
    if days is not None:
        min_date = cur["current_date"].iloc[0] - pd.Timedelta(days=days)
        hist = hist.loc[hist["as_of_date"] >= min_date]
    merged = hist.merge(cur, on=key_col, how="inner")
    merged = merged.loc[merged["as_of_date"] <= merged["current_date"]]
    merged["le_current"] = merged[value_col] <= merged["current_value"]
    out = merged.groupby(key_col, observed=True).agg(
        obs=(value_col, "count"),
        le_current=("le_current", "sum"),
    )
    out["percentile"] = np.where(out["obs"] > 0, 100.0 * out["le_current"] / out["obs"], np.nan)
    return out[["obs", "percentile"]]




def historical_distribution_stats(history: pd.DataFrame, current: pd.DataFrame, value_col: str, key_col: str, days: Optional[int]) -> pd.DataFrame:
    """Return obs, median, p75, p90 and current distance from those levels.


    Percentiles alone can compress many names at 100 in wide markets. Distance-from-median
    gives the RV model a second dimension: how wide, not just widest-in-window.
    """
    cols = [key_col, value_col, "as_of_date"]
    hist = history.loc[history[value_col].notna(), cols].copy()
    cur = current[[key_col, value_col, "as_of_date"]].rename(columns={value_col: "current_value", "as_of_date": "current_date"})
    if days is not None and not cur.empty:
        hist = hist.loc[hist["as_of_date"] >= cur["current_date"].iloc[0] - pd.Timedelta(days=days)]
    merged = hist.merge(cur, on=key_col, how="inner")
    merged = merged.loc[merged["as_of_date"] <= merged["current_date"]]
    if merged.empty:
        return pd.DataFrame(columns=[key_col, "obs", "median", "p75", "p90", "vs_median", "vs_p75", "vs_p90"]).set_index(key_col)
    out = merged.groupby(key_col, observed=True).agg(
        obs=(value_col, "count"),
        median=(value_col, "median"),
        p75=(value_col, lambda x: float(np.nanpercentile(x, 75))),
        p90=(value_col, lambda x: float(np.nanpercentile(x, 90))),
        current_value=("current_value", "first"),
    )
    out["vs_median"] = out["current_value"] - out["median"]
    out["vs_p75"] = out["current_value"] - out["p75"]
    out["vs_p90"] = out["current_value"] - out["p90"]
    return out[["obs", "median", "p75", "p90", "vs_median", "vs_p75", "vs_p90"]]




def percentile_of_current_external(
    history: pd.DataFrame,
    current: pd.DataFrame,
    hist_value_col: str,
    current_value_col: str,
    key_col: str,
    days: Optional[int],
) -> pd.DataFrame:
    """Percentile of the current value versus history when columns have different names."""
    cols = [key_col, hist_value_col, "as_of_date"]
    hist = history.loc[history[hist_value_col].notna(), cols].copy()
    cur = current[[key_col, current_value_col, "as_of_date"]].rename(
        columns={current_value_col: "current_value", "as_of_date": "current_date"}
    )
    cur = cur.loc[cur["current_value"].notna()]
    if days is not None and not cur.empty:
        min_date = cur["current_date"].iloc[0] - pd.Timedelta(days=days)
        hist = hist.loc[hist["as_of_date"] >= min_date]
    merged = hist.merge(cur, on=key_col, how="inner")
    merged = merged.loc[merged["as_of_date"] <= merged["current_date"]]
    if merged.empty:
        return pd.DataFrame(columns=["obs", "percentile"]).rename_axis(key_col)
    merged["le_current"] = merged[hist_value_col] <= merged["current_value"]
    out = merged.groupby(key_col, observed=True).agg(
        obs=(hist_value_col, "count"),
        le_current=("le_current", "sum"),
    )
    out["percentile"] = np.where(out["obs"] > 0, 100.0 * out["le_current"] / out["obs"], np.nan)
    return out[["obs", "percentile"]]




def historical_distribution_stats_external(
    history: pd.DataFrame,
    current: pd.DataFrame,
    hist_value_col: str,
    current_value_col: str,
    key_col: str,
    days: Optional[int],
) -> pd.DataFrame:
    """Distribution stats for historical residuals where current and history columns differ."""
    cols = [key_col, hist_value_col, "as_of_date"]
    hist = history.loc[history[hist_value_col].notna(), cols].copy()
    cur = current[[key_col, current_value_col, "as_of_date"]].rename(
        columns={current_value_col: "current_value", "as_of_date": "current_date"}
    )
    cur = cur.loc[cur["current_value"].notna()]
    if days is not None and not cur.empty:
        hist = hist.loc[hist["as_of_date"] >= cur["current_date"].iloc[0] - pd.Timedelta(days=days)]
    merged = hist.merge(cur, on=key_col, how="inner")
    merged = merged.loc[merged["as_of_date"] <= merged["current_date"]]
    if merged.empty:
        return pd.DataFrame(columns=[key_col, "obs", "median", "p75", "p90", "vs_median", "vs_p75", "vs_p90"]).set_index(key_col)
    out = merged.groupby(key_col, observed=True).agg(
        obs=(hist_value_col, "count"),
        median=(hist_value_col, "median"),
        p75=(hist_value_col, lambda x: float(np.nanpercentile(x, 75))),
        p90=(hist_value_col, lambda x: float(np.nanpercentile(x, 90))),
        current_value=("current_value", "first"),
    )
    out["vs_median"] = out["current_value"] - out["median"]
    out["vs_p75"] = out["current_value"] - out["p75"]
    out["vs_p90"] = out["current_value"] - out["p90"]
    return out[["obs", "median", "p75", "p90", "vs_median", "vs_p75", "vs_p90"]]




def add_historical_peer_residuals(history: pd.DataFrame) -> pd.DataFrame:
    """Add historical peer residuals for true market-adjusted RV history.


    Raw OAS history can make many bonds look cheap when the whole market widens. This
    function creates a daily peer benchmark for every historical row and stores the
    residual to that benchmark. Later, the latest residual is compared to the bond's
    own residual history.
    """
    if "hist_peer_oas_residual" in history.columns:
        return history
    log("Calculating historical peer residual series for market-adjusted RV...")
    result = history.copy()
    result["hist_peer_group_used"] = pd.Series([np.nan] * len(result), dtype=object)
    result["hist_peer_count"] = np.nan
    result["hist_peer_median_oas"] = np.nan


    normal = result.loc[result["normal_rv_eligible"] & result["oas"].notna()].copy()
    if normal.empty:
        result["hist_peer_oas_residual"] = np.nan
        return result


    assigned = pd.Series(False, index=normal.index)
    peer_specs = [
        ("sector_l3_rating_maturity", ["as_of_date", "sector_l3", "rating_bucket", "maturity_bucket"], 10),
        ("sector_l3_rating", ["as_of_date", "sector_l3", "rating_bucket"], 12),
        ("rating_maturity", ["as_of_date", "rating_bucket", "maturity_bucket"], 15),
        ("rating", ["as_of_date", "rating_bucket"], 25),
        ("all_normal_hy", ["as_of_date"], 50),
    ]


    for group_name, keys, min_count in peer_specs:
        group_obj = normal.groupby(keys, observed=True)["oas"]
        counts = group_obj.transform("count")
        medians = group_obj.transform("median")
        use = (~assigned) & counts.ge(min_count)
        idx = normal.index[use]
        if len(idx) == 0:
            continue
        result.loc[idx, "hist_peer_group_used"] = group_name
        result.loc[idx, "hist_peer_count"] = counts.loc[use].values
        result.loc[idx, "hist_peer_median_oas"] = medians.loc[use].values
        assigned.loc[idx] = True


    result["hist_peer_oas_residual"] = result["oas"] - result["hist_peer_median_oas"]
    return result




def add_residual_history_percentiles(history: pd.DataFrame, latest_rows: pd.DataFrame) -> pd.DataFrame:
    """Compare today's peer residual to each bond's own historical peer residual range."""
    log("Calculating bond-level residual-history percentiles...")
    result = latest_rows.copy()
    if "hist_peer_oas_residual" not in history.columns or "peer_oas_residual" not in result.columns:
        for label in ["1y", "2y", "full"]:
            result[f"residual_oas_obs_{label}"] = np.nan
            result[f"residual_oas_pct_{label}"] = np.nan
        return result
    for label, days in [("1y", 365), ("2y", 730), ("full", None)]:
        pct = percentile_of_current_external(history, result, "hist_peer_oas_residual", "peer_oas_residual", "bond_key", days)
        result = result.merge(
            pct.rename(columns={"obs": f"residual_oas_obs_{label}", "percentile": f"residual_oas_pct_{label}"}),
            on="bond_key",
            how="left",
        )
    stats_1y = historical_distribution_stats_external(
        history, result, "hist_peer_oas_residual", "peer_oas_residual", "bond_key", 365
    ).rename(
        columns={
            "median": "residual_oas_1y_median",
            "p75": "residual_oas_1y_p75",
            "p90": "residual_oas_1y_p90",
            "vs_median": "residual_oas_vs_1y_median",
            "vs_p75": "residual_oas_vs_1y_p75",
            "vs_p90": "residual_oas_vs_1y_p90",
        }
    )
    result = result.merge(stats_1y.drop(columns=["obs"], errors="ignore"), on="bond_key", how="left")
    return result




def distance_score(series: pd.Series) -> pd.Series:
    """Rank positive distance as cheap and negative distance as rich on a 0-100 scale."""
    out = pd.Series(np.nan, index=series.index, dtype=float)
    mask = series.notna()
    if mask.sum() > 0:
        out.loc[mask] = series.loc[mask].rank(pct=True) * 100.0
    return out




def peer_group_quality_label(count_value: object) -> str:
    n = pd.to_numeric(pd.Series([count_value]), errors="coerce").iloc[0]
    if pd.isna(n):
        return "Weak"
    if n >= 30:
        return "Strong"
    if n >= 15:
        return "Acceptable"
    if n >= 8:
        return "Thin"
    return "Weak"




def stability_label(bp_value: object) -> str:
    n = pd.to_numeric(pd.Series([bp_value]), errors="coerce").iloc[0]
    if pd.isna(n):
        return "Unknown"
    if n <= 5:
        return "Strong"
    if n <= 15:
        return "Moderate"
    return "Fragile"




def cohort_context_label_from_pct(value: object) -> str:
    n = pd.to_numeric(pd.Series([value]), errors="coerce").iloc[0]
    if pd.isna(n):
        return "Cohort Fair"
    if n >= 80:
        return "Cohort Cheap/Wide"
    if n <= 20:
        return "Cohort Rich/Tight"
    return "Cohort Fair"




def relative_vs_absolute_label(signal: object, cohort_context: object) -> str:
    s = str(signal or "")
    c = str(cohort_context or "")
    if "Cheap" in s:
        if c == "Cohort Rich/Tight":
            return "Cheap to peers inside rich cohort"
        if c == "Cohort Cheap/Wide":
            return "Cheap in cheap cohort"
        return "Cheap to peers in fair cohort"
    if "Rich" in s:
        if c == "Cohort Cheap/Wide":
            return "Rich to peers inside cheap cohort"
        if c == "Cohort Rich/Tight":
            return "Rich in rich cohort"
        return "Rich to peers in fair cohort"
    return "Relative signal only"




def make_signal_caution_note(row: pd.Series) -> str:
    reasons: List[str] = []
    cohort_context = str(row.get("cohort_context_label", ""))
    signal = str(row.get("rv_signal", ""))
    peer_quality = str(row.get("peer_group_quality", ""))
    stability = str(row.get("peer_median_stability_label", ""))
    oas_obs = pd.to_numeric(pd.Series([row.get("oas_obs_1y")]), errors="coerce").iloc[0]
    residual_obs = pd.to_numeric(pd.Series([row.get("residual_oas_obs_1y")]), errors="coerce").iloc[0]
    if "Cheap" in signal and cohort_context == "Cohort Rich/Tight":
        reasons.append("whole cohort is historically tight, so cheapness may be only relative")
    if "Rich" in signal and cohort_context == "Cohort Cheap/Wide":
        reasons.append("whole cohort is historically wide, so richness may be only relative")
    if peer_quality in {"Thin", "Weak"}:
        reasons.append("peer group is thin")
    if stability == "Fragile":
        reasons.append("peer median is fragile")
    if (not pd.isna(oas_obs) and oas_obs < 120) or (not pd.isna(residual_obs) and residual_obs < 120):
        reasons.append("history is limited")
    if bool(row.get("issuer_distressed_event", False)):
        reasons.append("issuer-level stress may justify the spread")
    seen = set()
    ordered = []
    for r in reasons:
        if r not in seen:
            seen.add(r)
            ordered.append(r)
    return "; ".join(ordered)




def mode_or_unknown(values: pd.Series) -> str:
    vals = values.dropna().astype(str)
    if vals.empty:
        return "UNKNOWN"
    return vals.mode().iloc[0] if not vals.mode().empty else vals.iloc[0]




def add_security_percentiles(history: pd.DataFrame, latest_rows: pd.DataFrame) -> pd.DataFrame:
    log("Calculating bond-level historical OAS percentiles and distance-from-range stats...")
    result = latest_rows.copy()
    for label, days in [("1y", 365), ("2y", 730), ("full", None)]:
        pct = percentile_of_current(history, result, "oas", "bond_key", days)
        result = result.merge(pct.rename(columns={"obs": f"oas_obs_{label}", "percentile": f"oas_pct_{label}"}), on="bond_key", how="left")
    stats_1y = historical_distribution_stats(history, result, "oas", "bond_key", 365).rename(
        columns={
            "median": "oas_1y_median",
            "p75": "oas_1y_p75",
            "p90": "oas_1y_p90",
            "vs_median": "oas_vs_1y_median",
            "vs_p75": "oas_vs_1y_p75",
            "vs_p90": "oas_vs_1y_p90",
        }
    )
    result = result.merge(stats_1y.drop(columns=["obs"], errors="ignore"), on="bond_key", how="left")
    return result




def add_peer_metrics(history: pd.DataFrame, latest_rows: pd.DataFrame) -> pd.DataFrame:
    log("Calculating latest-date peer medians and peer residuals...")
    result = latest_rows.copy()
    normal = result.loc[result["normal_rv_eligible"] & result["oas"].notna()].copy()


    # Broad fallback peer groups, all on latest date only. This is intentionally fast and explainable.
    peer_specs = [
        ("sector_l4_rating_maturity", ["sector_l4", "rating_bucket", "maturity_bucket"]),
        ("sector_l3_rating_maturity", ["sector_l3", "rating_bucket", "maturity_bucket"]),
        ("sector_l3_rating", ["sector_l3", "rating_bucket"]),
        ("rating_maturity", ["rating_bucket", "maturity_bucket"]),
        ("rating", ["rating_bucket"]),
    ]


    result["peer_group_used"] = pd.Series([np.nan] * len(result), dtype=object)
    result["peer_count"] = np.nan
    result["peer_median_oas"] = np.nan


    assigned = pd.Series(False, index=result.index)
    for group_name, keys in peer_specs:
        stats = normal.groupby(keys, observed=True)["oas"].agg(peer_count="count", peer_median_oas="median").reset_index()
        candidate = result[keys].merge(stats, on=keys, how="left")
        use = (~assigned) & candidate["peer_count"].fillna(0).ge(10) & result["normal_rv_eligible"]
        result.loc[use, "peer_group_used"] = group_name
        result.loc[use, "peer_count"] = candidate.loc[use, "peer_count"].values
        result.loc[use, "peer_median_oas"] = candidate.loc[use, "peer_median_oas"].values
        assigned |= use


    # Final fallback: all normal bonds on latest date.
    fallback_median = normal["oas"].median()
    fallback_count = normal["oas"].count()
    use = (~assigned) & result["normal_rv_eligible"]
    result.loc[use, "peer_group_used"] = "all_normal_hy"
    result.loc[use, "peer_count"] = fallback_count
    result.loc[use, "peer_median_oas"] = fallback_median


    result["peer_oas_residual"] = result["oas"] - result["peer_median_oas"]


    # Latest-date percentile inside the selected group. Since groups may vary by fallback, use global residual rank as robust score.
    result["peer_oas_percentile_today"] = np.nan
    normal_idx = result.index[result["normal_rv_eligible"] & result["peer_oas_residual"].notna()]
    result.loc[normal_idx, "peer_oas_percentile_today"] = result.loc[normal_idx, "peer_oas_residual"].rank(pct=True) * 100.0
    return result




def add_cohort_context_and_peer_quality(history: pd.DataFrame, latest_rows: pd.DataFrame) -> pd.DataFrame:
    """Add cohort context, peer-quality diagnostics, and shrunk fair spread estimates."""
    # Reset to a simple RangeIndex so peer-group labels remain aligned after merges.
    result = latest_rows.copy().reset_index(drop=True)
    normal_latest = result.loc[result["normal_rv_eligible"] & result["oas"].notna()].copy()
    if normal_latest.empty:
        for col in [
            "peer_group_quality", "peer_median_stability_bp", "peer_median_stability_label",
            "fair_oas_shrunk", "shrunk_peer_residual", "cohort_group_used", "cohort_current_count",
            "cohort_median_oas", "cohort_oas_pct_1y", "cohort_oas_pct_full", "cohort_context_label"
        ]:
            result[col] = np.nan
        return result


    rm_stats = normal_latest.groupby(["rating_bucket", "maturity_bucket"], observed=True)["oas"].agg(rm_peer_median="median", rm_peer_count="count").reset_index()
    r_stats = normal_latest.groupby(["rating_bucket"], observed=True)["oas"].agg(r_peer_median="median", r_peer_count="count").reset_index()
    all_median = normal_latest["oas"].median()
    result = result.merge(rm_stats, on=["rating_bucket", "maturity_bucket"], how="left")
    result = result.merge(r_stats, on=["rating_bucket"], how="left")


    result["peer_group_quality"] = result["peer_count"].apply(peer_group_quality_label)
    result["peer_median_stability_bp"] = np.nan
    result["peer_median_stability_label"] = "Unknown"


    group_keys_map = {
        "sector_l4_rating_maturity": ["sector_l4", "rating_bucket", "maturity_bucket"],
        "sector_l3_rating_maturity": ["sector_l3", "rating_bucket", "maturity_bucket"],
        "sector_l3_rating": ["sector_l3", "rating_bucket"],
        "rating_maturity": ["rating_bucket", "maturity_bucket"],
        "rating": ["rating_bucket"],
        "all_normal_hy": [],
    }


    for group_name, keys in group_keys_map.items():
        mask = normal_latest["peer_group_used"].eq(group_name)
        if mask.sum() == 0:
            continue
        if keys:
            grouped = normal_latest.loc[mask].groupby(keys, observed=True)
            group_indexes = grouped.groups.values()
        else:
            group_indexes = [normal_latest.loc[mask].index.tolist()]
        for idxs in group_indexes:
            idxs = list(idxs)
            vals = pd.to_numeric(normal_latest.loc[idxs, "oas"], errors="coerce").dropna().to_numpy()
            if len(vals) <= 2:
                result.loc[idxs, "peer_median_stability_bp"] = np.nan
                result.loc[idxs, "peer_median_stability_label"] = "Fragile"
                continue
            full_med = float(np.nanmedian(vals))
            shifts = []
            for j in range(len(vals)):
                loo = np.delete(vals, j)
                if len(loo) == 0:
                    continue
                shifts.append(abs(float(np.nanmedian(loo)) - full_med))
            max_shift = max(shifts) if shifts else np.nan
            result.loc[idxs, "peer_median_stability_bp"] = max_shift
            result.loc[idxs, "peer_median_stability_label"] = stability_label(max_shift)


    def _shrunk_fair(row: pd.Series) -> float:
        peer = pd.to_numeric(pd.Series([row.get("peer_median_oas")]), errors="coerce").iloc[0]
        peer_count = pd.to_numeric(pd.Series([row.get("peer_count")]), errors="coerce").iloc[0]
        group_name = str(row.get("peer_group_used", ""))
        rm_median = pd.to_numeric(pd.Series([row.get("rm_peer_median")]), errors="coerce").iloc[0]
        r_median = pd.to_numeric(pd.Series([row.get("r_peer_median")]), errors="coerce").iloc[0]
        target = np.nan
        weight = 1.0
        if group_name in {"sector_l4_rating_maturity", "sector_l3_rating_maturity", "sector_l3_rating"}:
            target = rm_median if not pd.isna(rm_median) else r_median if not pd.isna(r_median) else all_median
            weight = float(np.clip(((peer_count if not pd.isna(peer_count) else 0.0) - 8.0) / 22.0, 0.0, 1.0))
        elif group_name == "rating_maturity":
            target = r_median if not pd.isna(r_median) else all_median
            weight = float(np.clip(((peer_count if not pd.isna(peer_count) else 0.0) - 10.0) / 15.0, 0.0, 1.0))
        elif group_name == "rating":
            target = all_median
            weight = float(np.clip(((peer_count if not pd.isna(peer_count) else 0.0) - 20.0) / 20.0, 0.0, 1.0))
        else:
            target = peer
            weight = 1.0
        if pd.isna(peer) and pd.isna(target):
            return np.nan
        if pd.isna(peer):
            return float(target)
        if pd.isna(target):
            return float(peer)
        return float(weight * peer + (1.0 - weight) * target)


    result["fair_oas_shrunk"] = result.apply(_shrunk_fair, axis=1)
    result["shrunk_peer_residual"] = result["oas"] - result["fair_oas_shrunk"]


    hist = history.loc[history["normal_rv_eligible"] & history["oas"].notna()].copy()


    # Component 2 v2.0 cohort valuation context.
    # Cohorts are intentionally built on exact B/BB rating notches and Level-4
    # industry classification, not broad rating buckets or broad sectors.
    # Example: BB3 + Software + 4-6 years, not BB + Technology.
    hist = hist.loc[hist["rating"].astype(str).str.upper().str.strip().isin(CORE_RATINGS_EXACT)].copy()
    normal_latest = normal_latest.loc[
        normal_latest["rating"].astype(str).str.upper().str.strip().isin(CORE_RATINGS_EXACT)
    ].copy()
    cohort_specs = [
        ("sector_l4_rating_notch_maturity", ["sector_l4", "rating", "maturity_bucket"], 8),
        ("sector_l4_rating_notch", ["sector_l4", "rating"], 10),
        ("rating_notch_maturity", ["rating", "maturity_bucket"], 15),
        ("rating_notch", ["rating"], 20),
    ]
    result["cohort_group_used"] = pd.Series([np.nan] * len(result), dtype=object)
    result["cohort_current_count"] = np.nan
    result["cohort_median_oas"] = np.nan
    assigned = pd.Series(False, index=result.index)
    cohort_daily_map: Dict[str, pd.DataFrame] = {}


    for group_name, keys, min_count in cohort_specs:
        daily = hist.groupby(["as_of_date"] + keys, observed=True)["oas"].agg(cohort_current_count="count", cohort_median_oas_hist="median").reset_index()
        cohort_daily_map[group_name] = daily
        stats_latest = normal_latest.groupby(keys, observed=True)["oas"].agg(cohort_current_count="count", cohort_median_oas="median").reset_index()
        candidate = result[keys].merge(stats_latest, on=keys, how="left")
        use = (~assigned) & candidate["cohort_current_count"].fillna(0).ge(min_count) & result["normal_rv_eligible"]
        result.loc[use, "cohort_group_used"] = group_name
        result.loc[use, "cohort_current_count"] = candidate.loc[use, "cohort_current_count"].values
        result.loc[use, "cohort_median_oas"] = candidate.loc[use, "cohort_median_oas"].values
        assigned |= use


    result["cohort_oas_pct_1y"] = np.nan
    result["cohort_oas_pct_full"] = np.nan


    def make_key(df: pd.DataFrame, group_name: str, keys: List[str]) -> pd.Series:
        if not keys:
            return pd.Series([group_name] * len(df), index=df.index)
        return group_name + "|" + df[keys].fillna("NA").astype(str).agg("|".join, axis=1)


    for group_name, keys, _ in cohort_specs:
        mask = result["cohort_group_used"].eq(group_name) & result["cohort_median_oas"].notna()
        if mask.sum() == 0:
            continue
        daily = cohort_daily_map[group_name].copy()
        daily["cohort_key"] = make_key(daily, group_name, keys)
        current = result.loc[mask, ["cohort_median_oas", "as_of_date"] + keys].copy()
        current["cohort_key"] = make_key(current, group_name, keys)
        for cohort_key, sub in current.groupby("cohort_key", observed=True):
            current_val = pd.to_numeric(sub["cohort_median_oas"].iloc[0], errors="coerce")
            current_date = pd.to_datetime(sub["as_of_date"].iloc[0])
            hist_vals = daily.loc[(daily["cohort_key"] == cohort_key) & (daily["as_of_date"] <= current_date), ["as_of_date", "cohort_median_oas_hist"]].copy()
            if hist_vals.empty or pd.isna(current_val):
                continue
            full_vals = pd.to_numeric(hist_vals["cohort_median_oas_hist"], errors="coerce").dropna()
            vals_1y = pd.to_numeric(hist_vals.loc[hist_vals["as_of_date"] >= current_date - pd.Timedelta(days=365), "cohort_median_oas_hist"], errors="coerce").dropna()
            full_pct = 100.0 * (full_vals <= current_val).sum() / len(full_vals) if len(full_vals) else np.nan
            pct_1y = 100.0 * (vals_1y <= current_val).sum() / len(vals_1y) if len(vals_1y) else np.nan
            result.loc[sub.index, "cohort_oas_pct_full"] = full_pct
            result.loc[sub.index, "cohort_oas_pct_1y"] = pct_1y


    cohort_pct = result["cohort_oas_pct_1y"].fillna(result["cohort_oas_pct_full"])
    result["cohort_context_label"] = cohort_pct.apply(cohort_context_label_from_pct)
    return result




def compute_issuer_daily(history: pd.DataFrame) -> pd.DataFrame:
    log("Calculating issuer-level daily histories with rating/sector peer context...")
    use = history.loc[
        history["ticker"].notna()
        & (history["ticker"] != "UNKNOWN")
        & history["oas"].notna()
    ].copy()


    # Company RV should be a relative-value measure, so issuer daily history is built from
    # normal-RV-eligible bonds. Distressed/event bonds are still retained via issuer flags.
    normal = use.loc[use["normal_rv_eligible"]].copy()
    if normal.empty:
        return pd.DataFrame()


    normal["oas_x_weight"] = normal["oas"] * normal["rv_weight"]
    normal["price_x_weight"] = normal["price"] * normal["rv_weight"]
    normal["ytw_x_weight"] = normal["ytw"] * normal["rv_weight"]


    grouped = normal.groupby(["as_of_date", "ticker"], observed=True).agg(
        issuer_bond_count=("bond_key", "nunique"),
        issuer_index_weight=("index_weight", "sum"),
        issuer_median_oas=("oas", "median"),
        issuer_median_price=("price", "median"),
        issuer_median_ytw=("ytw", "median"),
        issuer_rating_bucket=("rating_bucket", mode_or_unknown),
        issuer_core_universe_status=("core_universe_status", mode_or_unknown),
        issuer_sector_l3=("sector_l3", mode_or_unknown),
        issuer_sector_l4=("sector_l4", mode_or_unknown),
        issuer_target_sector_group=("target_sector_group", mode_or_unknown),
        issuer_maturity_bucket=("maturity_bucket", mode_or_unknown),
        weight_sum=("rv_weight", "sum"),
        oas_x_weight=("oas_x_weight", "sum"),
        price_x_weight=("price_x_weight", "sum"),
        ytw_x_weight=("ytw_x_weight", "sum"),
    ).reset_index()


    grouped["issuer_weighted_oas"] = np.where(grouped["weight_sum"] > 0, grouped["oas_x_weight"] / grouped["weight_sum"], grouped["issuer_median_oas"])
    grouped["issuer_weighted_price"] = np.where(grouped["weight_sum"] > 0, grouped["price_x_weight"] / grouped["weight_sum"], grouped["issuer_median_price"])
    grouped["issuer_weighted_ytw"] = np.where(grouped["weight_sum"] > 0, grouped["ytw_x_weight"] / grouped["weight_sum"], grouped["issuer_median_ytw"])
    grouped = grouped.drop(columns=["oas_x_weight", "price_x_weight", "ytw_x_weight"])


    # Add daily issuer-level peer medians. These are based on issuer medians, not all bonds,
    # so a large issuer does not overwhelm the peer group.
    grouped["issuer_peer_group_used"] = pd.Series([np.nan] * len(grouped), dtype=object)
    grouped["issuer_peer_count"] = np.nan
    grouped["issuer_peer_median_oas"] = np.nan
    assigned = pd.Series(False, index=grouped.index)
    peer_specs = [
        ("sector_l3_rating", ["as_of_date", "issuer_sector_l3", "issuer_rating_bucket"]),
        ("target_sector_rating", ["as_of_date", "issuer_target_sector_group", "issuer_rating_bucket"]),
        ("rating", ["as_of_date", "issuer_rating_bucket"]),
        ("all_normal_issuers", ["as_of_date"]),
    ]
    for group_name, keys in peer_specs:
        stats = grouped.groupby(keys, observed=True)["issuer_median_oas"].agg(issuer_peer_count="count", issuer_peer_median_oas="median").reset_index()
        candidate = grouped[keys].merge(stats, on=keys, how="left")
        min_count = 8 if group_name in {"sector_l3_rating", "target_sector_rating"} else 15 if group_name == "rating" else 25
        use_mask = (~assigned) & candidate["issuer_peer_count"].fillna(0).ge(min_count)
        grouped.loc[use_mask, "issuer_peer_group_used"] = group_name
        grouped.loc[use_mask, "issuer_peer_count"] = candidate.loc[use_mask, "issuer_peer_count"].values
        grouped.loc[use_mask, "issuer_peer_median_oas"] = candidate.loc[use_mask, "issuer_peer_median_oas"].values
        assigned |= use_mask


    grouped["issuer_peer_oas_residual"] = grouped["issuer_median_oas"] - grouped["issuer_peer_median_oas"]
    return grouped




def add_issuer_percentiles(issuer_daily: pd.DataFrame, as_of_date: pd.Timestamp, latest_rows: pd.DataFrame) -> pd.DataFrame:
    current = issuer_daily.loc[issuer_daily["as_of_date"] == as_of_date].copy()
    if current.empty:
        return current


    for label, days in [("1y", 365), ("2y", 730), ("full", None)]:
        hist = issuer_daily.copy()
        if days is not None:
            hist = hist.loc[hist["as_of_date"] >= as_of_date - pd.Timedelta(days=days)]
        cur = current[["ticker", "issuer_median_oas", "issuer_peer_oas_residual"]].rename(
            columns={
                "issuer_median_oas": "current_issuer_oas",
                "issuer_peer_oas_residual": "current_issuer_residual",
            }
        )
        merged = hist.merge(cur, on="ticker", how="inner")
        merged = merged.loc[merged["as_of_date"] <= as_of_date]
        merged["le_current_oas"] = merged["issuer_median_oas"] <= merged["current_issuer_oas"]
        merged["le_current_residual"] = merged["issuer_peer_oas_residual"] <= merged["current_issuer_residual"]
        pct = merged.groupby("ticker", observed=True).agg(
            obs=("issuer_median_oas", "count"),
            le_current_oas=("le_current_oas", "sum"),
            le_current_residual=("le_current_residual", "sum"),
            issuer_oas_median=("issuer_median_oas", "median"),
            issuer_oas_p75=("issuer_median_oas", lambda x: float(np.nanpercentile(x, 75))),
            issuer_oas_p90=("issuer_median_oas", lambda x: float(np.nanpercentile(x, 90))),
        )
        pct[f"issuer_oas_obs_{label}"] = pct["obs"]
        pct[f"issuer_oas_pct_{label}"] = 100.0 * pct["le_current_oas"] / pct["obs"]
        pct[f"issuer_residual_pct_{label}"] = 100.0 * pct["le_current_residual"] / pct["obs"]
        if label == "1y":
            pct["issuer_oas_1y_median"] = pct["issuer_oas_median"]
            pct["issuer_oas_1y_p75"] = pct["issuer_oas_p75"]
            pct["issuer_oas_1y_p90"] = pct["issuer_oas_p90"]
        keep_cols = [f"issuer_oas_obs_{label}", f"issuer_oas_pct_{label}", f"issuer_residual_pct_{label}"]
        if label == "1y":
            keep_cols += ["issuer_oas_1y_median", "issuer_oas_1y_p75", "issuer_oas_1y_p90"]
        current = current.merge(pct[keep_cols], on="ticker", how="left")


    current["issuer_oas_vs_1y_median"] = current["issuer_median_oas"] - current["issuer_oas_1y_median"]
    current["issuer_oas_vs_1y_p75"] = current["issuer_median_oas"] - current["issuer_oas_1y_p75"]
    current["issuer_oas_vs_1y_p90"] = current["issuer_median_oas"] - current["issuer_oas_1y_p90"]


    # Issuer-level distress/event filter uses all current bonds, not just normal-RV bonds.
    all_current = latest_rows.loc[latest_rows["ticker"].notna() & (latest_rows["ticker"] != "UNKNOWN")].copy()
    all_current["core_bond_flag"] = all_current["core_universe_status"].eq("Core B/BB")
    all_current["stressed_watch_flag"] = all_current["core_universe_status"].eq("Stressed Watch")
    all_current["stressed_lender_flag"] = all_current["core_universe_status"].eq("Stressed Lender Watch")
    all_current["distressed_lender_flag"] = all_current["core_universe_status"].eq("Distressed Lender Review")
    all_stats = all_current.groupby("ticker", observed=True).agg(
        all_bond_count=("bond_key", "nunique"),
        all_median_oas=("oas", "median"),
        all_median_price=("price", "median"),
        distressed_bond_count=("distressed_event", "sum"),
        core_bond_count=("core_bond_flag", "sum"),
        stressed_watch_bond_count=("stressed_watch_flag", "sum"),
        stressed_lender_bond_count=("stressed_lender_flag", "sum"),
        distressed_lender_bond_count=("distressed_lender_flag", "sum"),
    ).reset_index()
    all_stats["issuer_distressed_share"] = np.where(all_stats["all_bond_count"] > 0, all_stats["distressed_bond_count"] / all_stats["all_bond_count"], np.nan)
    current = current.merge(all_stats, on="ticker", how="left")


    flags = pd.Series([""] * len(current), index=current.index, dtype=object)
    def add_flag(mask: pd.Series, label: str) -> None:
        nonlocal flags
        flags = flags.mask(mask.fillna(False), flags + label + ";")


    add_flag(current["core_bond_count"].fillna(0).le(0), "no_core_b_or_bb_bonds")
    add_flag(current["all_median_oas"].ge(1000), "issuer_oas_above_1000")
    add_flag(current["all_median_oas"].ge(1500), "issuer_oas_above_1500")
    add_flag(current["all_median_oas"].ge(2500), "issuer_oas_above_2500")
    add_flag(current["all_median_price"].lt(70), "issuer_price_below_70")
    add_flag(current["issuer_distressed_share"].ge(0.50), "majority_distressed_bonds")
    add_flag(current["issuer_median_oas"].ge(1000), "core_issuer_oas_above_1000")
    add_flag(current["stressed_watch_bond_count"].fillna(0).gt(0), "stressed_watch_bonds_present")
    add_flag(current["stressed_lender_bond_count"].fillna(0).gt(0), "stressed_lender_bonds_present")
    add_flag(current["distressed_lender_bond_count"].fillna(0).gt(0), "distressed_lender_bonds_present")
    current["issuer_flags"] = flags.str.strip(";").replace("", "none")
    if CORE_UNIVERSE_MODE == "bb_b_only":
        # Stressed-watch issuers are not normal RV issuers. They remain review items.
        current["issuer_distressed_event"] = current["issuer_flags"].ne("none")
    else:
        # v2.1: wide or low-price issuers stay scored. Only issuers with no scoreable
        # bonds, majority-distressed capital structures, or levels past the hard
        # tradability gates go to the issuer review book. The other flags stay on the
        # row as information.
        current["issuer_distressed_event"] = (
            current["core_bond_count"].fillna(0).le(0)
            | current["all_median_oas"].ge(ALL_RATED_MAX_OAS)
            | current["all_median_price"].lt(ALL_RATED_MIN_PRICE)
            | current["issuer_distressed_share"].fillna(0).ge(0.50)
            | current["distressed_lender_bond_count"].fillna(0).gt(0)
        )
    current.loc[current["issuer_distressed_event"], "issuer_core_universe_status"] = "Issuer Review"


    # Confidence is not investment advice; it tells the user how much history/support exists.
    current["issuer_confidence"] = np.select(
        [
            current["issuer_oas_obs_1y"].fillna(0).ge(180) & current["issuer_bond_count"].fillna(0).ge(3) & current["issuer_peer_count"].fillna(0).ge(10),
            current["issuer_oas_obs_1y"].fillna(0).ge(90) & current["issuer_bond_count"].fillna(0).ge(2) & current["issuer_peer_count"].fillna(0).ge(8),
        ],
        ["High", "Medium"],
        default="Low",
    )
    return current
def _weighted_median(values: np.ndarray, weights: np.ndarray) -> float:
    values = np.asarray(values, dtype=float)
    weights = np.asarray(weights, dtype=float)
    valid = np.isfinite(values) & np.isfinite(weights) & (weights > 0)
    if valid.sum() == 0:
        return float(np.nanmedian(values)) if np.isfinite(values).any() else np.nan
    values = values[valid]
    weights = weights[valid]
    order = np.argsort(values)
    values = values[order]
    weights = weights[order]
    cumsum = np.cumsum(weights)
    cutoff = 0.5 * weights.sum()
    return float(values[np.searchsorted(cumsum, cutoff, side="left")])




def _weighted_line_fit(x: np.ndarray, y: np.ndarray, w: np.ndarray) -> Optional[np.ndarray]:
    x = np.asarray(x, dtype=float)
    y = np.asarray(y, dtype=float)
    w = np.asarray(w, dtype=float)
    valid = np.isfinite(x) & np.isfinite(y) & np.isfinite(w) & (w > 0)
    if valid.sum() < 2 or np.nanstd(x[valid]) < 1e-8:
        return None
    try:
        return np.polyfit(x[valid], y[valid], deg=1, w=np.sqrt(np.maximum(w[valid], 1e-8)))
    except Exception:
        return None




def _robust_weighted_line_fit(x: np.ndarray, y: np.ndarray, w: np.ndarray) -> Optional[np.ndarray]:
    """Two-pass weighted line fit with mild outlier downweighting.


    This avoids letting one odd bond dominate the issuer curve. The method remains
    simple and transparent: first fit a weighted line, then downweight observations
    with large residuals and refit.
    """
    coeff = _weighted_line_fit(x, y, w)
    if coeff is None:
        return None
    for _ in range(2):
        pred = coeff[0] * x + coeff[1]
        resid = y - pred
        med = np.nanmedian(resid)
        mad = np.nanmedian(np.abs(resid - med))
        scale = 1.4826 * mad if np.isfinite(mad) and mad > 1e-8 else np.nanstd(resid)
        if not np.isfinite(scale) or scale < 1e-8:
            break
        robust_w = 1.0 / (1.0 + (np.abs(resid - med) / (3.0 * scale)) ** 2)
        new_coeff = _weighted_line_fit(x, y, w * robust_w)
        if new_coeff is None:
            break
        coeff = new_coeff
    return coeff




def _weighted_r2(y: np.ndarray, fair: np.ndarray, w: np.ndarray) -> float:
    y = np.asarray(y, dtype=float)
    fair = np.asarray(fair, dtype=float)
    w = np.asarray(w, dtype=float)
    valid = np.isfinite(y) & np.isfinite(fair) & np.isfinite(w) & (w > 0)
    if valid.sum() < 2:
        return np.nan
    yv = y[valid]
    fv = fair[valid]
    wv = w[valid]
    ybar = np.average(yv, weights=wv)
    ss_res = np.sum(wv * (yv - fv) ** 2)
    ss_tot = np.sum(wv * (yv - ybar) ** 2)
    if ss_tot <= 1e-8:
        return np.nan
    return float(1.0 - ss_res / ss_tot)




def _fit_curve_segment(segment: pd.DataFrame, method_context: str) -> Dict[str, object]:
    """Fit a quality-controlled issuer curve segment.


    Returns fair values for the segment's rows plus diagnostics. The preferred method
    is leave-one-out robust weighted line fitting, because it prevents the target bond
    from pulling its own fair value toward itself. For smaller or very flat curves, the
    method falls back to lower-confidence alternatives.
    """
    seg = segment.copy()
    x = seg["years_to_worst"].astype(float).to_numpy()
    y = seg["oas"].astype(float).to_numpy()
    w = seg["rv_weight"].fillna(1.0).astype(float).to_numpy()
    valid = np.isfinite(x) & np.isfinite(y) & np.isfinite(w) & (w > 0)
    idx = seg.index[valid]
    x = x[valid]
    y = y[valid]
    w = w[valid]
    n = int(valid.sum())


    if n < 3:
        return {"index": idx, "fair": np.array([]), "method": "not_calculated", "confidence": "Unavailable", "reason": "insufficient_clean_curve_bonds", "r2": np.nan, "x_span": np.nan, "model_n": n}


    x_span = float(np.nanmax(x) - np.nanmin(x)) if n else np.nan
    fair = np.full(n, np.nan, dtype=float)
    method = ""
    reason = ""


    if x_span < 0.50:
        fair[:] = _weighted_median(y, w)
        method = "flat_same_issuer_median"
        confidence = "Low"
        reason = "low_years_to_worst_dispersion"
    elif n >= 4:
        for i in range(n):
            train = np.ones(n, dtype=bool)
            train[i] = False
            train_x = x[train]
            train_y = y[train]
            train_w = w[train]
            if len(train_x) >= 3 and (np.nanmax(train_x) - np.nanmin(train_x)) >= 0.50:
                coeff = _robust_weighted_line_fit(train_x, train_y, train_w)
                if coeff is not None:
                    fair[i] = coeff[0] * x[i] + coeff[1]
                else:
                    fair[i] = _weighted_median(train_y, train_w)
            elif len(train_x) >= 2 and np.nanstd(train_x) > 1e-8:
                coeff = _weighted_line_fit(train_x, train_y, train_w)
                fair[i] = coeff[0] * x[i] + coeff[1] if coeff is not None else _weighted_median(train_y, train_w)
            else:
                fair[i] = _weighted_median(train_y, train_w)
        method = "leave_one_out_robust_line"
        if n >= 5 and x_span >= 1.50:
            confidence = "High"
            reason = "clean_multi_bond_curve"
        elif x_span >= 0.75:
            confidence = "Medium"
            reason = "adequate_multi_bond_curve"
        else:
            confidence = "Low"
            reason = "limited_curve_span"
    else:
        coeff = _robust_weighted_line_fit(x, y, w)
        if coeff is not None:
            fair[:] = coeff[0] * x + coeff[1]
            method = "three_bond_robust_line"
            confidence = "Medium" if x_span >= 1.50 else "Low"
            reason = "three_bond_curve_low_sample"
        else:
            fair[:] = _weighted_median(y, w)
            method = "flat_same_issuer_median"
            confidence = "Low"
            reason = "three_bond_curve_fallback_flat"


    r2 = _weighted_r2(y, fair, w)
    residuals = y - fair
    if np.isfinite(residuals).any() and np.nanmax(np.abs(residuals)) > 500:
        # Extreme residuals are often real credit information, but they are also a sign
        # that structure/call/liquidity may be distorting the curve. Keep the value but
        # demote confidence and explicitly flag it.
        if confidence == "High":
            confidence = "Medium"
        elif confidence == "Medium":
            confidence = "Low"
        reason = reason + ";extreme_curve_residual_review"


    return {
        "index": idx,
        "fair": fair,
        "method": f"{method_context}:{method}",
        "confidence": confidence,
        "reason": reason,
        "r2": r2,
        "x_span": x_span,
        "model_n": n,
    }




def add_issuer_curve_residuals(latest_rows: pd.DataFrame) -> pd.DataFrame:
    log("Calculating quality-controlled issuer-curve residuals on latest date...")
    result = latest_rows.copy()
    result["issuer_curve_fair_oas"] = np.nan
    result["issuer_curve_residual"] = np.nan
    result["issuer_curve_bond_count"] = np.nan
    result["issuer_curve_model_bond_count"] = np.nan
    result["issuer_curve_method"] = "not_calculated"
    result["issuer_curve_confidence"] = "Unavailable"
    result["issuer_curve_reason"] = "not_enough_clean_same_issuer_bonds"
    result["issuer_curve_r2"] = np.nan
    result["issuer_curve_x_span"] = np.nan
    result["issuer_curve_segment_key"] = ""


    normal = result.loc[
        result["normal_rv_eligible"]
        & result["ticker"].notna()
        & (result["ticker"] != "UNKNOWN")
        & result["oas"].notna()
        & result["years_to_worst"].notna()
    ].copy()


    if normal.empty:
        return result


    assigned = pd.Series(False, index=result.index)
    fitted_segments = 0


    for ticker, group in normal.groupby("ticker", observed=True):
        ticker_idx = group.index
        result.loc[ticker_idx, "issuer_curve_bond_count"] = len(group)
        if len(group) < 3:
            result.loc[ticker_idx, "issuer_curve_reason"] = f"insufficient_clean_curve_bonds_{len(group)}"
            continue


        bucket_counts = group["curve_structure_bucket"].fillna("UNSPECIFIED").value_counts()
        meaningful = bucket_counts.drop(labels=["UNSPECIFIED"], errors="ignore")
        segment_assigned = pd.Series(False, index=group.index)


        # First fit clean same-structure segments when available. This avoids comparing
        # secured, subordinated, preferred, and unsecured bonds on a single line when the
        # data supports a more comparable curve.
        for bucket, count in bucket_counts.items():
            if count < 3 or bucket == "UNSPECIFIED":
                continue
            segment = group.loc[group["curve_structure_bucket"].fillna("UNSPECIFIED") == bucket]
            fit = _fit_curve_segment(segment, "same_structure")
            if len(fit["index"]) == 0 or len(fit["fair"]) == 0:
                continue
            idx = fit["index"]
            result.loc[idx, "issuer_curve_fair_oas"] = fit["fair"]
            result.loc[idx, "issuer_curve_residual"] = result.loc[idx, "oas"].astype(float).to_numpy() - fit["fair"]
            result.loc[idx, "issuer_curve_model_bond_count"] = fit["model_n"]
            result.loc[idx, "issuer_curve_method"] = fit["method"]
            result.loc[idx, "issuer_curve_confidence"] = fit["confidence"]
            result.loc[idx, "issuer_curve_reason"] = fit["reason"]
            result.loc[idx, "issuer_curve_r2"] = fit["r2"]
            result.loc[idx, "issuer_curve_x_span"] = fit["x_span"]
            result.loc[idx, "issuer_curve_segment_key"] = f"{ticker}|{bucket}"
            segment_assigned.loc[idx] = True
            assigned.loc[idx] = True
            fitted_segments += 1


        remaining = group.loc[~segment_assigned]
        if remaining.empty:
            continue


        # If there are multiple known structure buckets and the remaining bonds do not
        # have enough same-structure observations, do not force an all-issuer curve.
        # This maximizes accuracy by preferring no curve over a contaminated curve.
        mixed_known_structures = len(meaningful) >= 2
        remaining_known = remaining.loc[
            remaining["curve_structure_bucket"].fillna("UNSPECIFIED") != "UNSPECIFIED",
            "curve_structure_bucket",
        ].nunique()
        if mixed_known_structures and remaining_known >= 2:
            result.loc[remaining.index, "issuer_curve_reason"] = "mixed_capital_structure_no_same_structure_curve"
            continue
        if mixed_known_structures and segment_assigned.any() and remaining_known >= 1:
            result.loc[remaining.index, "issuer_curve_reason"] = "rare_structure_insufficient_same_structure_bonds"
            continue
        if mixed_known_structures and len(remaining) < 3:
            result.loc[remaining.index, "issuer_curve_reason"] = "mixed_capital_structure_insufficient_same_structure_bonds"
            continue


        # Use an all-issuer curve when structure is unknown or mostly homogeneous. If it
        # mixes known and unknown structures, the fitted segment will still be retained,
        # but confidence is capped below High.
        if len(remaining) >= 3:
            method_context = "all_issuer"
            fit = _fit_curve_segment(remaining, method_context)
            if len(fit["index"]) == 0 or len(fit["fair"]) == 0:
                result.loc[remaining.index, "issuer_curve_reason"] = fit.get("reason", "curve_fit_failed")
                continue
            idx = fit["index"]
            confidence = str(fit["confidence"])
            reason = str(fit["reason"])
            remaining_structures = remaining["curve_structure_bucket"].fillna("UNSPECIFIED").nunique()
            if remaining_structures > 1 and confidence == "High":
                confidence = "Medium"
                reason = reason + ";mixed_or_unspecified_structure"


            result.loc[idx, "issuer_curve_fair_oas"] = fit["fair"]
            result.loc[idx, "issuer_curve_residual"] = result.loc[idx, "oas"].astype(float).to_numpy() - fit["fair"]
            result.loc[idx, "issuer_curve_model_bond_count"] = fit["model_n"]
            result.loc[idx, "issuer_curve_method"] = fit["method"]
            result.loc[idx, "issuer_curve_confidence"] = confidence
            result.loc[idx, "issuer_curve_reason"] = reason
            result.loc[idx, "issuer_curve_r2"] = fit["r2"]
            result.loc[idx, "issuer_curve_x_span"] = fit["x_span"]
            result.loc[idx, "issuer_curve_segment_key"] = f"{ticker}|ALL"
            assigned.loc[idx] = True
            fitted_segments += 1
        else:
            result.loc[remaining.index, "issuer_curve_reason"] = f"insufficient_remaining_curve_bonds_{len(remaining)}"


    curve_available = result["issuer_curve_residual"].notna().sum()
    log(f"Issuer-curve residual coverage: {curve_available:,}/{len(result):,} latest bonds across {fitted_segments:,} curve segment(s).")
    return result


def score_latest(latest_rows: pd.DataFrame, issuer_current: pd.DataFrame) -> pd.DataFrame:
    log("Scoring latest-date bonds with v2.0 internal-OAS weighted framework...")
    result = latest_rows.copy()
    result["bond_label"] = result.apply(make_bond_label, axis=1)
    issuer_cols = [
        "ticker",
        "issuer_oas_pct_1y", "issuer_oas_pct_2y", "issuer_oas_pct_full",
        "issuer_residual_pct_1y", "issuer_residual_pct_2y", "issuer_residual_pct_full",
        "issuer_oas_vs_1y_median", "issuer_oas_vs_1y_p75",
        "issuer_distressed_event", "issuer_confidence",
    ]
    result = result.merge(issuer_current[[c for c in issuer_cols if c in issuer_current.columns]].drop_duplicates("ticker"), on="ticker", how="left")


    normal_mask = result["normal_rv_eligible"].copy()


    def _weighted_component(frame: pd.DataFrame, weights: Dict[str, float]) -> pd.Series:
        """Weighted average of 0-100 component columns, normalized by active weights."""
        cols = [c for c in weights if c in frame.columns]
        if not cols:
            return pd.Series(np.nan, index=frame.index, dtype=float)
        w = pd.Series({c: weights[c] for c in cols}, dtype=float)
        part = frame[cols].apply(pd.to_numeric, errors="coerce")
        weighted = part.mul(w, axis=1).sum(axis=1, skipna=True)
        active_w = part.notna().mul(w, axis=1).sum(axis=1)
        return pd.Series(np.where(active_w > 0, weighted / active_w, np.nan), index=frame.index, dtype=float)


    # 1) Historical peer-adjusted residual score: the model's core edge.
    # This asks whether today's peer residual is unusual versus the bond's own
    # historical peer-residual relationship. Weighting favors recent regimes but
    # keeps 2Y/full-history context.
    residual_frame = pd.DataFrame(index=result.index)
    residual_frame["residual_1y"] = result.get("residual_oas_pct_1y", pd.Series(np.nan, index=result.index))
    residual_frame["residual_2y"] = result.get("residual_oas_pct_2y", pd.Series(np.nan, index=result.index))
    residual_frame["residual_full"] = result.get("residual_oas_pct_full", pd.Series(np.nan, index=result.index))
    result["historical_peer_adjusted_residual_score"] = _weighted_component(
        residual_frame,
        {"residual_1y": 0.45, "residual_2y": 0.35, "residual_full": 0.20},
    )


    # 2) Cohort valuation context: whether the whole yardstick is rich/fair/cheap.
    cohort_frame = pd.DataFrame(index=result.index)
    cohort_frame["cohort_1y"] = result.get("cohort_oas_pct_1y", pd.Series(np.nan, index=result.index))
    cohort_frame["cohort_full"] = result.get("cohort_oas_pct_full", pd.Series(np.nan, index=result.index))
    result["cohort_context_score"] = _weighted_component(
        cohort_frame,
        {"cohort_1y": 0.70, "cohort_full": 0.30},
    )


    # 3) Current shrunk peer residual: today's cross-sectional dislocation using
    # a shrunk fair OAS so thin cohorts do not over-control fair value.
    result["peer_score"] = np.nan
    peer_source = result.get("shrunk_peer_residual", result.get("peer_oas_residual", pd.Series(np.nan, index=result.index)))
    peer_mask = normal_mask & peer_source.notna()
    result.loc[peer_mask, "peer_score"] = peer_source.loc[peer_mask].rank(pct=True) * 100.0
    result["current_shrunk_peer_residual_score"] = result["peer_score"]


    # 4) Issuer context: issuer-level historical and peer-relative context.
    result["issuer_distance_score"] = distance_score(result["issuer_oas_vs_1y_median"])
    issuer_frame = pd.DataFrame(index=result.index)
    issuer_frame["issuer_residual_1y"] = result.get("issuer_residual_pct_1y", pd.Series(np.nan, index=result.index))
    issuer_frame["issuer_oas_1y"] = result.get("issuer_oas_pct_1y", pd.Series(np.nan, index=result.index))
    issuer_frame["issuer_distance"] = result.get("issuer_distance_score", pd.Series(np.nan, index=result.index))
    result["issuer_context_score"] = _weighted_component(
        issuer_frame,
        {"issuer_residual_1y": 0.60, "issuer_oas_1y": 0.20, "issuer_distance": 0.20},
    )


    # 5) Bond's own raw OAS history: useful as a sanity check, but intentionally
    # low weight because raw OAS can make every bond look cheap in a broad selloff.
    result["own_distance_score"] = distance_score(result["oas_vs_1y_median"])
    own_frame = pd.DataFrame(index=result.index)
    own_frame["own_oas_1y"] = result.get("oas_pct_1y", pd.Series(np.nan, index=result.index))
    own_frame["own_oas_2y"] = result.get("oas_pct_2y", pd.Series(np.nan, index=result.index))
    own_frame["own_oas_full"] = result.get("oas_pct_full", pd.Series(np.nan, index=result.index))
    own_frame["own_distance"] = result.get("own_distance_score", pd.Series(np.nan, index=result.index))
    result["own_oas_history_score"] = _weighted_component(
        own_frame,
        {"own_oas_1y": 0.50, "own_oas_2y": 0.20, "own_oas_full": 0.10, "own_distance": 0.20},
    )


    # 6) Spread momentum / recent cheapening. This remains a small overlay because
    # we do not yet have TRACE/volume/catalyst data. Higher score means more recent
    # widening/cheapening versus the rest of the normal RV universe.
    result["momentum_score"] = np.nan
    mom_frame = pd.DataFrame(index=result.index)
    for col, out_col in [("oas_1d_chg", "mom_1d"), ("oas_1w_chg", "mom_1w"), ("oas_1m_chg", "mom_1m")]:
        if col in result.columns:
            vals = pd.to_numeric(result[col], errors="coerce")
            ranked = pd.Series(np.nan, index=result.index, dtype=float)
            mask = normal_mask & vals.notna()
            if mask.sum() > 1:
                ranked.loc[mask] = vals.loc[mask].rank(pct=True) * 100.0
            mom_frame[out_col] = ranked
    if not mom_frame.empty:
        result["momentum_score"] = _weighted_component(mom_frame, {"mom_1d": 0.10, "mom_1w": 0.45, "mom_1m": 0.45})


    # Issuer curve residual remains available for diagnostics and switch candidates,
    # but it no longer contributes an overlay to rv_score in v2.0.
    result["curve_score_raw"] = np.nan
    result["curve_score"] = np.nan
    result["issuer_curve_overlay_points"] = 0.0
    curve_confidence = result.get("issuer_curve_confidence", pd.Series("Unavailable", index=result.index)).fillna("Unavailable")
    curve_mask = normal_mask & result["issuer_curve_residual"].notna() & curve_confidence.isin(["High", "Medium"])
    result.loc[curve_mask, "curve_score_raw"] = result.loc[curve_mask, "issuer_curve_residual"].rank(pct=True) * 100.0
    curve_factor = curve_confidence.map({"High": 1.00, "Medium": 0.65}).fillna(0.0)
    result.loc[curve_mask, "curve_score"] = 50.0 + (result.loc[curve_mask, "curve_score_raw"] - 50.0) * curve_factor.loc[curve_mask]


    # v2.0 internal-data weighting. These are true percentages that sum to
    # 100%, and the final average is still normalized by active weights when a
    # component is unavailable. Momentum is removed from the final score.
    # Issuer-curve residuals remain diagnostics only.
    components = pd.DataFrame(index=result.index)
    components["historical_peer_adjusted_residual_score"] = result["historical_peer_adjusted_residual_score"]
    components["cohort_context_score"] = result["cohort_context_score"]
    components["current_shrunk_peer_residual_score"] = result["current_shrunk_peer_residual_score"]
    components["issuer_context_score"] = result["issuer_context_score"]
    components["own_oas_history_score"] = result["own_oas_history_score"]
    weights = pd.Series({
        "historical_peer_adjusted_residual_score": 0.40,
        "cohort_context_score": 0.20,
        "current_shrunk_peer_residual_score": 0.15,
        "issuer_context_score": 0.15,
        "own_oas_history_score": 0.10,
    })


    weighted_sum = components.mul(weights, axis=1).sum(axis=1, skipna=True)
    weight_sum = components.notna().mul(weights, axis=1).sum(axis=1)
    result["rv_score_base"] = np.where(weight_sum > 0, weighted_sum / weight_sum, np.nan)
    result["rv_score"] = result["rv_score_base"].clip(lower=0.0, upper=100.0)
    result["rv_score_method"] = "v2.0: 40 hist residual + 20 Level-4/notch cohort + 15 shrunk peer + 15 issuer + 10 own OAS; no momentum; no curve overlay"


    # Do not allow issuer-level distressed/event companies to appear as normal bond RV candidates.
    issuer_distress = result.get("issuer_distressed_event", pd.Series(False, index=result.index)).fillna(False).astype(bool)
    exact_core_rating = result["rating"].astype(str).str.upper().str.strip().isin(CORE_RATINGS_EXACT)
    lender_status = result.get("lender_stress_status", pd.Series("Non-Lender", index=result.index)).astype(str)
    core_status = result.get("core_universe_status", pd.Series("", index=result.index)).astype(str)
    result.loc[issuer_distress | (~exact_core_rating) | core_status.ne("Core B/BB") | lender_status.isin(["Stressed Lender Watch", "Distressed Lender Review"]), "normal_rv_eligible"] = False
    normal_mask = result["normal_rv_eligible"]
    result.loc[~normal_mask, "rv_score"] = np.nan
    result.loc[~normal_mask, "rv_score_base"] = np.nan
    result.loc[~normal_mask, "issuer_curve_overlay_points"] = 0.0


    residual_obs = result.get("residual_oas_obs_1y", pd.Series(np.nan, index=result.index)).fillna(0)
    issuer_conf = result.get("issuer_confidence", pd.Series("Low", index=result.index)).fillna("Low")
    support_score = pd.Series(0.0, index=result.index, dtype=float)
    support_score += np.where(result["oas_obs_1y"].fillna(0).ge(180), 1.0, np.where(result["oas_obs_1y"].fillna(0).ge(90), 0.5, 0.0))
    support_score += np.where(residual_obs.ge(180), 1.0, np.where(residual_obs.ge(90), 0.5, 0.0))
    support_score += result.get("peer_group_quality", pd.Series("Weak", index=result.index)).map({"Strong": 1.0, "Acceptable": 0.7, "Thin": 0.3, "Weak": 0.0}).fillna(0.0)
    support_score += result.get("peer_median_stability_label", pd.Series("Unknown", index=result.index)).map({"Strong": 0.8, "Moderate": 0.4, "Fragile": 0.0}).fillna(0.0)
    support_score += issuer_conf.map({"High": 0.5, "Medium": 0.25}).fillna(0.0)


    result["model_support"] = np.select(
        [support_score >= 3.0, support_score >= 1.8],
        ["High", "Medium"],
        default="Low",
    )
    result["history_confidence"] = result["model_support"]
    result["rv_signal"] = result["rv_score"].apply(score_to_signal)
    result.loc[~normal_mask, "rv_signal"] = "Distressed/Event Review"
    result["relative_vs_absolute_label"] = result.apply(lambda r: relative_vs_absolute_label(r.get("rv_signal"), r.get("cohort_context_label")), axis=1)
    result["why_signal_may_be_wrong"] = result.apply(make_signal_caution_note, axis=1)
    result["rv_note"] = result.apply(make_bond_note, axis=1)
    return result


def score_to_signal(score: object) -> str:
    if pd.isna(score):
        return "No Score"
    score = float(score)
    if score >= 80:
        return "Very Cheap"
    if score >= 65:
        return "Cheap"
    if score > 35:
        return "Neutral"
    if score > 20:
        return "Rich"
    return "Very Rich"




def issuer_signal_from_score(score: object, distressed: object = False) -> str:
    if bool(distressed):
        return "Distressed/Event Review"
    if pd.isna(score):
        return "No Score"
    score = float(score)
    if score >= 80:
        return "Very Cheap"
    if score >= 65:
        return "Cheap"
    if score > 35:
        return "Neutral"
    if score > 20:
        return "Rich"
    return "Very Rich"




def fmt_num(value: object, decimals: int = 1) -> str:
    if pd.isna(value):
        return "n/a"
    return f"{float(value):,.{decimals}f}"


def fmt_coupon(value: object) -> str:
    if pd.isna(value):
        return ""
    try:
        return f"{float(value):.3f}".rstrip("0").rstrip(".") + "%"
    except Exception:
        return str(value)




def fmt_maturity(value: object) -> str:
    if pd.isna(value):
        return ""
    try:
        return pd.to_datetime(value).strftime("%m/%d/%Y")
    except Exception:
        return str(value)




def make_bond_label(row: pd.Series) -> str:
    coupon = fmt_coupon(row.get("coupon"))
    maturity = fmt_maturity(row.get("maturity_date"))
    pieces = [str(row.get("ticker", "")).strip(), coupon, maturity]
    label = " ".join([p for p in pieces if p and p.lower() != "nan"])
    return label if label else str(row.get("description", ""))




def make_bond_snippet(row: pd.Series) -> str:
    cusip = str(row.get("cusip", "")).strip()
    isin = str(row.get("isin", "")).strip()
    ident = cusip if cusip and cusip.lower() != "nan" else isin
    label = row.get("bond_label") or make_bond_label(row)
    oas = fmt_num(row.get("oas"), 0)
    score = fmt_num(row.get("rv_score"), 0)
    return f"{ident} | {label} | OAS {oas} | score {score}"




def make_curve_note(row: pd.Series) -> str:
    """Return an issuer-curve note only when the curve was actually used.


    Most high-yield issuers do not have enough clean, comparable same-issuer bonds
    to support a reliable curve. For those bonds, the curve component is not used in
    scoring and should not clutter the dashboard note. The raw fields remain in the
    CSV for audit, but the user-facing note is intentionally silent unless the curve
    signal is High or Medium confidence.
    """
    residual = row.get("issuer_curve_residual")
    confidence = str(row.get("issuer_curve_confidence", "Unavailable"))
    if pd.isna(residual) or confidence not in {"High", "Medium"}:
        return ""
    model_n = fmt_num(row.get("issuer_curve_model_bond_count"), 0)
    return f"issuer-curve residual {fmt_num(residual, 0)} bp; curve confidence {confidence}; curve bonds {model_n}"




def make_bond_note(row: pd.Series) -> str:
    if not bool(row.get("normal_rv_eligible", False)):
        status = row.get("core_universe_status", "Review")
        lender_status = row.get("lender_stress_status", "")
        lender_text = f"; lender filter: {lender_status}" if str(lender_status) in {"Stressed Lender Watch", "Distressed Lender Review"} else ""
        return f"Excluded from Core B/BB RV monitor ({status}{lender_text}). Flags: {row.get('flags', 'none')}."


    parts = [
        f"OAS {fmt_num(row.get('oas'), 0)} bp",
        f"own 1Y OAS percentile {fmt_num(row.get('oas_pct_1y'), 0)}%",
        f"issuer 1Y percentile {fmt_num(row.get('issuer_oas_pct_1y'), 0)}%",
        f"shrunk peer residual {fmt_num(row.get('shrunk_peer_residual', row.get('peer_oas_residual')), 0)} bp versus {row.get('peer_group_used', 'peer group')}",
        f"residual-history percentile {fmt_num(row.get('residual_oas_pct_1y'), 0)}%",
        f"residual is {fmt_num(row.get('residual_oas_vs_1y_median'), 0)} bp versus its 1Y residual median",
        f"issuer rating/sector residual percentile {fmt_num(row.get('issuer_residual_pct_1y'), 0)}%",
        f"bond is {fmt_num(row.get('oas_vs_1y_median'), 0)} bp versus its 1Y median",
        f"cohort context {row.get('cohort_context_label', 'Cohort Fair')} ({fmt_num(row.get('cohort_oas_pct_1y'), 0)}% 1Y cohort percentile)",
        f"peer quality {row.get('peer_group_quality', 'n/a')} / stability {row.get('peer_median_stability_label', 'n/a')}",
        f"model support {row.get('model_support', row.get('history_confidence', 'n/a'))}",
    ]
    curve_note = make_curve_note(row)
    if curve_note:
        parts.append(curve_note)
    caution = str(row.get('why_signal_may_be_wrong', '')).strip()
    note = f"{row.get('ticker', '')} screens {row.get('rv_signal', 'No Score')} ({row.get('relative_vs_absolute_label', 'relative signal only')}): " + "; ".join(parts) + "."
    if caution:
        note += f" Why this may be misleading: {caution}."
    return note




def build_issuer_output(issuer_current: pd.DataFrame) -> pd.DataFrame:
    result = issuer_current.copy()
    result["issuer_distance_score"] = distance_score(result["issuer_oas_vs_1y_median"])
    components = pd.DataFrame(index=result.index)
    components["raw_oas_pct"] = result["issuer_oas_pct_1y"]
    components["rating_sector_residual_pct"] = result["issuer_residual_pct_1y"]
    components["distance_score"] = result["issuer_distance_score"]
    # Issuer score for the core monitor is also peer-aware: similar-rating/similar-sector
    # residuals carry more weight than raw issuer spread percentiles.
    weights = pd.Series({"raw_oas_pct": 0.25, "rating_sector_residual_pct": 0.55, "distance_score": 0.20})
    weighted_sum = components.mul(weights, axis=1).sum(axis=1, skipna=True)
    weight_sum = components.notna().mul(weights, axis=1).sum(axis=1)
    result["issuer_score"] = np.where(weight_sum > 0, weighted_sum / weight_sum, np.nan)
    result.loc[result["issuer_distressed_event"].fillna(False), "issuer_score"] = np.nan
    result["issuer_signal"] = result.apply(lambda r: issuer_signal_from_score(r.get("issuer_score"), r.get("issuer_distressed_event", False)), axis=1)
    result["issuer_relative_vs_absolute_label"] = result.apply(lambda r: relative_vs_absolute_label(r.get("issuer_signal"), "Cohort Fair"), axis=1)
    result["issuer_note"] = result.apply(
        lambda r: (
            f"{r['ticker']} screens {r['issuer_signal']}: issuer median OAS {fmt_num(r['issuer_median_oas'], 0)} bp; "
            f"1Y OAS percentile {fmt_num(r.get('issuer_oas_pct_1y'), 0)}%; "
            f"rating/sector residual percentile {fmt_num(r.get('issuer_residual_pct_1y'), 0)}%; "
            f"issuer is {fmt_num(r.get('issuer_oas_vs_1y_median'), 0)} bp versus its 1Y median; "
            f"peer residual {fmt_num(r.get('issuer_peer_oas_residual'), 0)} bp versus {r.get('issuer_peer_group_used', 'issuer peers')}; "
            f"bond count {int(r['issuer_bond_count']) if pd.notna(r['issuer_bond_count']) else 0}; model support {r.get('issuer_confidence', 'n/a')}; flags {r.get('issuer_flags', 'none')}."
        ),
        axis=1,
    )
    return result




def add_issuer_bond_summaries(issuer_output: pd.DataFrame, bond_latest: pd.DataFrame) -> pd.DataFrame:
    """Attach the specific CUSIPs driving each issuer's cheap/rich signal."""
    result = issuer_output.copy()
    if result.empty or bond_latest.empty or "ticker" not in bond_latest.columns:
        result["cheap_bond_cusips"] = ""
        result["rich_bond_cusips"] = ""
        return result


    bonds = bond_latest.loc[bond_latest["normal_rv_eligible"] & bond_latest["rv_score"].notna()].copy()
    if "bond_label" not in bonds.columns:
        bonds["bond_label"] = bonds.apply(make_bond_label, axis=1)


    cheap_map: Dict[str, str] = {}
    rich_map: Dict[str, str] = {}
    for ticker, group in bonds.groupby("ticker", observed=True):
        cheap_rows = group.sort_values("rv_score", ascending=False).head(3)
        rich_rows = group.sort_values("rv_score", ascending=True).head(3)
        cheap_map[ticker] = "; ".join(make_bond_snippet(r) for _, r in cheap_rows.iterrows())
        rich_map[ticker] = "; ".join(make_bond_snippet(r) for _, r in rich_rows.iterrows())


    result["cheap_bond_cusips"] = result["ticker"].map(cheap_map).fillna("")
    result["rich_bond_cusips"] = result["ticker"].map(rich_map).fillna("")
    return result




def build_switch_candidates(bond_latest: pd.DataFrame) -> pd.DataFrame:
    log("Building same-issuer switch candidates...")
    rows: List[Dict[str, object]] = []
    use = bond_latest.loc[
        bond_latest["normal_rv_eligible"]
        & bond_latest["issuer_curve_residual"].notna()
        & bond_latest.get("issuer_curve_confidence", pd.Series("Unavailable", index=bond_latest.index)).isin(["High", "Medium"])
        & bond_latest["ticker"].notna()
    ].copy()
    for ticker, group in use.groupby("ticker", observed=True):
        if len(group) < 3:
            continue
        cheap = group.sort_values("issuer_curve_residual", ascending=False).head(1).iloc[0]
        rich = group.sort_values("issuer_curve_residual", ascending=True).head(1).iloc[0]
        pickup = cheap["issuer_curve_residual"] - rich["issuer_curve_residual"]
        if pd.isna(pickup) or pickup < 40:
            continue
        rows.append(
            {
                "as_of_date": cheap["as_of_date"],
                "ticker": ticker,
                "buy_description": cheap.get("description"),
                "buy_cusip": cheap.get("cusip"),
                "buy_isin": cheap.get("isin"),
                "buy_price": cheap.get("price"),
                "buy_oas": cheap.get("oas"),
                "buy_ytw": cheap.get("ytw"),
                "buy_years_to_worst": cheap.get("years_to_worst"),
                "buy_curve_residual": cheap.get("issuer_curve_residual"),
                "sell_description": rich.get("description"),
                "sell_cusip": rich.get("cusip"),
                "sell_isin": rich.get("isin"),
                "sell_price": rich.get("price"),
                "sell_oas": rich.get("oas"),
                "sell_ytw": rich.get("ytw"),
                "sell_years_to_worst": rich.get("years_to_worst"),
                "sell_curve_residual": rich.get("issuer_curve_residual"),
                "gross_residual_pickup_bp": pickup,
                "note": f"Buy bond is {fmt_num(cheap.get('issuer_curve_residual'), 0)} bp wide to issuer curve; sell bond is {fmt_num(rich.get('issuer_curve_residual'), 0)} bp rich. Gross curve-residual pickup {fmt_num(pickup, 0)} bp.",
            }
        )
    return pd.DataFrame(rows).sort_values("gross_residual_pickup_bp", ascending=False) if rows else pd.DataFrame()




def review_priority_score(df: pd.DataFrame) -> pd.Series:
    """Prioritize the HY / Distressed Review page.


    Higher score means more urgent review. This is not a rich/cheap score; it is a
    risk / attention score driven by OAS, price stress, rating weakness, and recent
    spread movement when available.
    """
    oas = pd.to_numeric(df.get("oas", pd.Series(np.nan, index=df.index)), errors="coerce").clip(lower=0, upper=5000)
    price = pd.to_numeric(df.get("price", pd.Series(np.nan, index=df.index)), errors="coerce")
    rating_s = pd.to_numeric(df.get("rating_score", pd.Series(np.nan, index=df.index)), errors="coerce").fillna(6)
    one_m = pd.to_numeric(df.get("oas_1m_chg", pd.Series(np.nan, index=df.index)), errors="coerce").fillna(0).clip(lower=-500, upper=1000)
    one_w = pd.to_numeric(df.get("oas_1w_chg", pd.Series(np.nan, index=df.index)), errors="coerce").fillna(0).clip(lower=-250, upper=500)
    price_stress = (100 - price).clip(lower=0, upper=100).fillna(0)
    return (oas / 25.0) + (price_stress * 1.5) + (rating_s * 8.0) + (one_m.clip(lower=0) / 5.0) + (one_w.clip(lower=0) / 4.0)




def issuer_review_summary(latest_rows: pd.DataFrame) -> pd.DataFrame:
    """Issuer-level summary for the non-core / stressed / distressed review page."""
    use = latest_rows.loc[latest_rows["ticker"].notna() & (latest_rows["ticker"] != "UNKNOWN")].copy()
    if use.empty:
        return pd.DataFrame(columns=ISSUER_OUTPUT_COLUMNS)
    use["review_priority_score"] = review_priority_score(use)
    use["core_flag"] = use["core_universe_status"].eq("Core B/BB")
    use["distress_flag"] = use["core_universe_status"].isin(["Distressed/Event", "CCC / Distressed Review"])
    use["stressed_flag"] = use["core_universe_status"].eq("Stressed Watch")
    use["stressed_lender_flag"] = use["core_universe_status"].eq("Stressed Lender Watch")
    use["distressed_lender_flag"] = use["core_universe_status"].eq("Distressed Lender Review")
    grouped = use.groupby("ticker", observed=True).agg(
        as_of_date=("as_of_date", "max"),
        issuer_bond_count=("bond_key", "nunique"),
        issuer_index_weight=("index_weight", "sum"),
        issuer_median_oas=("oas", "median"),
        issuer_weighted_oas=("oas", "median"),
        issuer_median_price=("price", "median"),
        issuer_median_ytw=("ytw", "median"),
        issuer_rating_bucket=("rating_bucket", mode_or_unknown),
        issuer_core_universe_status=("core_universe_status", mode_or_unknown),
        issuer_sector_l3=("sector_l3", mode_or_unknown),
        issuer_target_sector_group=("target_sector_group", mode_or_unknown),
        issuer_review_priority_score=("review_priority_score", "max"),
        distressed_bond_count=("distress_flag", "sum"),
        stressed_watch_bond_count=("stressed_flag", "sum"),
        stressed_lender_bond_count=("stressed_lender_flag", "sum"),
        distressed_lender_bond_count=("distressed_lender_flag", "sum"),
        core_bond_count=("core_flag", "sum"),
    ).reset_index()
    grouped["issuer_review_bucket"] = np.select(
        [
            grouped["distressed_lender_bond_count"].fillna(0).gt(0),
            grouped["stressed_lender_bond_count"].fillna(0).gt(0),
            grouped["distressed_bond_count"].fillna(0).gt(0),
            grouped["stressed_watch_bond_count"].fillna(0).gt(0),
            grouped["core_bond_count"].fillna(0).gt(0),
        ],
        ["Distressed Lender Review", "Stressed Lender Watch", "Distressed / CCC Review", "Stressed Watch", "Mixed Core / Review"],
        default="Other HY Review",
    )
    grouped["issuer_distressed_event"] = grouped["issuer_review_bucket"].isin(["Distressed / CCC Review", "Stressed Watch", "Stressed Lender Watch", "Distressed Lender Review"])
    grouped["issuer_flags"] = grouped["issuer_review_bucket"]
    grouped["issuer_confidence"] = "Review"
    grouped["issuer_signal"] = grouped["issuer_review_bucket"]
    grouped["issuer_score"] = np.nan
    grouped["issuer_note"] = grouped.apply(
        lambda r: (
            f"{r['ticker']} is in {r['issuer_review_bucket']}: median OAS {fmt_num(r.get('issuer_median_oas'),0)} bp; "
            f"median price {fmt_num(r.get('issuer_median_price'),2)}; rating bucket {r.get('issuer_rating_bucket','n/a')}; "
            f"review bonds {int(r.get('issuer_bond_count') or 0)}."
        ), axis=1
    )
    grouped["cheap_bond_cusips"] = ""
    grouped["rich_bond_cusips"] = ""
    return grouped






def write_dashboard_history_files(
    output_dir: Path,
    history: pd.DataFrame,
    core_bonds: pd.DataFrame,
    review_bonds: pd.DataFrame,
    as_of_date: pd.Timestamp,
    lookback_days: int = 395,
) -> None:
    """Write chart-ready history files for dashboard tear sheets.


    These files are intentionally limited to the latest monitor constituents and a
    roughly one-year lookback so the static dashboard can load them quickly. They
    power the Price/OAS Momentum Strip and the OAS-vs-peer history chart shown
    when a user clicks a bond.
    """
    log("Writing dashboard visual history files...")
    if history.empty:
        return
    start_date = as_of_date - pd.Timedelta(days=lookback_days)


    chart_cols = [
        "as_of_date",
        "bond_key",
        "ticker",
        "description",
        "cusip",
        "isin",
        "bond_label",
        "rating",
        "rating_bucket",
        "sector_l3",
        "sector_l4",
        "maturity_bucket",
        "years_to_worst",
        "price",
        "prev_price",
        "oas",
        "ytw",
        "index_weight",
        "hist_peer_group_used",
        "hist_peer_count",
        "hist_peer_median_oas",
        "hist_peer_oas_residual",
    ]


    def safe_chart_frame(keys: pd.Series) -> pd.DataFrame:
        keys = keys.dropna().astype(str).unique()
        if len(keys) == 0:
            return pd.DataFrame(columns=chart_cols)
        hist = history.loc[
            history["bond_key"].astype(str).isin(keys)
            & history["as_of_date"].ge(start_date)
            & history["as_of_date"].le(as_of_date)
        ].copy()
        if hist.empty:
            return pd.DataFrame(columns=chart_cols)
        # Keep only the columns the browser needs. If an older cache lacks a field,
        # create it as blank so the dashboard can still render.
        for col in chart_cols:
            if col not in hist.columns:
                hist[col] = np.nan
        if "bond_label" not in hist.columns or hist["bond_label"].isna().all():
            hist["bond_label"] = hist.apply(make_bond_label, axis=1)
        hist["price_change_1d_from_prev"] = hist["price"] - hist["prev_price"]
        hist = hist.sort_values(["bond_key", "as_of_date"])
        return hist[chart_cols + ["price_change_1d_from_prev"]]


    core_hist = safe_chart_frame(core_bonds.get("bond_key", pd.Series(dtype=object)))
    review_hist = safe_chart_frame(review_bonds.get("bond_key", pd.Series(dtype=object)))


    core_hist.to_csv(output_dir / "bond_history_core_dashboard.csv", index=False)
    review_hist.to_csv(output_dir / "bond_history_other_dashboard.csv", index=False)


def core_exclusion_reason(df: pd.DataFrame) -> pd.Series:
    """Explain, per bond, why a latest-date row is not in the Core B/BB output.

    Written for the core_exclusion_report.csv diagnostic so missing tickers can
    be traced to a specific rule instead of silently disappearing.
    """
    rating = df.get("rating", pd.Series("", index=df.index)).astype(str).str.upper().str.strip()
    rating_bucket_col = df.get("rating_bucket", pd.Series("OTHER", index=df.index)).astype(str).str.upper()
    sector_l2 = df.get("sector_l2", pd.Series("", index=df.index)).astype(str).str.upper()
    status = df.get("core_universe_status", pd.Series("", index=df.index)).astype(str)
    lender_status = df.get("lender_stress_status", pd.Series("Non-Lender", index=df.index)).astype(str)
    price = pd.to_numeric(df.get("price", pd.Series(np.nan, index=df.index)), errors="coerce")
    oas = pd.to_numeric(df.get("oas", pd.Series(np.nan, index=df.index)), errors="coerce")
    ytw = pd.to_numeric(df.get("ytw", pd.Series(np.nan, index=df.index)), errors="coerce")
    years = pd.to_numeric(df.get("years_to_worst", pd.Series(np.nan, index=df.index)), errors="coerce")

    reasons = pd.Series("", index=df.index, dtype=object)

    def add(cond: pd.Series, label: str) -> None:
        nonlocal reasons
        reasons = reasons.mask(cond & reasons.ne(""), reasons + "; " + label)
        reasons = reasons.mask(cond & reasons.eq(""), label)

    add((sector_l2 == "CASH") | (rating == "CASH"), "cash row")
    add(price.isna(), "missing price")
    add(oas.isna(), "missing OAS")
    add(~rating.isin(scored_ratings_exact()), "rating '" + rating + "' not scoreable")
    add(rating.isin(scored_ratings_exact()) & ~rating_bucket_col.isin(scored_rating_buckets()), "rating bucket " + rating_bucket_col)
    add(price.lt(universe_min_price()), f"price < {universe_min_price():g}")
    add(oas.ge(universe_max_oas()), f"OAS >= {universe_max_oas():g}")
    add(oas.le(0), "OAS <= 0")
    add(ytw.fillna(0).ge(50), "YTW >= 50")
    add(years.fillna(0).lt(0.25), "years to worst < 0.25")
    excluded_lender = (
        ["Stressed Lender Watch", "Distressed Lender Review"]
        if CORE_UNIVERSE_MODE == "bb_b_only"
        else ["Distressed Lender Review"]
    )
    add(lender_status.isin(excluded_lender), "lender stress: " + lender_status)
    add(reasons.eq("") & status.ne("Core B/BB"), "status: " + status)
    return reasons.replace("", "unknown (check normal_rv_eligible)")


def write_outputs(
    output_dir: Path,
    history: pd.DataFrame,
    bond_latest: pd.DataFrame,
    issuer_output: pd.DataFrame,
    switches: pd.DataFrame,
    parse_report: List[Dict[str, object]],
    used_cache: bool,
) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    log(f"Writing outputs to: {output_dir}")


    bond_latest = bond_latest.copy()
    bond_latest["review_bucket"] = bond_latest["core_universe_status"]
    bond_latest["review_priority_score"] = review_priority_score(bond_latest)


    exact_core_rating = bond_latest["rating"].astype(str).str.upper().str.strip().isin(scored_ratings_exact())
    normal_bonds = bond_latest.loc[bond_latest["normal_rv_eligible"] & exact_core_rating & bond_latest["core_universe_status"].eq("Core B/BB")].copy()
    review_bonds = bond_latest.loc[~(bond_latest.index.isin(normal_bonds.index))].copy()


    def safe_cols(df: pd.DataFrame, cols: Sequence[str]) -> pd.DataFrame:
        out = df.copy()
        for col in cols:
            if col not in out.columns:
                out[col] = np.nan
        return out[list(cols)]


    bond_out = safe_cols(normal_bonds.sort_values("rv_score", ascending=False), OUTPUT_BOND_COLUMNS)
    review_bond_out = safe_cols(review_bonds.sort_values("review_priority_score", ascending=False), OUTPUT_BOND_COLUMNS)


    issuer_core_rating = issuer_output.get("issuer_rating_bucket", pd.Series("", index=issuer_output.index)).astype(str).str.upper().isin(scored_rating_buckets())
    issuer_distress_flag = issuer_output.get("issuer_distressed_event", pd.Series(False, index=issuer_output.index)).fillna(False)
    issuer_normal = issuer_output.loc[(~issuer_distress_flag) & issuer_core_rating].copy()
    issuer_out = safe_cols(issuer_normal.sort_values("issuer_score", ascending=False), ISSUER_OUTPUT_COLUMNS)


    issuer_review = issuer_review_summary(review_bonds)
    issuer_review_out = safe_cols(issuer_review.sort_values("issuer_review_priority_score", ascending=False), ISSUER_OUTPUT_COLUMNS)


    # Backward-compatible core outputs plus explicit split-monitor outputs.
    bond_out.to_csv(output_dir / "bond_rv_latest.csv", index=False)
    issuer_out.to_csv(output_dir / "issuer_rv_latest.csv", index=False)
    bond_out.to_csv(output_dir / "bond_rv_core_latest.csv", index=False)
    issuer_out.to_csv(output_dir / "issuer_rv_core_latest.csv", index=False)
    review_bond_out.to_csv(output_dir / "bond_rv_other_latest.csv", index=False)

    # Diagnostic: every latest-date bond excluded from core, with the exact reason.
    exclusion = review_bonds.copy()
    exclusion["exclusion_reason"] = core_exclusion_reason(exclusion)
    exclusion_cols = [
        c for c in [
            "as_of_date", "ticker", "description", "cusip", "isin",
            "rating_raw", "rating", "rating_bucket", "core_universe_status",
            "lender_stress_status", "price", "oas", "ytw", "years_to_worst",
            "flags", "exclusion_reason",
        ] if c in exclusion.columns
    ]
    exclusion[exclusion_cols].sort_values(["ticker", "oas"], na_position="last").to_csv(
        output_dir / "core_exclusion_report.csv", index=False
    )

    core_tickers = set(normal_bonds.get("ticker", pd.Series(dtype=object)).astype(str))
    missing = (
        exclusion.loc[~exclusion["ticker"].astype(str).isin(core_tickers)]
        .groupby("ticker", observed=True)["exclusion_reason"]
        .agg(lambda v: v.value_counts().index[0])
    )
    if len(missing):
        log(f"Tickers in latest H0A0 with zero Core B/BB bonds: {len(missing):,} (see core_exclusion_report.csv)")
        for tk, reason in missing.items():
            log(f"  {tk}: {reason}")
    issuer_review_out.to_csv(output_dir / "issuer_rv_other_latest.csv", index=False)
    review_bond_out.to_csv(output_dir / "distressed_event_latest.csv", index=False)
    issuer_review_out.to_csv(output_dir / "issuer_distressed_event_latest.csv", index=False)
    switches.to_csv(output_dir / "issuer_switch_candidates.csv", index=False)


    latest_dt = pd.Timestamp(bond_latest["as_of_date"].max()) if not bond_latest.empty else pd.NaT
    if not pd.isna(latest_dt):
        write_dashboard_history_files(output_dir, history, bond_out, review_bond_out, latest_dt)


    report_lines = []
    report_lines.append("H0A0 Historical Relative Value Data Quality Report")
    report_lines.append(f"Model version: {MODEL_VERSION}")
    report_lines.append("RV score weights: 40% historical peer-adjusted residual, 20% Level-4 rating-notch cohort context, 15% current shrunk peer residual, 15% issuer context, 10% own OAS history. Momentum and issuer-curve overlay are excluded from the RV score.")
    report_lines.append(f"Generated: {datetime.now()}")
    report_lines.append(f"Used cache: {used_cache}")
    report_lines.append(f"Latest date: {bond_latest['as_of_date'].max().date() if not bond_latest.empty else 'n/a'}")
    report_lines.append(f"Latest-date bonds: {len(bond_latest):,}")
    report_lines.append(f"Core B/BB Performing RV bonds: {len(normal_bonds):,}")
    report_lines.append(f"HY / Distressed Review bonds: {len(review_bonds):,}")
    report_lines.append("Per-bond exclusion reasons: core_exclusion_report.csv")
    if "lender_stress_status" in review_bonds.columns:
        lender_counts = review_bonds["lender_stress_status"].value_counts(dropna=False)
        report_lines.append("Lender stress filter distribution in review universe:")
        for label, count in lender_counts.items():
            if str(label) != "Non-Lender":
                report_lines.append(f"  {label}: {count:,}")
    report_lines.append("Core bond output rating distribution:")
    if not normal_bonds.empty and "rating" in normal_bonds.columns:
        for rating_label, count in normal_bonds["rating"].astype(str).str.upper().str.strip().value_counts().sort_index().items():
            report_lines.append(f"  {rating_label}: {count:,}")
    if not normal_bonds.empty and "cohort_context_label" in normal_bonds.columns:
        report_lines.append("Core bond cohort context distribution:")
        for label, count in normal_bonds["cohort_context_label"].fillna("Cohort Fair").value_counts().items():
            report_lines.append(f"  {label}: {count:,}")
    if not normal_bonds.empty and "peer_group_quality" in normal_bonds.columns:
        report_lines.append("Core bond peer group quality distribution:")
        for label, count in normal_bonds["peer_group_quality"].fillna("Unknown").value_counts().items():
            report_lines.append(f"  {label}: {count:,}")
    report_lines.append("HY / Distressed Review status distribution:")
    if "core_universe_status" in review_bonds.columns:
        for label, count in review_bonds["core_universe_status"].value_counts(dropna=False).items():
            report_lines.append(f"  {label}: {count:,}")
    report_lines.append(f"Core B/BB RV issuers: {issuer_out['ticker'].nunique() if not issuer_out.empty else 0:,}")
    report_lines.append(f"HY / Distressed Review issuers: {issuer_review_out['ticker'].nunique() if not issuer_review_out.empty else 0:,}")
    report_lines.append(f"Switch candidates: {len(switches):,}")
    report_lines.append("Dashboard visual history files: bond_history_core_dashboard.csv, bond_history_other_dashboard.csv")
    if "issuer_curve_residual" in bond_latest.columns:
        curve_available = bond_latest["issuer_curve_residual"].notna().sum()
        curve_used = (
            bond_latest["issuer_curve_residual"].notna()
            & bond_latest.get("issuer_curve_confidence", pd.Series("Unavailable", index=bond_latest.index)).isin(["High", "Medium"])
        ).sum()
        report_lines.append(f"Issuer curve residual available: {curve_available:,} / {len(bond_latest):,}")
        report_lines.append(f"Issuer curve residual available as diagnostic/note: {curve_used:,} / {len(bond_latest):,}")
    if parse_report:
        errors = [r for r in parse_report if r.get("error")]
        report_lines.append(f"Parsed files: {len(parse_report):,}")
        report_lines.append(f"Parse errors: {len(errors):,}")
        for err in errors[:25]:
            report_lines.append(f"  {err.get('file')}: {err.get('error')}")
    (output_dir / "data_quality_report.txt").write_text("\n".join(report_lines), encoding="utf-8")


    xlsx_path = output_dir / "h0a0_relative_value_report.xlsx"
    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
        bond_out.head(150).to_excel(writer, sheet_name="Core BB B Cheap", index=False)
        bond_out.sort_values("rv_score", ascending=True).head(150).to_excel(writer, sheet_name="Core BB B Rich", index=False)
        bond_out.sort_values(["model_support", "rv_score"], ascending=[False, False]).head(150).to_excel(writer, sheet_name="Signal Quality", index=False)
        issuer_out.head(150).to_excel(writer, sheet_name="Core Issuer RV", index=False)
        review_bond_out.head(200).to_excel(writer, sheet_name="HY Distressed Review", index=False)
        issuer_review_out.head(150).to_excel(writer, sheet_name="Review Issuers", index=False)
        switches.head(100).to_excel(writer, sheet_name="Issuer Switches", index=False)


def run_pipeline(
    history_dir: Path,
    output_dir: Path,
    as_of_date: Optional[str] = None,
    rebuild_cache: bool = False,
    max_files: Optional[int] = None,
    history_lookback_months: int = DEFAULT_HISTORY_LOOKBACK_MONTHS,
) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    log(f"Running {MODEL_VERSION}")
    history, parse_report, used_cache = load_or_parse_history(
        history_dir,
        output_dir,
        rebuild_cache=rebuild_cache,
        max_files=max_files,
        as_of_date=as_of_date,
        history_lookback_months=history_lookback_months,
    )


    dt = latest_date(history, as_of_date)
    log(f"Using as-of date: {dt.date()}")
    if not history.empty:
        log(
            f"RV score history window loaded: {history['as_of_date'].min().date()} to "
            f"{history['as_of_date'].max().date()} ({history['as_of_date'].nunique():,} date(s))."
        )
    history = add_historical_peer_residuals(history)
    latest_rows = history.loc[history["as_of_date"] == dt].copy()
    if latest_rows.empty:
        raise RuntimeError(f"No latest rows found for {dt.date()}")


    latest_rows = add_security_percentiles(history, latest_rows)
    latest_rows = add_peer_metrics(history, latest_rows)
    latest_rows = add_cohort_context_and_peer_quality(history, latest_rows)
    latest_rows = add_residual_history_percentiles(history, latest_rows)
    latest_rows = add_issuer_curve_residuals(latest_rows)


    issuer_daily = compute_issuer_daily(history)
    issuer_current = add_issuer_percentiles(issuer_daily, dt, latest_rows)
    issuer_output = build_issuer_output(issuer_current)


    bond_latest = score_latest(latest_rows, issuer_current)
    issuer_output = add_issuer_bond_summaries(issuer_output, bond_latest)
    switches = build_switch_candidates(bond_latest)
    write_outputs(output_dir, history, bond_latest, issuer_output, switches, parse_report, used_cache)


    log("H0A0 Relative Value finished successfully.")
    log(f"Open: {output_dir / 'h0a0_relative_value_report.xlsx'}")




def parse_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Standalone H0A0 historical spread relative value engine.")
    parser.add_argument("--history-dir", default=DEFAULT_HISTORY_DIR, help="Folder containing historical ICE_H0A0 files.")
    parser.add_argument("--output-dir", default=DEFAULT_OUTPUT_DIR, help="Output folder. Defaults to ./output.")
    parser.add_argument("--as-of-date", default=None, help="Optional as-of date in YYYY-MM-DD format. Defaults to latest available.")
    parser.add_argument("--rebuild-cache", action="store_true", help="Force reparsing source files and rebuilding cache.")
    parser.add_argument("--max-files", type=int, default=None, help="Test mode: use only the latest N files.")
    parser.add_argument(
        "--core-universe",
        choices=["all_rated", "bb_b_only"],
        default="all_rated",
        help=(
            "Which bonds get scored into the core output. 'all_rated' (default) scores every "
            "performing rated bond including BBB crossover and CCC notches. 'bb_b_only' restores "
            "the pre-v2.1 BB1-B3 restriction."
        ),
    )
    parser.add_argument(
        "--min-price",
        type=float,
        default=ALL_RATED_MIN_PRICE,
        help=f"Dollar price floor for the scored universe under all_rated. Default {ALL_RATED_MIN_PRICE:g}.",
    )
    parser.add_argument(
        "--max-oas",
        type=float,
        default=ALL_RATED_MAX_OAS,
        help=f"OAS ceiling (bp) for the scored universe under all_rated. Default {ALL_RATED_MAX_OAS:g}.",
    )
    parser.add_argument(
        "--history-lookback-months", "--history-months",
        dest="history_lookback_months",
        type=int,
        default=DEFAULT_HISTORY_LOOKBACK_MONTHS,
        help="Number of months of ICE files to parse/use for RV scoring. Default: 12. Use 0 to use all files.",
    )
    return parser.parse_args(argv)




def main(argv: Optional[Sequence[str]] = None) -> int:
    global CORE_UNIVERSE_MODE, ALL_RATED_MIN_PRICE, ALL_RATED_MAX_OAS
    args = parse_args(argv)
    CORE_UNIVERSE_MODE = args.core_universe
    ALL_RATED_MIN_PRICE = float(args.min_price)
    ALL_RATED_MAX_OAS = float(args.max_oas)
    log(
        f"Scored universe: {CORE_UNIVERSE_MODE} "
        f"(price >= {universe_min_price():g}, OAS < {universe_max_oas():g})"
    )
    history_dir = Path(args.history_dir)
    output_dir = Path(args.output_dir)


    if not history_dir.exists():
        log(f"ERROR: history directory does not exist: {history_dir}")
        return 2


    try:
        run_pipeline(
            history_dir=history_dir,
            output_dir=output_dir,
            as_of_date=args.as_of_date,
            rebuild_cache=args.rebuild_cache,
            max_files=args.max_files,
            history_lookback_months=args.history_lookback_months,
        )
    except KeyboardInterrupt:
        log("Stopped by user.")
        return 130
    except Exception as exc:
        log(f"ERROR: {exc}")
        raise
    return 0




if __name__ == "__main__":
    raise SystemExit(main())
