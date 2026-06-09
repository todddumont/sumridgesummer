#!/usr/bin/env python3
"""
H0A0 Historical Relative Value Engine

Standalone analytics script for ICE/BAML H0A0 historical constituent files.

Purpose:
    Identify rich/cheap bonds and issuers using historical OAS, peer residuals,
    and issuer-curve residuals.

Default data path:
    P:\\jmorris\\ICE H0A0 Historical Index Data

Outputs:
    output/bond_rv_latest.csv
    output/issuer_rv_latest.csv
    output/issuer_switch_candidates.csv
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
    "coupon": ["Coupon"],
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
    "rating",
    "rating_bucket",
    "sector_l3",
    "sector_l4",
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
    "peer_median_oas",
    "peer_oas_residual",
    "peer_oas_percentile_today",
    "issuer_curve_residual",
    "history_confidence",
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
    "issuer_sector_l3",
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
            normalized = {clean_col_name(x).lower() for x in row}
            if "cusip" in normalized and "ticker" in normalized:
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


def build_manifest(files: Sequence[Path]) -> Dict[str, object]:
    return {
        "version": 3,
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
) -> Tuple[pd.DataFrame, List[Dict[str, object]], bool]:
    files = list_history_files(history_dir)
    if max_files is not None and max_files > 0:
        files = files[-max_files:]
        log(f"Test mode: using latest {len(files)} file(s).")

    if not files:
        raise FileNotFoundError(f"No supported history files found in {history_dir}")

    cache_dir = output_dir / CACHE_DIR_NAME
    cache_dir.mkdir(parents=True, exist_ok=True)
    cache_file = cache_dir / CACHE_FILE_NAME
    manifest_file = cache_dir / MANIFEST_FILE_NAME
    current_manifest = build_manifest(files)

    if not rebuild_cache and max_files is None and cache_file.exists() and manifest_file.exists():
        try:
            existing_manifest = json.loads(manifest_file.read_text(encoding="utf-8"))
            if manifest_matches(existing_manifest, current_manifest):
                log(f"Loading normalized history cache: {cache_file}")
                history = pd.read_pickle(cache_file)
                log(f"Loaded {len(history):,} normalized row(s) from cache.")
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


def finalize_history(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["as_of_date"] = pd.to_datetime(df["as_of_date"], errors="coerce")
    df = df.loc[df["as_of_date"].notna()].copy()

    df["ticker"] = df["ticker"].fillna("UNKNOWN").astype(str).str.upper().str.strip()
    df["rating"] = df["rating"].fillna("NR").astype(str).str.upper().str.strip()
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


def normal_rv_eligible(df: pd.DataFrame) -> pd.Series:
    rating = df["rating"].astype(str).str.upper()
    sector_l2 = df["sector_l2"].astype(str).str.upper()
    oas = df["oas"]
    price = df["price"]
    ytw = df["ytw"]
    years = df["years_to_worst"]

    return (
        (sector_l2 != "CASH")
        & (rating != "CASH")
        & (~rating.isin(["D", "C", "CC"]))
        & price.notna()
        & oas.notna()
        & (price >= 70)
        & (oas > 0)
        & (oas < 1000)
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
    return base.str.strip(";").replace("", "none")


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


def distance_score(series: pd.Series) -> pd.Series:
    """Rank positive distance as cheap and negative distance as rich on a 0-100 scale."""
    out = pd.Series(np.nan, index=series.index, dtype=float)
    mask = series.notna()
    if mask.sum() > 0:
        out.loc[mask] = series.loc[mask].rank(pct=True) * 100.0
    return out


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
        issuer_sector_l3=("sector_l3", mode_or_unknown),
        issuer_sector_l4=("sector_l4", mode_or_unknown),
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
        ("rating", ["as_of_date", "issuer_rating_bucket"]),
        ("all_normal_issuers", ["as_of_date"]),
    ]
    for group_name, keys in peer_specs:
        stats = grouped.groupby(keys, observed=True)["issuer_median_oas"].agg(issuer_peer_count="count", issuer_peer_median_oas="median").reset_index()
        candidate = grouped[keys].merge(stats, on=keys, how="left")
        min_count = 8 if group_name == "sector_l3_rating" else 15 if group_name == "rating" else 25
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
    all_stats = all_current.groupby("ticker", observed=True).agg(
        all_bond_count=("bond_key", "nunique"),
        all_median_oas=("oas", "median"),
        all_median_price=("price", "median"),
        distressed_bond_count=("distressed_event", "sum"),
    ).reset_index()
    all_stats["issuer_distressed_share"] = np.where(all_stats["all_bond_count"] > 0, all_stats["distressed_bond_count"] / all_stats["all_bond_count"], np.nan)
    current = current.merge(all_stats, on="ticker", how="left")

    flags = pd.Series([""] * len(current), index=current.index, dtype=object)
    def add_flag(mask: pd.Series, label: str) -> None:
        nonlocal flags
        flags = flags.mask(mask.fillna(False), flags + label + ";")

    add_flag(current["all_median_oas"].ge(1500), "issuer_oas_above_1500")
    add_flag(current["all_median_oas"].ge(2500), "issuer_oas_above_2500")
    add_flag(current["all_median_price"].lt(70), "issuer_price_below_70")
    add_flag(current["issuer_distressed_share"].ge(0.50), "majority_distressed_bonds")
    add_flag(current["issuer_median_oas"].ge(1500), "normal_issuer_oas_above_1500")
    current["issuer_flags"] = flags.str.strip(";").replace("", "none")
    current["issuer_distressed_event"] = current["issuer_flags"].ne("none")

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
def add_issuer_curve_residuals(latest_rows: pd.DataFrame) -> pd.DataFrame:
    log("Calculating issuer-curve residuals on latest date...")
    result = latest_rows.copy()
    result["issuer_curve_fair_oas"] = np.nan
    result["issuer_curve_residual"] = np.nan
    result["issuer_curve_bond_count"] = np.nan

    normal = result.loc[
        result["normal_rv_eligible"]
        & result["ticker"].notna()
        & result["oas"].notna()
        & result["years_to_worst"].notna()
    ]

    for ticker, group in normal.groupby("ticker", observed=True):
        if len(group) < 3:
            continue
        x = group["years_to_worst"].astype(float).values
        y = group["oas"].astype(float).values
        w = group["rv_weight"].fillna(1.0).astype(float).values
        valid = np.isfinite(x) & np.isfinite(y) & np.isfinite(w)
        if valid.sum() < 3 or np.nanstd(x[valid]) < 0.05:
            fair = np.repeat(np.nanmedian(y[valid]), valid.sum())
            idx = group.index[valid]
        else:
            try:
                coeff = np.polyfit(x[valid], y[valid], deg=1, w=np.sqrt(np.maximum(w[valid], 0.0001)))
                fair = coeff[0] * x[valid] + coeff[1]
                idx = group.index[valid]
            except Exception:
                fair = np.repeat(np.nanmedian(y[valid]), valid.sum())
                idx = group.index[valid]
        result.loc[idx, "issuer_curve_fair_oas"] = fair
        result.loc[idx, "issuer_curve_residual"] = result.loc[idx, "oas"] - fair
        result.loc[idx, "issuer_curve_bond_count"] = len(group)

    return result


def score_latest(latest_rows: pd.DataFrame, issuer_current: pd.DataFrame) -> pd.DataFrame:
    log("Scoring latest-date bonds...")
    result = latest_rows.copy()
    issuer_cols = [
        "ticker",
        "issuer_oas_pct_1y", "issuer_oas_pct_2y", "issuer_oas_pct_full",
        "issuer_residual_pct_1y", "issuer_residual_pct_2y", "issuer_residual_pct_full",
        "issuer_oas_vs_1y_median", "issuer_oas_vs_1y_p75",
        "issuer_distressed_event", "issuer_confidence",
    ]
    result = result.merge(issuer_current[[c for c in issuer_cols if c in issuer_current.columns]].drop_duplicates("ticker"), on="ticker", how="left")

    normal_mask = result["normal_rv_eligible"]
    result["peer_score"] = np.nan
    result["curve_score"] = np.nan
    result.loc[normal_mask & result["peer_oas_residual"].notna(), "peer_score"] = (
        result.loc[normal_mask & result["peer_oas_residual"].notna(), "peer_oas_residual"].rank(pct=True) * 100.0
    )
    result.loc[normal_mask & result["issuer_curve_residual"].notna(), "curve_score"] = (
        result.loc[normal_mask & result["issuer_curve_residual"].notna(), "issuer_curve_residual"].rank(pct=True) * 100.0
    )

    result["own_distance_score"] = distance_score(result["oas_vs_1y_median"])
    result["issuer_distance_score"] = distance_score(result["issuer_oas_vs_1y_median"])

    components = pd.DataFrame(index=result.index)
    components["own_pct"] = result["oas_pct_1y"]
    components["own_distance"] = result["own_distance_score"]
    components["issuer_pct"] = result["issuer_oas_pct_1y"]
    components["issuer_residual"] = result["issuer_residual_pct_1y"]
    components["issuer_distance"] = result["issuer_distance_score"]
    components["peer"] = result["peer_score"]
    components["curve"] = result["curve_score"]
    weights = pd.Series({
        "own_pct": 0.20,
        "own_distance": 0.10,
        "issuer_pct": 0.15,
        "issuer_residual": 0.15,
        "issuer_distance": 0.10,
        "peer": 0.20,
        "curve": 0.10,
    })

    weighted_sum = components.mul(weights, axis=1).sum(axis=1, skipna=True)
    weight_sum = components.notna().mul(weights, axis=1).sum(axis=1)
    result["rv_score"] = np.where(weight_sum > 0, weighted_sum / weight_sum, np.nan)

    # Do not allow issuer-level distressed/event companies to appear as normal bond RV candidates.
    issuer_distress = result.get("issuer_distressed_event", pd.Series(False, index=result.index)).fillna(False).astype(bool)
    result.loc[issuer_distress, "normal_rv_eligible"] = False
    normal_mask = result["normal_rv_eligible"]
    result.loc[~normal_mask, "rv_score"] = np.nan

    result["history_confidence"] = np.select(
        [
            result["oas_obs_1y"].fillna(0).ge(180) & result["peer_count"].fillna(0).ge(20),
            result["oas_obs_1y"].fillna(0).ge(90) & result["peer_count"].fillna(0).ge(10),
        ],
        ["High", "Medium"],
        default="Low",
    )
    result["rv_signal"] = result["rv_score"].apply(score_to_signal)
    result.loc[~normal_mask, "rv_signal"] = "Distressed/Event Review"
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


def make_bond_note(row: pd.Series) -> str:
    if not bool(row.get("normal_rv_eligible", False)):
        return f"Moved to distressed/event review. Flags: {row.get('flags', 'none')}."
    return (
        f"{row.get('ticker', '')} screens {row.get('rv_signal', 'No Score')}: "
        f"OAS {fmt_num(row.get('oas'), 0)} bp; own 1Y OAS percentile {fmt_num(row.get('oas_pct_1y'), 0)}%; "
        f"issuer 1Y percentile {fmt_num(row.get('issuer_oas_pct_1y'), 0)}%; "
        f"peer residual {fmt_num(row.get('peer_oas_residual'), 0)} bp versus {row.get('peer_group_used', 'peer group')}; "
        f"issuer rating/sector residual percentile {fmt_num(row.get('issuer_residual_pct_1y'), 0)}%; "
        f"bond is {fmt_num(row.get('oas_vs_1y_median'), 0)} bp versus its 1Y median; "
        f"issuer-curve residual {fmt_num(row.get('issuer_curve_residual'), 0)} bp; "
        f"confidence {row.get('history_confidence', 'n/a')}."
    )


def build_issuer_output(issuer_current: pd.DataFrame) -> pd.DataFrame:
    result = issuer_current.copy()
    result["issuer_distance_score"] = distance_score(result["issuer_oas_vs_1y_median"])
    components = pd.DataFrame(index=result.index)
    components["raw_oas_pct"] = result["issuer_oas_pct_1y"]
    components["rating_sector_residual_pct"] = result["issuer_residual_pct_1y"]
    components["distance_score"] = result["issuer_distance_score"]
    weights = pd.Series({"raw_oas_pct": 0.35, "rating_sector_residual_pct": 0.45, "distance_score": 0.20})
    weighted_sum = components.mul(weights, axis=1).sum(axis=1, skipna=True)
    weight_sum = components.notna().mul(weights, axis=1).sum(axis=1)
    result["issuer_score"] = np.where(weight_sum > 0, weighted_sum / weight_sum, np.nan)
    result.loc[result["issuer_distressed_event"].fillna(False), "issuer_score"] = np.nan
    result["issuer_signal"] = result.apply(lambda r: issuer_signal_from_score(r.get("issuer_score"), r.get("issuer_distressed_event", False)), axis=1)
    result["issuer_note"] = result.apply(
        lambda r: (
            f"{r['ticker']} screens {r['issuer_signal']}: issuer median OAS {fmt_num(r['issuer_median_oas'], 0)} bp; "
            f"1Y OAS percentile {fmt_num(r.get('issuer_oas_pct_1y'), 0)}%; "
            f"rating/sector residual percentile {fmt_num(r.get('issuer_residual_pct_1y'), 0)}%; "
            f"issuer is {fmt_num(r.get('issuer_oas_vs_1y_median'), 0)} bp versus its 1Y median; "
            f"peer residual {fmt_num(r.get('issuer_peer_oas_residual'), 0)} bp versus {r.get('issuer_peer_group_used', 'issuer peers')}; "
            f"bond count {int(r['issuer_bond_count']) if pd.notna(r['issuer_bond_count']) else 0}; flags {r.get('issuer_flags', 'none')}."
        ),
        axis=1,
    )
    return result


def build_switch_candidates(bond_latest: pd.DataFrame) -> pd.DataFrame:
    log("Building same-issuer switch candidates...")
    rows: List[Dict[str, object]] = []
    use = bond_latest.loc[
        bond_latest["normal_rv_eligible"]
        & bond_latest["issuer_curve_residual"].notna()
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


def write_outputs(
    output_dir: Path,
    bond_latest: pd.DataFrame,
    issuer_output: pd.DataFrame,
    switches: pd.DataFrame,
    parse_report: List[Dict[str, object]],
    used_cache: bool,
) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    log(f"Writing outputs to: {output_dir}")

    normal_bonds = bond_latest.loc[bond_latest["normal_rv_eligible"]].copy()
    distressed = bond_latest.loc[~bond_latest["normal_rv_eligible"]].copy()

    def safe_cols(df: pd.DataFrame, cols: Sequence[str]) -> pd.DataFrame:
        out = df.copy()
        for col in cols:
            if col not in out.columns:
                out[col] = np.nan
        return out[list(cols)]

    bond_out = safe_cols(normal_bonds.sort_values("rv_score", ascending=False), OUTPUT_BOND_COLUMNS)
    issuer_normal = issuer_output.loc[~issuer_output.get("issuer_distressed_event", pd.Series(False, index=issuer_output.index)).fillna(False)].copy()
    issuer_distressed = issuer_output.loc[issuer_output.get("issuer_distressed_event", pd.Series(False, index=issuer_output.index)).fillna(False)].copy()
    issuer_out = safe_cols(issuer_normal.sort_values("issuer_score", ascending=False), ISSUER_OUTPUT_COLUMNS)
    issuer_distressed_out = safe_cols(issuer_distressed.sort_values("issuer_median_oas", ascending=False), ISSUER_OUTPUT_COLUMNS)
    distressed_out = safe_cols(distressed.sort_values("oas", ascending=False), OUTPUT_BOND_COLUMNS)

    bond_out.to_csv(output_dir / "bond_rv_latest.csv", index=False)
    issuer_out.to_csv(output_dir / "issuer_rv_latest.csv", index=False)
    issuer_distressed_out.to_csv(output_dir / "issuer_distressed_event_latest.csv", index=False)
    switches.to_csv(output_dir / "issuer_switch_candidates.csv", index=False)
    distressed_out.to_csv(output_dir / "distressed_event_latest.csv", index=False)

    report_lines = []
    report_lines.append("H0A0 Historical Relative Value Data Quality Report")
    report_lines.append(f"Generated: {datetime.now()}")
    report_lines.append(f"Used cache: {used_cache}")
    report_lines.append(f"Latest date: {bond_latest['as_of_date'].max().date() if not bond_latest.empty else 'n/a'}")
    report_lines.append(f"Latest-date bonds: {len(bond_latest):,}")
    report_lines.append(f"Normal RV bonds: {len(normal_bonds):,}")
    report_lines.append(f"Distressed/event review bonds: {len(distressed):,}")
    report_lines.append(f"Issuers: {issuer_output['ticker'].nunique() if not issuer_output.empty else 0:,}")
    if "issuer_distressed_event" in issuer_output.columns:
        report_lines.append(f"Normal RV issuers: {(~issuer_output['issuer_distressed_event'].fillna(False)).sum():,}")
        report_lines.append(f"Distressed/event issuers: {issuer_output['issuer_distressed_event'].fillna(False).sum():,}")
    report_lines.append(f"Switch candidates: {len(switches):,}")
    if parse_report:
        errors = [r for r in parse_report if r.get("error")]
        report_lines.append(f"Parsed files: {len(parse_report):,}")
        report_lines.append(f"Parse errors: {len(errors):,}")
        for err in errors[:25]:
            report_lines.append(f"  {err.get('file')}: {err.get('error')}")
    (output_dir / "data_quality_report.txt").write_text("\n".join(report_lines), encoding="utf-8")

    xlsx_path = output_dir / "h0a0_relative_value_report.xlsx"
    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
        bond_out.head(100).to_excel(writer, sheet_name="Top Cheap Bonds", index=False)
        bond_out.sort_values("rv_score", ascending=True).head(100).to_excel(writer, sheet_name="Top Rich Bonds", index=False)
        issuer_out.head(100).to_excel(writer, sheet_name="Issuer RV", index=False)
        issuer_distressed_out.head(100).to_excel(writer, sheet_name="Issuer Distressed", index=False)
        switches.head(100).to_excel(writer, sheet_name="Issuer Switches", index=False)
        distressed_out.head(100).to_excel(writer, sheet_name="Distressed Event", index=False)


def run_pipeline(
    history_dir: Path,
    output_dir: Path,
    as_of_date: Optional[str] = None,
    rebuild_cache: bool = False,
    max_files: Optional[int] = None,
) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    history, parse_report, used_cache = load_or_parse_history(history_dir, output_dir, rebuild_cache=rebuild_cache, max_files=max_files)

    dt = latest_date(history, as_of_date)
    log(f"Using as-of date: {dt.date()}")
    latest_rows = history.loc[history["as_of_date"] == dt].copy()
    if latest_rows.empty:
        raise RuntimeError(f"No latest rows found for {dt.date()}")

    latest_rows = add_security_percentiles(history, latest_rows)
    latest_rows = add_peer_metrics(history, latest_rows)
    latest_rows = add_issuer_curve_residuals(latest_rows)

    issuer_daily = compute_issuer_daily(history)
    issuer_current = add_issuer_percentiles(issuer_daily, dt, latest_rows)
    issuer_output = build_issuer_output(issuer_current)

    bond_latest = score_latest(latest_rows, issuer_current)
    switches = build_switch_candidates(bond_latest)
    write_outputs(output_dir, bond_latest, issuer_output, switches, parse_report, used_cache)

    log("H0A0 Relative Value finished successfully.")
    log(f"Open: {output_dir / 'h0a0_relative_value_report.xlsx'}")


def parse_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Standalone H0A0 historical spread relative value engine.")
    parser.add_argument("--history-dir", default=DEFAULT_HISTORY_DIR, help="Folder containing historical ICE_H0A0 files.")
    parser.add_argument("--output-dir", default=DEFAULT_OUTPUT_DIR, help="Output folder. Defaults to ./output.")
    parser.add_argument("--as-of-date", default=None, help="Optional as-of date in YYYY-MM-DD format. Defaults to latest available.")
    parser.add_argument("--rebuild-cache", action="store_true", help="Force reparsing source files and rebuilding cache.")
    parser.add_argument("--max-files", type=int, default=None, help="Test mode: use only the latest N files.")
    return parser.parse_args(argv)


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = parse_args(argv)
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
