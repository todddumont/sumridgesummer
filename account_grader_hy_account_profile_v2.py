#!/usr/bin/env python3 

r""" 

account_grader_hy_account_improved.py 

 

Standalone HY-only account grader build for SumRidge AI Tool Shed / Rel Val Screener. 

 

Inputs (folders of monthly CSVs; these are the defaults, override with --pnl/--rfq/--hy): 

  - PNL folder:  N:\toddddumont\MARV_sheet_monitor\PNL_DATA 

                   e.g. 26_JUNE_PNL.csv, 26_MAY_PNL.csv, ... 25_JULY_PNL.csv 

  - RFQ folder:  N:\toddddumont\MARV_sheet_monitor\RFQ_DATA 

                   e.g. 26_JUNE_RFQ.csv, 26_MAY_RFQ.csv, ... 25_JULY_RFQ.csv 

  - HY ICE:      P:\jmorris\ICE H0A0 Historical Index Data   (folder of ICE H0A0 csv files) 

 

All CSVs in each folder are read and combined. TRADEDATE cells may be either Excel 

serial numbers or text dates (e.g. 2024-01-15 or 1/15/2024); both are handled. 

 

HY-only mode: IG/C0A0 is intentionally ignored unless --ig is explicitly passed. 

 

Outputs: 

  - account_grader_output/account_grades.csv 

  - account_grader_output/account_bond_preferences.csv 

  - account_grader_output/account_trade_history.csv 

  - account_grader_output/account_alias_map_template.csv 

  - account_grader_output/mpid_alias_candidates.csv 

  - account_grader_output/account_grader.html 

  - account_grader_output/data_quality_report.json 

 

No third party packages required. Uses only the Python standard library. 

""" 

 

from __future__ import annotations 

 

import argparse 

import bisect 

import csv 

import html 

import json 

import math 

import os 

import re 

import sys 

import time 

import zipfile 

from collections import Counter, defaultdict 

from datetime import datetime, timedelta 

from difflib import SequenceMatcher 

import xml.etree.ElementTree as ET 

 

XLSX_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}" 

 

DEFAULT_PNL = r"N:\toddddumont\MARV_sheet_monitor\PNL_DATA" 

DEFAULT_RFQ = r"N:\toddddumont\MARV_sheet_monitor\RFQ_DATA" 

DEFAULT_IG = "" 

DEFAULT_HY = r"P:\jmorris\ICE H0A0 Historical Index Data" 

# Assumption: BUYSELL is from the firm's perspective. 

# If BUYSELL = S, the firm sold and the client bought. 

# If BUYSELL = B, the firm bought and the client sold. 

def client_flow_from_buysell(buysell: str) -> str: 

    s = str(buysell or "").strip().upper() 

    if s == "S": 

        return "Client_Bought" 

    if s == "B": 

        return "Client_Sold" 

    return "Unknown" 

 

 

def clean_header(x) -> str: 

    return str(x or "").strip() 

 

 

def key_header(x) -> str: 

    return clean_header(x).upper() 

 

 

def norm_cusip(x) -> str: 

    s = str(x or "").strip().upper() 

    s = re.sub(r"[^A-Z0-9]", "", s) 

    if s in ("", "NULL", "NONE", "NAN"): 

        return "" 

    return s 

 

 

def cusip_from_isin(isin: str, fallback: str = "") -> str: 

    # Standard US ISIN format: US + 9-char CUSIP + 1 check digit. 

    s = str(isin or "").strip().upper() 

    s = re.sub(r"[^A-Z0-9]", "", s) 

    if len(s) >= 11: 

        c = norm_cusip(s[2:11]) 

        if len(c) == 9: 

            return c 

    f = norm_cusip(fallback) 

    return f if len(f) == 9 else "" 

 

 

def num(x, default=None): 

    if x is None: 

        return default 

    if isinstance(x, (int, float)): 

        if isinstance(x, float) and math.isnan(x): 

            return default 

        return float(x) 

    s = str(x).strip() 

    if s == "" or s.lower() in ("(null)", "null", "none", "nan"): 

        return default 

    # Normalize common export / accounting quirks before parsing: 

    #   unicode minus, thousands separators, currency, stray spaces. 

    s = s.replace("\u2212", "-").replace(",", "").replace("$", "").replace(" ", "") 

    neg = False 

    # Accounting-style negatives: "(1234.56)" means -1234.56. 

    if s.startswith("(") and s.endswith(")"): 

        neg = True 

        s = s[1:-1] 

    # Trailing-minus negatives: "1234.56-" means -1234.56. 

    if s.endswith("-"): 

        neg = True 

        s = s[:-1] 

    if s.startswith("+"): 

        s = s[1:] 

    try: 

        v = float(s) 

    except ValueError: 

        return default 

    return -v if neg else v 

 

 

def excel_date(x) -> str: 

    # TRADEDATE/SETTLEDATE come through as Excel serial numbers in xlsx exports, 

    # but as plain text (e.g. 2024-01-15 or 1/15/2024) in csv exports. If we only 

    # handle the serial case, csv text dates return "", every account ends up with 

    # a blank last_active_date, and the current-activity column collapses to 0 days. 

    if x is None: 

        return "" 

    if isinstance(x, datetime): 

        return x.date().isoformat() 

    s = str(x).strip() 

    if s == "" or s.lower() in ("(null)", "null", "none", "nan"): 

        return "" 

    # Try text date formats first so a value like 2024-01-15 is not misread as a serial. 

    d = parse_iso_date_text(s) 

    if d is not None: 

        return d.isoformat() 

    # Fall back to Excel serial number (numeric date storage). 

    v = num(s) 

    if v is None: 

        return "" 

    try: 

        return (datetime(1899, 12, 30) + timedelta(days=v)).date().isoformat() 

    except Exception: 

        return "" 

 

 

def safe_div(a, b, default=None): 

    try: 

        if b == 0 or b is None: 

            return default 

        return a / b 

    except Exception: 

        return default 

 

 

def clamp(x, lo=1.0, hi=10.0): 

    if x is None: 

        return None 

    return max(lo, min(hi, x)) 

 

 

def round2(x): 

    if x is None: 

        return "" 

    try: 

        return round(float(x), 2) 

    except Exception: 

        return "" 

 

 

def round4(x): 

    if x is None: 

        return "" 

    try: 

        return round(float(x), 4) 

    except Exception: 

        return "" 

 

 

def desk_from_book(book: str) -> str: 

    b = str(book or "").upper().strip() 

    if not b or b in ("(NULL)", "NULL"): 

        return "OTHER" 

    if any(t in b for t in ["MUNI", "MUN", "PUBFIN", "TAXEX", "TAXEXEMPT", "TAX-EX"]): 

        return "MUNI" 

    if any(t in b for t in ["HY", "HIGHYIELD", "HIGH_YIELD", "EBANK", "DISTRESSED", "LEVFIN", "AUTO"]): 

        return "HY" 

    if any(t in b for t in ["IG", "INVESTMENT", "BANKFIN", "FRONT", "CORP", "CREDIT", "FIN"]): 

        return "IG" 

    return "OTHER" 

 

 

def rating_bucket(rating: str) -> str: 

    r = str(rating or "").upper().strip() 

    if not r: 

        return "Unknown" 

    # ICE style examples: BBB2, BB1, B3. 

    if r.startswith("AAA"): 

        return "AAA" 

    if r.startswith("AA"): 

        return "AA" 

    if r.startswith("A") and not r.startswith("AA"): 

        return "A" 

    if r.startswith("BBB"): 

        return "BBB" 

    if r.startswith("BB"): 

        return "BB" 

    if r.startswith("B") and not r.startswith("BB"): 

        return "B" 

    if r.startswith("CCC") or r.startswith("CC") or r.startswith("C") or r.startswith("D"): 

        return "CCC_or_Lower" 

    return r[:6] 

 

 

def tenor_bucket(years) -> str: 

    y = num(years) 

    if y is None: 

        return "Unknown" 

    if y < 2: 

        return "0-2Y" 

    if y < 5: 

        return "2-5Y" 

    if y < 10: 

        return "5-10Y" 

    if y < 15: 

        return "10-15Y" 

    if y < 25: 

        return "15-25Y" 

    return "25Y+" 

 

 

def size_bucket(qty) -> str: 

    q = abs(num(qty, 0) or 0) 

    if q < 100000: 

        return "<100k" 

    if q < 500000: 

        return "100k-500k" 

    if q < 1000000: 

        return "500k-1mm" 

    if q < 5000000: 

        return "1mm-5mm" 

    return "5mm+" 

 

 

def tier_from_grade(g) -> str: 

    if g is None: 

        return "Unscored" 

    if g < 3: 

        return "Bad / avoid unless very defensive" 

    if g < 5: 

        return "Weak / price defensively" 

    if g < 7: 

        return "Neutral" 

    if g < 8.5: 

        return "Good account" 

    return "Priority account" 

 

 

def name_key(s: str) -> str: 

    s = str(s or "").upper().strip() 

    s = re.sub(r"[^A-Z0-9 ]+", " ", s) 

    s = re.sub(r"\s+", " ", s).strip() 

    return s 

 

 

# Execution venues / dealer tags that come through the RFQ BROKERINITIALS 

# (or PNL MPID) field but are NOT real client accounts. Rows attributed to 

# these are dropped during parsing so they never become graded accounts, 

# preference rows, or trade-history rows. 

#   - MarketAxess: MKTX / MARKET AXESS 

#   - Trumid:      TRUMID / TMID 

#   - Tradeweb:    TRADEWEB / TWDS / TW 

#   - DLR:         generic dealer / inter-dealer venue tag 

EXCLUDED_VENUE_KEYS = { 

    "MKTX", "MARKETAXESS", "MARKETAXXESS", 

    "TRUMID", "TMID", 

    "TRADEWEB", "TWDS", "TW", 

    "DLR", "DEALER", 

} 

EXCLUDED_VENUE_SUBSTRINGS = ("MARKETAXESS", "MARKETAXXESS", "TRADEWEB", "TRUMID") 

 

 

def is_excluded_venue(raw) -> bool: 

    """True if a raw broker/MPID name is an execution venue, not a client.""" 

    k = re.sub(r"[^A-Z0-9]", "", str(raw or "").upper()) 

    if not k: 

        return False 

    if k in EXCLUDED_VENUE_KEYS: 

        return True 

    return any(sub in k for sub in EXCLUDED_VENUE_SUBSTRINGS) 

 

 

def alias_match_key(s: str) -> str: 

    s = name_key(s) 

    remove_words = { 

        "LLC", "L", "LP", "INC", "CO", "COMPANY", "CORP", "CORPORATION", "LTD", "LIMITED", 

        "SECURITIES", "SECURITY", "CAPITAL", "MARKETS", "MARKET", "BANK", "TRUST", "ASSET", 

        "MANAGEMENT", "MANAGER", "ADVISORS", "ADVISERS", "INVESTMENT", "INVESTMENTS", "NA", "N", "A" 

    } 

    toks = [t for t in s.split() if t not in remove_words] 

    return " ".join(toks) 

 

 

def load_alias_map(path: str | None) -> dict: 

    aliases = {} 

    if not path or not os.path.exists(path): 

        return aliases 

    with open(path, newline="", encoding="utf-8-sig", errors="replace") as f: 

        rdr = csv.DictReader(f) 

        headers = {h.lower().strip(): h for h in (rdr.fieldnames or [])} 

        raw_col = headers.get("raw_name") or headers.get("entity_raw") or headers.get("raw") or headers.get("mpid") or headers.get("brokerinitials") 

        canon_col = headers.get("canonical_account") or headers.get("canonical") or headers.get("account") 

        if not raw_col or not canon_col: 

            print(f"WARNING: alias map {path} missing raw_name/canonical_account columns. Ignoring.") 

            return aliases 

        for row in rdr: 

            raw = row.get(raw_col, "") 

            canon = row.get(canon_col, "") 

            if raw and canon and canon.strip(): 

                aliases[name_key(raw)] = canon.strip() 

    print(f"Loaded {len(aliases):,} alias mappings from {path}") 

    return aliases 

 

 

def canonical_name(raw: str, aliases: dict) -> str: 

    raw_clean = name_key(raw) 

    if not raw_clean: 

        return "UNKNOWN" 

    return aliases.get(raw_clean, raw_clean) 

 

 

# -------------------------- XLSX streaming reader -------------------------- 

 

def load_shared_strings(z: zipfile.ZipFile) -> list[str]: 

    try: 

        with z.open("xl/sharedStrings.xml") as f: 

            out = [] 

            for event, elem in ET.iterparse(f, events=("end",)): 

                if elem.tag == XLSX_NS + "si": 

                    out.append("".join((t.text or "") for t in elem.iter(XLSX_NS + "t"))) 

                    elem.clear() 

            return out 

    except KeyError: 

        return [] 

 

 

def cell_to_col(ref: str) -> int: 

    m = re.match(r"([A-Z]+)", ref or "") 

    if not m: 

        return 0 

    n = 0 

    for ch in m.group(1): 

        n = n * 26 + ord(ch) - 64 

    return n - 1 

 

 

def cell_value(cell, shared: list[str]): 

    t = cell.attrib.get("t") 

    v = cell.find(XLSX_NS + "v") 

    if v is None: 

        isel = cell.find(XLSX_NS + "is") 

        if isel is not None: 

            return "".join((tx.text or "") for tx in isel.iter(XLSX_NS + "t")) 

        return "" 

    text = v.text or "" 

    if t == "s": 

        try: 

            return shared[int(text)] 

        except Exception: 

            return text 

    if t == "b": 

        return "TRUE" if text == "1" else "FALSE" 

    return text 

 

 

def iter_xlsx_rows(path: str, max_rows: int = 0): 

    # CSV support 

    if path.lower().endswith(".csv"): 

        with open(path, newline="", encoding="utf-8-sig", errors="replace") as f: 

            rdr = csv.reader(f) 

            for i, row in enumerate(rdr): 

                yield row 

                if max_rows and i >= max_rows: 

                    return 

        return 

 

    # XLSX support 

    with zipfile.ZipFile(path) as z: 

        shared = load_shared_strings(z) 

        sheet_path = "xl/worksheets/sheet1.xml" 

 

        with z.open(sheet_path) as f: 

            yielded = 0 

 

            for event, row in ET.iterparse(f, events=("end",)): 

                if row.tag == XLSX_NS + "row": 

                    vals = {} 

                    max_col = -1 

 

                    for c in row.findall(XLSX_NS + "c"): 

                        col = cell_to_col(c.attrib.get("r", "")) 

                        vals[col] = cell_value(c, shared) 

                        if col > max_col: 

                            max_col = col 

 

                    if max_col >= 0: 

                        yield [vals.get(i, "") for i in range(max_col + 1)] 

                        yielded += 1 

 

                        if max_rows and yielded > max_rows: 

                            return 

 

                    row.clear() 

 

 

def header_index(header: list) -> dict: 

    return {key_header(h): i for i, h in enumerate(header)} 

 

 

def get(row: list, idx: dict, col: str, default=""): 

    i = idx.get(key_header(col)) 

    if i is None or i >= len(row): 

        return default 

    v = row[i] 

    if v is None: 

        return default 

    return v 

 

 

def require_cols(idx: dict, cols: list[str], file_label: str): 

    missing = [c for c in cols if key_header(c) not in idx] 

    if missing: 

        raise ValueError(f"{file_label} is missing required columns: {', '.join(missing)}") 

 

 

# -------------------------- ICE reference data -------------------------- 

 

def read_ice_file(path: str, ref_desk: str) -> dict: 

    bonds = {} 

    with open(path, newline="", encoding="utf-8-sig", errors="replace") as f: 

        rdr = csv.reader(f) 

        header = None 

        idx = None 

        for row in rdr: 

            if header is None: 

                if "ISIN number" in row and "Cusip" in row: 

                    header = row 

                    idx = {h: i for i, h in enumerate(header)} 

                continue 

            if not any(row): 

                continue 

 

            def c(col): 

                i = idx.get(col) if idx else None 

                return row[i] if i is not None and i < len(row) else "" 

 

            full_cusip = cusip_from_isin(c("ISIN number"), c("Cusip")) 

            if not full_cusip: 

                continue 

            info = { 

                "cusip": full_cusip, 

                "ref_desk": ref_desk, 

                "index_name": c("Index Name"), 

                "as_of_date": c("As of Date"), 

                "description": c("Description"), 

                "ticker": c("Ticker"), 

                "coupon": c("Par Wtd Coupon") or c("Current Coupon"), 

                "maturity_date": c("Maturity Date"), 

                "years_to_worst": c("Yrs To Worst"), 

                "tenor_bucket": tenor_bucket(c("Yrs To Worst")), 

                "rating": c("Rating"), 

                "rating_bucket": rating_bucket(c("Rating")), 

                "sector_l1": c("ML Industry Lvl 1"), 

                "sector_l2": c("ML Industry Lvl 2"), 

                "sector_l3": c("ML Industry Lvl 3"), 

                "sector_l4": c("ML Industry Lvl 4"), 

                "effective_yield": c("Effective Yield"), 

                "oas": c("OAS"), 

                "spread_to_worst": c("Spread to Worst"), 

                "price": c("Price"), 

            } 

            bonds[full_cusip] = info 

    return bonds 

 

 

def collect_ice_csv_files(path: str) -> list[str]: 

    """Return ICE csv files from either a single file path or a folder path. 

 

    If a folder is provided, all .csv files under that folder are loaded. This is 

    intentional for the public ICE historical index folders; the grader should 

    build a broad CUSIP reference from the full history, not just one day. 

    """ 

    if not path: 

        return [] 

    if os.path.isfile(path): 

        return [path] if path.lower().endswith(".csv") else [] 

    if os.path.isdir(path): 

        files = [] 

        for root, _dirs, names in os.walk(path): 

            for name in names: 

                if name.lower().endswith(".csv"): 

                    files.append(os.path.join(root, name)) 

        # Filename sorting makes ICE_H0A0_20210118 load before ICE_H0A0_20210119, 

        # so newer rows overwrite older rows for duplicate CUSIPs. 

        return sorted(files) 

    return [] 

 

 

def read_ice_source(path: str, ref_desk: str) -> dict: 

    files = collect_ice_csv_files(path) 

    if not files: 

        print(f"WARNING: No {ref_desk} ICE csv files found at: {path}") 

        return {} 

 

    combined = {} 

    print(f"Loading {len(files):,} {ref_desk} ICE csv file(s) from {path}") 

    for i, file_path in enumerate(files, start=1): 

        try: 

            one = read_ice_file(file_path, ref_desk) 

            combined.update(one) 

            if i == 1 or i == len(files) or i % 25 == 0: 

                print(f"  {ref_desk}: {i:,}/{len(files):,} files, {len(combined):,} unique CUSIPs") 

        except Exception as e: 

            print(f"WARNING: Skipping {ref_desk} ICE file due to read error: {file_path} ({e})") 

    return combined 

 

 

def read_bond_reference(ig_path: str, hy_path: str) -> dict: 

    ref = {} 

    if ig_path: 

        if os.path.exists(ig_path): 

            ref.update(read_ice_source(ig_path, "IG")) 

        else: 

            print(f"WARNING: IG path not found: {ig_path}") 

    if hy_path and os.path.exists(hy_path): 

        ref.update(read_ice_source(hy_path, "HY")) 

    else: 

        print(f"WARNING: HY path not found: {hy_path}") 

    return ref 

 

 

# -------------------------- account metrics -------------------------- 

 

def weighted_post_trade_pnl(p1, p5, p20): 

    pairs = [(0.20, num(p1)), (0.50, num(p5)), (0.30, num(p20))] 

    pairs = [(w, v) for w, v in pairs if v is not None] 

    if not pairs: 

        return None 

    total_w = sum(w for w, _ in pairs) 

    return sum(w * v for w, v in pairs) / total_w 

 

 

def new_metric(): 

    return { 

        "raw_names": Counter(), 

        "books": Counter(), 

        "platforms": Counter(), 

        "pnl_trades": 0, 

        "pnl_qty": 0.0, 

        "weighted_pnl_sum": 0.0, 

        "weighted_pnl_count": 0, 

        "pnl_per_mm_sum": 0.0, 

        "pnl_per_mm_count": 0, 

        "positive_trades": 0, 

        "negative_trades": 0, 

        "rfq_trades": 0, 

        "rfq_qty": 0.0, 

        "linked_events": 0, 

        "linked_qty": 0.0, 

        "buy_pnl_trades": 0, 

        "sell_pnl_trades": 0, 

        "buy_rfq_rows": 0, 

        "sell_rfq_rows": 0, 

        "buy_flow_rows": 0, 

        "sell_flow_rows": 0, 

        "buy_flow_qty": 0.0, 

        "sell_flow_qty": 0.0, 

        "buy_interest_weight": 0.0, 

        "sell_interest_weight": 0.0, 

        "last_active_date": "", 

    } 

 

 

def latest_date(a, b): 

    a = str(a or "").strip() 

    b = str(b or "").strip() 

    if not a: 

        return b 

    if not b: 

        return a 

    return max(a, b) 

 

 

def parse_iso_date_text(x): 

    s = str(x or "").strip() 

    if not s: 

        return None 

    if s.lower() in ("(null)", "null", "none", "nan"): 

        return None 

    # Strip a trailing time component ("... 00:00:00", "...T13:45", "... 1:05 PM", trailing Z/offset) 

    # WITHOUT discarding month-name dates that legitimately contain spaces. 

    s = re.sub( 

        r"[ T]\d{1,2}:\d{2}(:\d{2})?(\.\d+)?\s*([AaPp][Mm])?\s*(Z|[+-]\d{2}:?\d{2})?$", 

        "", 

        s, 

    ).strip() 

    fmts = ( 

        "%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y", "%Y/%m/%d", 

        "%m-%d-%Y", "%m-%d-%y", "%Y%m%d", 

        "%d-%b-%Y", "%d-%b-%y", "%d %b %Y", "%d %B %Y", 

        "%b %d %Y", "%b %d, %Y", "%B %d %Y", "%B %d, %Y", 

        "%Y.%m.%d", "%m.%d.%Y", 

    ) 

    for fmt in fmts: 

        try: 

            return datetime.strptime(s, fmt).date() 

        except Exception: 

            pass 

    try: 

        return datetime.fromisoformat(s).date() 

    except Exception: 

        return None 

     

 

def is_recent_direction_trade(date_text: str, months: int = 4) -> bool: 

    """ 

    Return True if the trade falls within the last N months. 

    Used ONLY for buyer/seller direction classification. 

    """ 

    d = parse_iso_date_text(date_text) 

    if d is None: 

        return False 

 

    cutoff = datetime.today().date() - timedelta(days=months * 30) 

    return d >= cutoff 

 

def enrich_account_activity_states(rows: list[dict]) -> list[dict]: 

    """Add historical-quality vs current-activity fields for the Account Grader tab. 

 

    Recency is relative to the newest dated account activity in the output file, so 

    backtests do not become stale simply because the site is opened later. 

    """ 

    dated = [parse_iso_date_text(r.get("last_active_date")) for r in rows] 

    dated = [d for d in dated if d is not None] 

    max_d = max(dated) if dated else None 

 

    for r in rows: 

        pnl_score = float(r.get("pnl_score") or 5.5) 

        adverse = float(r.get("adverse_selection_score") or 5.5) 

        confidence = float(r.get("data_confidence_score") or 5.5) 

        flow = float(r.get("flow_score") or 5.5) 

        bad_rate = r.get("bad_trade_rate") 

        try: 

            bad_rate = float(bad_rate) if bad_rate not in (None, "") else None 

        except Exception: 

            bad_rate = None 

 

        historical_quality_score = clamp(0.50 * pnl_score + 0.30 * adverse + 0.20 * confidence) 

        d = parse_iso_date_text(r.get("last_active_date")) 

        days_since = (max_d - d).days if max_d and d else None 

        if days_since is None: 

            recency_score = 1.0 

            current_activity_label = "No dated activity" 

        elif days_since <= 30: 

            recency_score = 10.0 

            current_activity_label = "Active now" 

        elif days_since <= 90: 

            recency_score = 8.0 

            current_activity_label = "Recently active" 

        elif days_since <= 180: 

            recency_score = 5.0 

            current_activity_label = "Fading activity" 

        else: 

            recency_score = 2.0 

            current_activity_label = "Stale" 

 

        current_activity_score = clamp(0.65 * recency_score + 0.35 * flow) 

 

        if bad_rate is not None and bad_rate >= 0.35: 

            historical_quality_label = "Weak history" 

        elif historical_quality_score >= 7.5: 

            historical_quality_label = "Strong history" 

        elif historical_quality_score >= 6.5: 

            historical_quality_label = "Good history" 

        elif historical_quality_score <= 4.5: 

            historical_quality_label = "Weak history" 

        else: 

            historical_quality_label = "Mixed history" 

 

        active_now = current_activity_label in ("Active now", "Recently active") 

        staleish = current_activity_label in ("Fading activity", "Stale", "No dated activity") 

        good_hist = historical_quality_label in ("Strong history", "Good history") 

        weak_hist = historical_quality_label == "Weak history" 

 

        if active_now and good_hist: 

            state = "Active + high quality" 

            use_note = "Good current call candidate; account is active and historical quality is supportive." 

        elif active_now and weak_hist: 

            state = "Active but price defensively" 

            use_note = "Account is active, but historical P&L/adverse-selection quality is weak. Use tighter sizing or more defensive levels." 

        elif staleish and good_hist: 

            state = "Historically good but stale" 

            use_note = "Quality has been good historically, but activity is stale. Useful as a re-engagement target rather than first call." 

        elif staleish and weak_hist: 

            state = "Stale / low priority" 

            use_note = "Limited current reason to prioritize; history and recency are both weak." 

        else: 

            state = "Monitor" 

            use_note = "Mixed account profile. Use bond-specific buyer/seller evidence before prioritizing." 

 

        r["historical_quality_score"] = historical_quality_score 

        r["historical_quality_label"] = historical_quality_label 

        r["current_activity_score"] = current_activity_score 

        r["current_activity_label"] = current_activity_label 

        r["days_since_last_active"] = "" if days_since is None else days_since 

        r["account_state_label"] = state 

        r["account_use_note"] = use_note 

    return rows 

 

 

def new_pref(): 

    return { 

        "pnl_trades": 0, 

        "rfq_trades": 0, 

        "pnl_qty": 0.0, 

        "rfq_qty": 0.0, 

        "total_qty": 0.0, 

        "interest_weight": 0.0, 

        "pnl_per_mm_sum": 0.0, 

        "pnl_per_mm_count": 0, 

        "last_active_date": "", 

        "example_cusips": set(), 

    } 

 

 

 

def force_single_account_grouping() -> bool: 

    return True 

 

def add_preference(preferences, acct, desk, ref_info, client_flow, qty, source, pnl_per_mm=None, trade_date=""): 

    if not ref_info: 

        return 

    key = ( 

        acct, 

        desk, 

        ref_info.get("ref_desk", ""), 

        client_flow, 

        ref_info.get("ticker", ""), 

        ref_info.get("sector_l2", ""), 

        ref_info.get("sector_l3", ""), 

        ref_info.get("rating_bucket", "Unknown"), 

        ref_info.get("tenor_bucket", "Unknown"), 

    ) 

    p = preferences[key] 

    qabs = abs(num(qty, 0) or 0) 

    if source == "PNL": 

        p["pnl_trades"] += 1 

        p["pnl_qty"] += qabs 

        p["interest_weight"] += 1.0 

    else: 

        p["rfq_trades"] += 1 

        p["rfq_qty"] += qabs 

        p["interest_weight"] += 0.35 

    p["total_qty"] += qabs 

    p["last_active_date"] = latest_date(p.get("last_active_date", ""), trade_date) 

    if pnl_per_mm is not None: 

        p["pnl_per_mm_sum"] += pnl_per_mm 

        p["pnl_per_mm_count"] += 1 

    c = ref_info.get("cusip") 

    if c and len(p["example_cusips"]) < 8: 

        p["example_cusips"].add(c) 

 

 

def get_any(row: list, idx: dict, cols: list[str], default=""): 

    for col in cols: 

        v = get(row, idx, col, "") 

        if str(v or "").strip() not in ("", "(null)", "NULL", "None"): 

            return v 

    return default 

 

 

def clean_text(x) -> str: 

    s = str(x or "").strip() 

    if s.lower() in ("(null)", "null", "none", "nan"): 

        return "" 

    return s 

 

 

def build_transaction_row(source: str, acct: str, raw: str, desk: str, book: str, ref_info: dict, 

                          client_flow: str, buy_sell: str, trade_date: str, settle_date: str, 

                          quantity, price="", platform="", pnladj1="", pnladj5="", pnladj20="", 

                          weighted_pnl=None, pnl_per_mm=None, ticket_number="", cpid="") -> dict: 

    return { 

        "source": source, 

        "canonical_account": acct, 

        "raw_account": raw, 

        "desk": desk, 

        "book": clean_text(book), 

        "platform": clean_text(platform), 

        "ticket_number": clean_text(ticket_number), 

        "cpid": clean_text(cpid), 

        "cusip": ref_info.get("cusip", ""), 

        "ticker": ref_info.get("ticker", ""), 

        "issuer": ref_info.get("issuer", ""), 

        "coupon": ref_info.get("coupon", ""), 

        "maturity_date": ref_info.get("maturity_date", ""), 

        "reference_desk": ref_info.get("ref_desk", ""), 

        "sector_l2": ref_info.get("sector_l2", ""), 

        "sector_l3": ref_info.get("sector_l3", ""), 

        "rating_bucket": ref_info.get("rating_bucket", "Unknown"), 

        "tenor_bucket": ref_info.get("tenor_bucket", "Unknown"), 

        "client_flow": client_flow, 

        "buy_sell": clean_text(buy_sell), 

        "trade_date": trade_date, 

        "settle_date": settle_date, 

        "quantity": abs(num(quantity, 0) or 0), 

        "price": clean_text(price), 

        "pnladj1": clean_text(pnladj1), 

        "pnladj5": clean_text(pnladj5), 

        "pnladj20": clean_text(pnladj20), 

        "weighted_pnl": weighted_pnl, 

        "pnl_per_mm": pnl_per_mm, 

    } 

 

def process_pnl(path, bond_ref, aliases, max_data_rows=0): 

    metrics = defaultdict(new_metric) 

    preferences = defaultdict(new_pref) 

 

    raw_entities = defaultdict( 

        lambda: { 

            "sources": Counter(), 

            "desks": Counter(), 

            "activity": 0, 

        } 

    ) 

 

    quality = { 

        "pnl_rows_read": 0, 

        "pnl_rows_used": 0, 

        "pnl_rows_missing_entity": 0, 

        "pnl_rows_missing_cusip": 0, 

        "pnl_rows_excluded_venue": 0, 

        "pnl_rows_ice_matched": 0, 

        "pnl_date_min": "", 

        "pnl_date_max": "", 

    } 

 

    dates = [] 

    transactions = [] 

 

    rows = iter_xlsx_rows(path) 

 

    try: 

        header = next(rows) 

    except StopIteration: 

        raise ValueError(f"PNL file appears empty: {path}") 

 

    idx = header_index(header) 

 

    require_cols( 

        idx, 

        [ 

            "MPID", 

            "BOOK", 

            "CUSIP", 

            "BUYSELL", 

            "TRADEDATE", 

            "QUANTITY", 

            "PNLADJ1", 

            "PNLADJ5", 

            "PNLADJ20", 

        ], 

        "PNL file", 

    ) 

 

    for rnum, row in enumerate(rows, start=1): 

        if max_data_rows and rnum > max_data_rows: 

            break 

 

        quality["pnl_rows_read"] += 1 

 

        raw = name_key(get(row, idx, "MPID")) 

        if not raw: 

            quality["pnl_rows_missing_entity"] += 1 

            continue 

 

        if is_excluded_venue(raw): 

            quality["pnl_rows_excluded_venue"] = quality.get("pnl_rows_excluded_venue", 0) + 1 

            continue 

 

        book = str(get(row, idx, "BOOK") or "").strip() 

        desk = desk_from_book(book) 

 

        cusip = norm_cusip(get(row, idx, "CUSIP")) 

        if not cusip: 

            quality["pnl_rows_missing_cusip"] += 1 

 

        ref_info = bond_ref.get(cusip) 

 

        if desk == "OTHER" and ref_info: 

            desk = ref_info.get("ref_desk", "OTHER") 

 

        if force_single_account_grouping(): 

            desk = "ALL" 

 

        if ref_info: 

            quality["pnl_rows_ice_matched"] += 1 

 

        acct = canonical_name(raw, aliases) 

 

        qty = num(get(row, idx, "QUANTITY"), 0) or 0 

        qabs = abs(qty) 

 

        m = metrics[(acct, desk)] 

 

        m["raw_names"][raw] += 1 

        m["books"][book] += 1 

        m["pnl_trades"] += 1 

        m["pnl_qty"] += qabs 

 

        quality["pnl_rows_used"] += 1 

 

        raw_entities[raw]["sources"]["PNL_MPID"] += 1 

        raw_entities[raw]["desks"][desk] += 1 

        raw_entities[raw]["activity"] += 1 

 

        d = excel_date(get(row, idx, "TRADEDATE")) 

 

        if quality["pnl_rows_read"] <= 5: 

            print(f"    TRADEDATE sample: {get(row, idx, 'TRADEDATE')!r} -> {d!r}") 

 

        if d: 

            dates.append(d) 

            m["last_active_date"] = latest_date( 

                m.get("last_active_date", ""), 

                d, 

            ) 

 

        wpnl = weighted_post_trade_pnl( 

            get(row, idx, "PNLADJ1"), 

            get(row, idx, "PNLADJ5"), 

            get(row, idx, "PNLADJ20"), 

        ) 

 

        pnl_per_mm = None 

 

        if wpnl is not None: 

            m["weighted_pnl_sum"] += wpnl 

            m["weighted_pnl_count"] += 1 

 

            if wpnl > 0: 

                m["positive_trades"] += 1 

 

            if wpnl < 0: 

                m["negative_trades"] += 1 

 

            size_mm = qabs / 1_000_000 if qabs else None 

 

            if size_mm and size_mm > 0: 

                pnl_per_mm = wpnl / size_mm 

                m["pnl_per_mm_sum"] += pnl_per_mm 

                m["pnl_per_mm_count"] += 1 

 

        client_flow = client_flow_from_buysell( 

            get(row, idx, "BUYSELL") 

        ) 

 

        # Only use recent trades (last 4 months) for buyer/seller direction 

        if is_recent_direction_trade(d): 

 

            if client_flow == "Client_Bought": 

                m["buy_pnl_trades"] += 1 

                m["buy_flow_rows"] += 1 

                m["buy_flow_qty"] += qabs 

                m["buy_interest_weight"] += 1.0 

 

            elif client_flow == "Client_Sold": 

                m["sell_pnl_trades"] += 1 

                m["sell_flow_rows"] += 1 

                m["sell_flow_qty"] += qabs 

                m["sell_interest_weight"] += 1.0 

 

        add_preference( 

            preferences, 

            acct, 

            desk, 

            ref_info, 

            client_flow, 

            qty, 

            "PNL", 

            pnl_per_mm=pnl_per_mm, 

            trade_date=d, 

        ) 

 

        if ref_info: 

            m["linked_events"] += 1 

            m["linked_qty"] += qabs 

 

            transactions.append( 

                build_transaction_row( 

                    source="PNL", 

                    acct=acct, 

                    raw=raw, 

                    desk=desk, 

                    book=book, 

                    ref_info=ref_info, 

                    client_flow=client_flow, 

                    buy_sell=get(row, idx, "BUYSELL"), 

                    trade_date=d, 

                    settle_date=excel_date( 

                        get(row, idx, "SETTLEDATE") 

                    ), 

                    quantity=qty, 

                    price=get(row, idx, "PRICE"), 

                    platform="", 

                    pnladj1=get(row, idx, "PNLADJ1"), 

                    pnladj5=get(row, idx, "PNLADJ5"), 

                    pnladj20=get(row, idx, "PNLADJ20"), 

                    weighted_pnl=wpnl, 

                    pnl_per_mm=pnl_per_mm, 

                    ticket_number=get(row, idx, "TICKETNUMBER"), 

                ) 

            ) 

 

    if dates: 

        quality["pnl_date_min"] = min(dates) 

        quality["pnl_date_max"] = max(dates) 

 

    used = quality.get("pnl_rows_used", 0) 

    print(f"    dates parsed: {len(dates):,} of {used:,} rows" 

          + (f"  (range {quality['pnl_date_min']} to {quality['pnl_date_max']})" if dates 

             else "   <-- ALL EMPTY: TRADEDATE format not recognized (see sample above)")) 

 

    scored = sum(mm["weighted_pnl_count"] for mm in metrics.values()) 

    negative = sum(mm["negative_trades"] for mm in metrics.values()) 

    print(f"    P&L scored: {scored:,} trade(s); losses (weighted P&L < 0): {negative:,}" 

          + ("   <-- 0 losses is suspicious; check PNLADJ number format" if scored and not negative else "")) 

 

    return ( 

        metrics, 

        preferences, 

        raw_entities, 

        quality, 

        transactions, 

    ) 

 

def process_rfq(path, bond_ref, aliases, max_data_rows=0): 

    metrics = defaultdict(new_metric) 

    preferences = defaultdict(new_pref) 

 

    raw_entities = defaultdict( 

        lambda: { 

            "sources": Counter(), 

            "desks": Counter(), 

            "activity": 0, 

        } 

    ) 

 

    quality = { 

        "rfq_rows_read": 0, 

        "rfq_rows_used": 0, 

        "rfq_rows_missing_entity": 0, 

        "rfq_rows_missing_cusip": 0, 

        "rfq_rows_excluded_venue": 0, 

        "rfq_rows_ice_matched": 0, 

        "rfq_date_min": "", 

        "rfq_date_max": "", 

    } 

 

    dates = [] 

    transactions = [] 

 

    rows = iter_xlsx_rows(path) 

 

    try: 

        header = next(rows) 

    except StopIteration: 

        raise ValueError(f"RFQ file appears empty: {path}") 

 

    idx = header_index(header) 

 

    require_cols( 

        idx, 

        [ 

            "PLATFORMID", 

            "BROKERINITIALS", 

            "BOOK", 

            "CUSIP", 

            "BUYSELL", 

            "TRADEDATE", 

            "QUANTITY", 

        ], 

        "RFQ file", 

    ) 

 

    for rnum, row in enumerate(rows, start=1): 

        if max_data_rows and rnum > max_data_rows: 

            break 

 

        quality["rfq_rows_read"] += 1 

 

        raw = name_key(get(row, idx, "BROKERINITIALS")) 

        if not raw: 

            quality["rfq_rows_missing_entity"] += 1 

            continue 

 

        if is_excluded_venue(raw): 

            quality["rfq_rows_excluded_venue"] = quality.get("rfq_rows_excluded_venue", 0) + 1 

            continue 

 

        book = str(get(row, idx, "BOOK") or "").strip() 

        desk = desk_from_book(book) 

 

        cusip = norm_cusip(get(row, idx, "CUSIP")) 

        if not cusip: 

            quality["rfq_rows_missing_cusip"] += 1 

 

        ref_info = bond_ref.get(cusip) 

 

        if desk == "OTHER" and ref_info: 

            desk = ref_info.get("ref_desk", "OTHER") 

 

        if force_single_account_grouping(): 

            desk = "ALL" 

 

        if ref_info: 

            quality["rfq_rows_ice_matched"] += 1 

 

        acct = canonical_name(raw, aliases) 

 

        qty = num(get(row, idx, "QUANTITY"), 0) or 0 

        qabs = abs(qty) 

 

        platform = str(get(row, idx, "PLATFORMID") or "").strip() 

 

        m = metrics[(acct, desk)] 

 

        m["raw_names"][raw] += 1 

        m["books"][book] += 1 

        m["platforms"][platform] += 1 

        m["rfq_trades"] += 1 

        m["rfq_qty"] += qabs 

 

        quality["rfq_rows_used"] += 1 

 

        raw_entities[raw]["sources"]["RFQ_BROKER"] += 1 

        raw_entities[raw]["desks"][desk] += 1 

        raw_entities[raw]["activity"] += 1 

 

        d = excel_date(get(row, idx, "TRADEDATE")) 

 

        if quality["rfq_rows_read"] <= 5: 

            print(f"    TRADEDATE sample: {get(row, idx, 'TRADEDATE')!r} -> {d!r}") 

 

        if d: 

            dates.append(d) 

            m["last_active_date"] = latest_date( 

                m.get("last_active_date", ""), 

                d, 

            ) 

 

        client_flow = client_flow_from_buysell( 

            get(row, idx, "BUYSELL") 

        ) 

 

        # Only use recent trades (last 4 months) for buyer/seller direction 

        if is_recent_direction_trade(d): 

 

            if client_flow == "Client_Bought": 

                m["buy_rfq_rows"] += 1 

                m["buy_flow_rows"] += 1 

                m["buy_flow_qty"] += qabs 

                m["buy_interest_weight"] += 0.35 

 

            elif client_flow == "Client_Sold": 

                m["sell_rfq_rows"] += 1 

                m["sell_flow_rows"] += 1 

                m["sell_flow_qty"] += qabs 

                m["sell_interest_weight"] += 0.35 

 

        add_preference( 

            preferences, 

            acct, 

            desk, 

            ref_info, 

            client_flow, 

            qty, 

            "RFQ", 

            pnl_per_mm=None, 

            trade_date=d, 

        ) 

 

        if ref_info: 

            m["linked_events"] += 1 

            m["linked_qty"] += qabs 

 

            transactions.append( 

                build_transaction_row( 

                    source="RFQ", 

                    acct=acct, 

                    raw=raw, 

                    desk=desk, 

                    book=book, 

                    ref_info=ref_info, 

                    client_flow=client_flow, 

                    buy_sell=get(row, idx, "BUYSELL"), 

                    trade_date=d, 

                    settle_date=excel_date( 

                        get(row, idx, "SETTLEDATE") 

                    ), 

                    quantity=qty, 

                    price=get_any( 

                        row, 

                        idx, 

                        ["PRICE", "P0", "P1", "P5"], 

                    ), 

                    platform=platform, 

                    cpid=get(row, idx, "CPID"), 

                ) 

            ) 

 

    if dates: 

        quality["rfq_date_min"] = min(dates) 

        quality["rfq_date_max"] = max(dates) 

 

    used = quality.get("rfq_rows_used", 0) 

    print(f"    dates parsed: {len(dates):,} of {used:,} rows" 

          + (f"  (range {quality['rfq_date_min']} to {quality['rfq_date_max']})" if dates 

             else "   <-- ALL EMPTY: TRADEDATE format not recognized (see sample above)")) 

 

    return ( 

        metrics, 

        preferences, 

        raw_entities, 

        quality, 

        transactions, 

    ) 

def merge_metrics(a, b): 

    out = defaultdict(new_metric) 

    for src in (a, b): 

        for key, m in src.items(): 

            o = out[key] 

            for ckey in ["raw_names", "books", "platforms"]: 

                o[ckey].update(m[ckey]) 

            for nkey in [ 

                "pnl_trades", "pnl_qty", "weighted_pnl_sum", "weighted_pnl_count", "pnl_per_mm_sum", "pnl_per_mm_count", 

                "positive_trades", "negative_trades", "rfq_trades", "rfq_qty", "linked_events", "linked_qty", 

                "buy_pnl_trades", "sell_pnl_trades", "buy_rfq_rows", "sell_rfq_rows", "buy_flow_rows", "sell_flow_rows", 

                "buy_flow_qty", "sell_flow_qty", "buy_interest_weight", "sell_interest_weight" 

            ]: 

                o[nkey] += m[nkey] 

            o["last_active_date"] = latest_date(o.get("last_active_date", ""), m.get("last_active_date", "")) 

    return out 

 

 

def merge_preferences(a, b): 

    out = defaultdict(new_pref) 

    for src in (a, b): 

        for key, p in src.items(): 

            o = out[key] 

            o["pnl_trades"] += p["pnl_trades"] 

            o["rfq_trades"] += p["rfq_trades"] 

            o["pnl_qty"] += p.get("pnl_qty", 0.0) 

            o["rfq_qty"] += p.get("rfq_qty", 0.0) 

            o["total_qty"] += p["total_qty"] 

            o["interest_weight"] += p["interest_weight"] 

            o["last_active_date"] = latest_date(o.get("last_active_date", ""), p.get("last_active_date", "")) 

            o["pnl_per_mm_sum"] += p["pnl_per_mm_sum"] 

            o["pnl_per_mm_count"] += p["pnl_per_mm_count"] 

            o["example_cusips"].update(list(p["example_cusips"])[:8]) 

            if len(o["example_cusips"]) > 8: 

                o["example_cusips"] = set(list(o["example_cusips"])[:8]) 

    return out 

 

 

def merge_raw_entities(a, b): 

    out = defaultdict(lambda: {"sources": Counter(), "desks": Counter(), "activity": 0}) 

    for src in (a, b): 

        for raw, d in src.items(): 

            out[raw]["sources"].update(d["sources"]) 

            out[raw]["desks"].update(d["desks"]) 

            out[raw]["activity"] += d["activity"] 

    return out 

 

 

 

def account_direction_from_flow(m: dict) -> dict: 

    buy_w = float(m.get("buy_interest_weight", 0.0) or 0.0) 

    sell_w = float(m.get("sell_interest_weight", 0.0) or 0.0) 

    buy_rows = int(m.get("buy_flow_rows", 0) or 0) 

    sell_rows = int(m.get("sell_flow_rows", 0) or 0) 

    rfq_rows = int(m.get("rfq_trades", 0) or 0) 

    pnl_rows = int(m.get("pnl_trades", 0) or 0) 

    total_w = buy_w + sell_w 

    if total_w <= 0: 

        buyer_pct = 0.0 

        seller_pct = 0.0 

        direction = "RFQ-Heavy" if rfq_rows > 0 else "Low Signal" 

    else: 

        buyer_pct = buy_w / total_w 

        seller_pct = sell_w / total_w 

        if buyer_pct >= 0.62 and sell_rows <= max(2, buy_rows * 0.45): 

            direction = "Natural Buyer" 

        elif seller_pct >= 0.62 and buy_rows <= max(2, sell_rows * 0.45): 

            direction = "Natural Seller" 

        elif buyer_pct >= 0.28 and seller_pct >= 0.28: 

            direction = "Two-Way" 

        elif rfq_rows >= max(25, pnl_rows * 4) and pnl_rows < 10: 

            direction = "RFQ-Heavy" 

        else: 

            direction = "Natural Buyer" if buyer_pct >= seller_pct else "Natural Seller" 

    return { 

        "buyer_pct": buyer_pct, 

        "seller_pct": seller_pct, 

        "direction": direction, 

    } 

 

# -------------------------- scoring -------------------------- 

 

def percentile_lookup(values_by_key: dict) -> dict: 

    vals = sorted(v for v in values_by_key.values() if v is not None) 

    if not vals: 

        return {} 

    if len(vals) == 1: 

        return {k: 0.5 for k, v in values_by_key.items() if v is not None} 

    out = {} 

    for k, v in values_by_key.items(): 

        if v is None: 

            continue 

        lo = bisect.bisect_left(vals, v) 

        hi = bisect.bisect_right(vals, v) 

        avg_rank = (lo + hi - 1) / 2.0 

        out[k] = avg_rank / (len(vals) - 1) 

    return out 

 

 

def score_from_percentile(pct, neutral=5.5): 

    if pct is None: 

        return neutral 

    return 1.0 + 9.0 * pct 

 

 

def shrink_to_neutral(score, obs, full_obs=10, neutral=5.5): 

    r = min(1.0, max(0.0, obs / float(full_obs))) 

    return neutral + r * (score - neutral) 

 

 

def compute_scores(metrics: dict) -> list[dict]: 

    by_desk = defaultdict(list) 

    for key in metrics: 

        by_desk[key[1]].append(key) 

 

    scores = {} 

    for desk, keys in by_desk.items(): 

        pnl_avgs = {} 

        flow_counts = {} 

        flow_volumes = {} 

        interest_vals = {} 

        for key in keys: 

            m = metrics[key] 

            pnl_avgs[key] = safe_div(m["pnl_per_mm_sum"], m["pnl_per_mm_count"]) 

            flow_counts[key] = m["pnl_trades"] + 0.25 * m["rfq_trades"] 

            flow_volumes[key] = m["pnl_qty"] + 0.25 * m["rfq_qty"] 

            interest_vals[key] = m["linked_events"] 

        pnl_pct = percentile_lookup({k: v for k, v in pnl_avgs.items() if v is not None}) 

        flow_count_pct = percentile_lookup(flow_counts) 

        flow_volume_pct = percentile_lookup(flow_volumes) 

        interest_pct = percentile_lookup(interest_vals) 

 

        for key in keys: 

            m = metrics[key] 

            pnl_avg = pnl_avgs[key] 

            raw_pnl_score = score_from_percentile(pnl_pct.get(key), neutral=5.5) 

            pnl_score = shrink_to_neutral(raw_pnl_score, m["pnl_per_mm_count"], full_obs=10, neutral=5.5) 

 

            bad_rate = safe_div(m["negative_trades"], m["weighted_pnl_count"]) 

            if bad_rate is None: 

                adverse_score = 5.5 

            else: 

                adverse_score = 10.0 - 9.0 * bad_rate 

                adverse_score = shrink_to_neutral(adverse_score, m["weighted_pnl_count"], full_obs=10, neutral=5.5) 

 

            fc = flow_count_pct.get(key, 0.5) 

            fv = flow_volume_pct.get(key, 0.5) 

            flow_score = 1.0 + 9.0 * (0.5 * fc + 0.5 * fv) 

 

            interest_score = score_from_percentile(interest_pct.get(key), neutral=5.5) 

 

            effective_obs = m["pnl_trades"] + 0.25 * m["rfq_trades"] 

            conf01 = min(1.0, math.log1p(effective_obs) / math.log1p(100.0)) if effective_obs > 0 else 0.0 

            confidence_score = 1.0 + 9.0 * conf01 

 

            account_grade = ( 

                0.50 * pnl_score 

                + 0.20 * adverse_score 

                + 0.15 * flow_score 

                + 0.10 * interest_score 

                + 0.05 * confidence_score 

            ) 

            scores[key] = { 

                "canonical_account": key[0], 

                "desk": key[1], 

                "account_grade": clamp(account_grade), 

                "pnl_score": clamp(pnl_score), 

                "adverse_selection_score": clamp(adverse_score), 

                "flow_score": clamp(flow_score), 

                "similar_interest_score": clamp(interest_score), 

                "data_confidence_score": clamp(confidence_score), 

                "avg_weighted_pnl_per_mm": pnl_avg, 

                "avg_weighted_pnl": safe_div(m["weighted_pnl_sum"], m["weighted_pnl_count"]), 

                "pnl_trades": m["pnl_trades"], 

                "positive_trade_rate": safe_div(m["positive_trades"], m["weighted_pnl_count"]), 

                "bad_trade_rate": bad_rate, 

                "total_trade_qty": m["pnl_qty"], 

                "rfq_trade_count": m["rfq_trades"], 

                "rfq_qty": m["rfq_qty"], 

                "linked_events": m["linked_events"], 

                "linked_qty": m["linked_qty"], 

                "last_active_date": m.get("last_active_date", ""), 

                **account_direction_from_flow(m), 

                "buy_pnl_trades": m.get("buy_pnl_trades", 0), 

                "sell_pnl_trades": m.get("sell_pnl_trades", 0), 

                "buy_rfq_rows": m.get("buy_rfq_rows", 0), 

                "sell_rfq_rows": m.get("sell_rfq_rows", 0), 

                "buy_flow_rows": m.get("buy_flow_rows", 0), 

                "sell_flow_rows": m.get("sell_flow_rows", 0), 

                "buy_flow_qty": m.get("buy_flow_qty", 0.0), 

                "sell_flow_qty": m.get("sell_flow_qty", 0.0), 

                "buy_interest_weight": m.get("buy_interest_weight", 0.0), 

                "sell_interest_weight": m.get("sell_interest_weight", 0.0), 

                "top_raw_names": "; ".join(f"{k}:{v}" for k, v in m["raw_names"].most_common(6)), 

                "top_books": "; ".join(f"{k}:{v}" for k, v in m["books"].most_common(6)), 

                "top_platforms": "; ".join(f"{k}:{v}" for k, v in m["platforms"].most_common(6)), 

            } 

 

    # 

    # Normalize the final account grades. 

    # Goal: 

    #   • Average account ≈ 5 

    #   • Few accounts near 1 

    #   • Few accounts near 10 

    # 

 

        # 

    # Normalize the final account grades. 

    # Goal: 

    #   • Average account ≈ 5 

    #   • Few accounts near 1 

    #   • Few accounts near 10 

    # 

 

    rows = list(scores.values()) 

 

    # Normalize independently within each desk 

    by_desk_rows = defaultdict(list) 

 

    for r in rows: 

        by_desk_rows[r["desk"]].append(r) 

 

    for desk_rows in by_desk_rows.values(): 

 

        desk_rows.sort(key=lambda r: r["account_grade"]) 

 

        n = len(desk_rows) 

 

        for i, r in enumerate(desk_rows): 

 

            pct = i / (n - 1) if n > 1 else 0.5 

 

            # Convert percentile to range (-1, +1) 

            x = (pct - 0.5) * 2.0 

 

            # Smooth S-curve 

            x = math.tanh(1.15 * x) 

 

            # Scale to 1–10 with center at 5 

            grade = 5.0 + 4.5 * x 

 

            r["account_grade"] = round(clamp(grade, 1.0, 10.0), 4) 

 

    enrich_account_activity_states(rows) 

 

    rows.sort( 

        key=lambda r: ( 

            r["desk"], 

            -r["account_grade"], 

            -r["pnl_trades"] - r["rfq_trade_count"], 

        ) 

    ) 

 

    return rows 

 

 

# -------------------------- output helpers -------------------------- 

 

def write_account_grades(path: str, score_rows: list[dict]): 

    fields = [ 

        "canonical_account", "desk", "account_grade", "tier", "pnl_score", "adverse_selection_score", "flow_score", 

        "similar_interest_score", "data_confidence_score", "avg_weighted_pnl_per_mm", "avg_weighted_pnl", 

        "pnl_trades", "positive_trade_rate", "bad_trade_rate", "total_trade_qty", "rfq_trade_count", "rfq_qty", 

        "linked_events", "linked_qty", "last_active_date", "days_since_last_active", 

        "historical_quality_score", "historical_quality_label", "current_activity_score", "current_activity_label", 

        "account_state_label", "account_use_note", 

        "direction", "buyer_pct", "seller_pct", "buy_pnl_trades", "sell_pnl_trades", "buy_rfq_rows", "sell_rfq_rows", 

        "buy_flow_rows", "sell_flow_rows", "buy_flow_qty", "sell_flow_qty", "buy_interest_weight", "sell_interest_weight", 

        "top_raw_names", "top_books", "top_platforms" 

    ] 

    with open(path, "w", newline="", encoding="utf-8") as f: 

        w = csv.DictWriter(f, fieldnames=fields) 

        w.writeheader() 

        for r in score_rows: 

            rr = dict(r) 

            rr["tier"] = tier_from_grade(rr["account_grade"]) 

            for k in ["account_grade", "pnl_score", "adverse_selection_score", "flow_score", "similar_interest_score", "data_confidence_score", "avg_weighted_pnl_per_mm", "avg_weighted_pnl", "positive_trade_rate", "bad_trade_rate", "historical_quality_score", "current_activity_score", "buyer_pct", "seller_pct", "buy_interest_weight", "sell_interest_weight"]: 

                rr[k] = round4(rr[k]) 

            for k in ["total_trade_qty", "rfq_qty", "linked_qty", "buy_flow_qty", "sell_flow_qty"]: 

                rr[k] = round2(rr[k]) 

            w.writerow({k: rr.get(k, "") for k in fields}) 

 

 

def preferences_to_rows(preferences: dict, grade_lookup: dict) -> list[dict]: 

    out = [] 

    for key, p in preferences.items(): 

        acct, desk, ref_desk, client_flow, ticker, sector_l2, sector_l3, rb, tb = key 

        avg_pnl_per_mm = safe_div(p["pnl_per_mm_sum"], p["pnl_per_mm_count"]) 

        out.append({ 

            "canonical_account": acct, 

            "desk": desk, 

            "reference_desk": ref_desk, 

            "account_grade": grade_lookup.get((acct, desk), {}).get("account_grade", ""), 

            "client_flow": client_flow, 

            "ticker": ticker, 

            "sector_l2": sector_l2, 

            "sector_l3": sector_l3, 

            "rating_bucket": rb, 

            "tenor_bucket": tb, 

            "pnl_trades": p["pnl_trades"], 

            "rfq_trades": p["rfq_trades"], 

            "pnl_qty": p.get("pnl_qty", 0.0), 

            "rfq_qty": p.get("rfq_qty", 0.0), 

            "total_interest_weight": p["interest_weight"], 

            "total_qty": p["total_qty"], 

            "last_active_date": p.get("last_active_date", ""), 

            "avg_pnl_per_mm": avg_pnl_per_mm, 

            "appetite_side": "buyer" if client_flow == "Client_Bought" else ("seller" if client_flow == "Client_Sold" else "flow"), 

            "base_appetite_score": round4((float(grade_lookup.get((acct, desk), {}).get("account_grade") or 5.5) * 0.45) + min(10.0, 1.0 + 1.5 * math.log1p(p["interest_weight"])) * 0.40 + min(10.0, 1.0 + 1.2 * math.log1p(p["pnl_trades"] + 0.25 * p["rfq_trades"])) * 0.15), 

            "example_cusips": "; ".join(sorted(p["example_cusips"])), 

        }) 

    out.sort(key=lambda r: (-float(r["total_interest_weight"]), -float(r["account_grade"] or 0), r["canonical_account"])) 

    return out 

 

 

def write_preferences(path: str, pref_rows: list[dict]): 

    fields = [ 

        "canonical_account", "desk", "reference_desk", "account_grade", "client_flow", 

        "match_scope", "cusip", "ticker", "issuer", "sector_l2", "sector_l3", 

        "rating_bucket", "tenor_bucket", "match_label", "match_basis", "match_detail", 

        "pnl_trades", "rfq_trades", "pnl_qty", "rfq_qty", "total_interest_weight", "total_qty", 

        "last_active_date", "appetite_side", "base_appetite_score", "avg_pnl_per_mm", 

        "example_cusips", "example_cusip_count", "avg_trade_size", "activity_summary", "why_summary" 

    ] 

 

    formatted_rows = [] 

 

    for r in pref_rows: 

        rr = dict(r) 

 

        rr["account_grade"] = round4(rr["account_grade"]) if rr["account_grade"] != "" else "" 

        rr["total_interest_weight"] = round4(rr["total_interest_weight"]) 

        rr["total_qty"] = round2(rr["total_qty"]) 

        rr["pnl_qty"] = round2(rr.get("pnl_qty")) 

        rr["rfq_qty"] = round2(rr.get("rfq_qty")) 

        rr["base_appetite_score"] = round4(rr.get("base_appetite_score")) 

        rr["avg_pnl_per_mm"] = round4(rr.get("avg_pnl_per_mm")) 

        rr["avg_trade_size"] = round2(rr.get("avg_trade_size")) 

 

        formatted_rows.append({k: rr.get(k, "") for k in fields}) 

 

    write_csv_chunks( 

        path, 

        formatted_rows, 

        fields, 

        rows_per_file=100000, 

    ) 

 

def preference_rows_from_transactions(transactions: list[dict], grade_lookup: dict, fallback_rows: list[dict] | None = None) -> list[dict]: 

    """Build more tangible preference rows from row-level PNL/RFQ history. 

 

    The older preference file rolled everything into one broad ticker/sector/rating/tenor bucket. 

    This function keeps the same output filename/shape but adds explicit match_scope rows: 

      - EXACT_CUSIP: selected-bond evidence 

      - SAME_ISSUER: same ticker / issuer evidence across CUSIPs 

      - SIMILAR_SECTOR_RATING_TENOR: closest true bucket match 

      - SIMILAR_SECTOR_RATING: broader industry/rating match 

      - SIMILAR_RATING_TENOR: risk/maturity match when sector is sparse 

    """ 

    def add(agg, key, tx, scope, basis, scope_weight): 

        a = agg[key] 

        q = abs(num(tx.get("quantity"), 0) or 0) 

        src = str(tx.get("source") or "").upper() 

        if src == "PNL": 

            a["pnl_trades"] += 1 

            a["pnl_qty"] += q 

            a["interest_weight"] += 1.0 * scope_weight 

        else: 

            a["rfq_trades"] += 1 

            a["rfq_qty"] += q 

            a["interest_weight"] += 0.35 * scope_weight 

        a["total_qty"] += q 

        a["last_active_date"] = latest_date(a.get("last_active_date", ""), tx.get("trade_date", "")) 

        c = clean_text(tx.get("cusip")) 

        if c: 

            a["cusips"][c] += 1 

        iss = clean_text(tx.get("issuer")) 

        if iss: 

            a["issuers"][iss] += 1 

        t = clean_text(tx.get("ticker")) 

        if t: 

            a["tickers"][t] += 1 

        s2 = clean_text(tx.get("sector_l2")) 

        if s2: 

            a["sector_l2s"][s2] += 1 

        s3 = clean_text(tx.get("sector_l3")) 

        if s3: 

            a["sector_l3s"][s3] += 1 

        rb = clean_text(tx.get("rating_bucket")) or "Unknown" 

        if rb: 

            a["ratings"][rb] += 1 

        tb = clean_text(tx.get("tenor_bucket")) or "Unknown" 

        if tb: 

            a["tenors"][tb] += 1 

        a["scope"] = scope 

        a["basis"] = basis 

        pnlmm = tx.get("pnl_per_mm") 

        if pnlmm not in (None, ""): 

            v = num(pnlmm) 

            if v is not None: 

                a["pnl_per_mm_sum"] += v 

                a["pnl_per_mm_count"] += 1 

 

    def new_agg(): 

        return { 

            "pnl_trades": 0, "rfq_trades": 0, "pnl_qty": 0.0, "rfq_qty": 0.0, 

            "total_qty": 0.0, "interest_weight": 0.0, "last_active_date": "", 

            "cusips": Counter(), "issuers": Counter(), "tickers": Counter(), 

            "sector_l2s": Counter(), "sector_l3s": Counter(), "ratings": Counter(), "tenors": Counter(), 

            "scope": "", "basis": "", "pnl_per_mm_sum": 0.0, "pnl_per_mm_count": 0, 

        } 

 

    aggs = defaultdict(new_agg) 

    for tx in transactions: 

        acct = clean_text(tx.get("canonical_account")) 

        if not acct: 

            continue 

        desk = clean_text(tx.get("desk")) or "ALL" 

        ref_desk = clean_text(tx.get("reference_desk")) 

        flow = clean_text(tx.get("client_flow")) 

        side = "buyer" if flow == "Client_Bought" else ("seller" if flow == "Client_Sold" else "flow") 

        cusip = clean_text(tx.get("cusip")) 

        ticker = clean_text(tx.get("ticker")) 

        issuer = clean_text(tx.get("issuer")) 

        l2 = clean_text(tx.get("sector_l2")) 

        l3 = clean_text(tx.get("sector_l3")) 

        rb = clean_text(tx.get("rating_bucket")) or "Unknown" 

        tb = clean_text(tx.get("tenor_bucket")) or "Unknown" 

        coupon = clean_text(tx.get("coupon")) 

        maturity = clean_text(tx.get("maturity_date")) 

        if not (cusip or ticker or l2 or l3 or rb or tb): 

            continue 

 

        if cusip: 

            key = (acct, desk, ref_desk, flow, "EXACT_CUSIP", cusip, ticker, issuer, l2, l3, rb, tb) 

            add(aggs, key, tx, "EXACT_CUSIP", "Selected CUSIP", 1.35) 

        if ticker: 

            key = (acct, desk, ref_desk, flow, "SAME_ISSUER", "", ticker, issuer, "", "", "", "") 

            add(aggs, key, tx, "SAME_ISSUER", "Same ticker / issuer", 1.15) 

        if (l3 or l2) and rb and rb != "Unknown" and tb and tb != "Unknown": 

            key = (acct, desk, ref_desk, flow, "SIMILAR_SECTOR_RATING_TENOR", "", "", "", l2, l3, rb, tb) 

            add(aggs, key, tx, "SIMILAR_SECTOR_RATING_TENOR", "Same sector + rating + tenor", 1.00) 

        if (l3 or l2) and rb and rb != "Unknown": 

            key = (acct, desk, ref_desk, flow, "SIMILAR_SECTOR_RATING", "", "", "", l2, l3, rb, "") 

            add(aggs, key, tx, "SIMILAR_SECTOR_RATING", "Same sector + rating", 0.78) 

        if rb and rb != "Unknown" and tb and tb != "Unknown": 

            key = (acct, desk, ref_desk, flow, "SIMILAR_RATING_TENOR", "", "", "", "", "", rb, tb) 

            add(aggs, key, tx, "SIMILAR_RATING_TENOR", "Same rating + tenor", 0.58) 

 

    rows = [] 

    for key, a in aggs.items(): 

        acct, desk, ref_desk, flow, scope, cusip, ticker, issuer, l2, l3, rb, tb = key 

        grade = grade_lookup.get((acct, desk), {}).get("account_grade", "") 

        grade_num = float(grade) if grade not in (None, "") else 5.5 

        event_count = a["pnl_trades"] + 0.25 * a["rfq_trades"] 

        activity_score = min(10.0, 1.0 + 1.65 * math.log1p(max(0.0, a["interest_weight"]))) 

        count_score = min(10.0, 1.0 + 1.25 * math.log1p(max(0.0, event_count))) 

        confidence = min(10.0, 1.0 + 9.0 * (math.log1p(max(0.0, event_count)) / math.log1p(60.0))) if event_count else 1.0 

        scope_bonus = { 

            "EXACT_CUSIP": 1.25, 

            "SAME_ISSUER": 0.85, 

            "SIMILAR_SECTOR_RATING_TENOR": 0.45, 

            "SIMILAR_SECTOR_RATING": 0.15, 

            "SIMILAR_RATING_TENOR": -0.10, 

        }.get(scope, 0.0) 

        base_appetite_score = max(1.0, min(10.0, 0.35 * grade_num + 0.38 * activity_score + 0.17 * count_score + 0.10 * confidence + scope_bonus)) 

        example_cusips = "; ".join(c for c, _ in a["cusips"].most_common(10)) 

        example_count = len(a["cusips"]) 

        main_issuer = issuer or (a["issuers"].most_common(1)[0][0] if a["issuers"] else "") 

        main_ticker = ticker or (a["tickers"].most_common(1)[0][0] if a["tickers"] else "") 

        main_l2 = l2 or (a["sector_l2s"].most_common(1)[0][0] if a["sector_l2s"] else "") 

        main_l3 = l3 or (a["sector_l3s"].most_common(1)[0][0] if a["sector_l3s"] else "") 

        main_rb = rb or (a["ratings"].most_common(1)[0][0] if a["ratings"] else "") 

        main_tb = tb or (a["tenors"].most_common(1)[0][0] if a["tenors"] else "") 

        avg_pnl_per_mm = safe_div(a["pnl_per_mm_sum"], a["pnl_per_mm_count"]) 

        scope_label = { 

            "EXACT_CUSIP": "Exact selected bond", 

            "SAME_ISSUER": "Same issuer", 

            "SIMILAR_SECTOR_RATING_TENOR": "Same sector/rating/tenor", 

            "SIMILAR_SECTOR_RATING": "Same sector/rating", 

            "SIMILAR_RATING_TENOR": "Same rating/tenor", 

        }.get(scope, scope) 

        detail_bits = [] 

        if scope == "EXACT_CUSIP": 

            detail_bits = [cusip, main_ticker, main_issuer] 

        elif scope == "SAME_ISSUER": 

            detail_bits = [main_ticker, main_issuer, f"{example_count} CUSIP(s)"] 

        elif scope == "SIMILAR_SECTOR_RATING_TENOR": 

            detail_bits = [main_l3 or main_l2, main_rb, main_tb] 

        elif scope == "SIMILAR_SECTOR_RATING": 

            detail_bits = [main_l3 or main_l2, main_rb] 

        elif scope == "SIMILAR_RATING_TENOR": 

            detail_bits = [main_rb, main_tb] 

        match_detail = " · ".join(str(x) for x in detail_bits if x) 

        activity_summary = f"{a['pnl_trades']} PNL / {a['rfq_trades']} RFQ; {round2(a['total_qty'])} total qty; last {a['last_active_date'] or '-'}" 

        why_summary = f"{scope_label}: {match_detail}. Evidence: {a['pnl_trades']} PNL trade(s), {a['rfq_trades']} RFQ row(s), {example_count} CUSIP(s), last active {a['last_active_date'] or '-'}" 

        rows.append({ 

            "canonical_account": acct, 

            "desk": desk, 

            "reference_desk": ref_desk, 

            "account_grade": grade, 

            "client_flow": flow, 

            "ticker": main_ticker, 

            "sector_l2": main_l2, 

            "coupon": coupon, 

            "maturity_date": maturity, 

            "sector_l3": main_l3, 

            "rating_bucket": main_rb, 

            "tenor_bucket": main_tb, 

            "pnl_trades": a["pnl_trades"], 

            "rfq_trades": a["rfq_trades"], 

            "pnl_qty": a["pnl_qty"], 

            "rfq_qty": a["rfq_qty"], 

            "total_interest_weight": a["interest_weight"], 

            "total_qty": a["total_qty"], 

            "last_active_date": a["last_active_date"], 

            "appetite_side": side, 

            "base_appetite_score": round4(base_appetite_score), 

            "avg_pnl_per_mm": avg_pnl_per_mm, 

            "example_cusips": example_cusips, 

            "match_scope": scope, 

            "cusip": cusip, 

            "issuer": main_issuer, 

            "match_label": scope_label, 

            "match_basis": a["basis"], 

            "match_detail": match_detail, 

            "activity_summary": activity_summary, 

            "why_summary": why_summary, 

            "example_cusip_count": example_count, 

            "avg_trade_size": safe_div(a["total_qty"], a["pnl_trades"] + a["rfq_trades"]), 

        }) 

 

    if not rows and fallback_rows: 

        rows = list(fallback_rows) 

    rows.sort(key=lambda r: ( 

        -float(r.get("base_appetite_score") or 0), 

        -float(r.get("total_interest_weight") or 0), 

        -float(r.get("account_grade") or 0), 

        r.get("canonical_account", "") 

    )) 

    return rows 

 

 

def write_transactions(path: str, rows: list[dict]): 

    fields = [ 

        "source", "canonical_account", "raw_account", "desk", "book", "platform", 

        "ticket_number", "cpid", "cusip", "ticker", "issuer", "reference_desk", 

        "sector_l2", "sector_l3", "rating_bucket", "tenor_bucket", 

        "client_flow", "buy_sell", "trade_date", "settle_date", 

        "quantity", "price", 

        "pnladj1", "pnladj5", "pnladj20", 

        "weighted_pnl", "pnl_per_mm" 

    ] 

 

    rows_sorted = sorted( 

        rows, 

        key=lambda r: ( 

            r.get("canonical_account", ""), 

            r.get("desk", ""), 

            r.get("cusip", ""), 

            r.get("trade_date", "") 

        ) 

    ) 

 

    formatted_rows = [] 

 

    for r in rows_sorted: 

        rr = dict(r) 

 

        for k in ["quantity", "weighted_pnl", "pnl_per_mm"]: 

            rr[k] = round4(rr.get(k)) if rr.get(k) not in (None, "") else "" 

 

        formatted_rows.append({k: rr.get(k, "") for k in fields}) 

 

    write_csv_chunks( 

        path, 

        formatted_rows, 

        fields, 

        rows_per_file=100000, 

    ) 

 

 

def write_alias_template(path: str, raw_entities: dict): 

    fields = ["raw_name", "source_list", "desk_list", "total_activity", "suggested_canonical", "canonical_account"] 

    rows = [] 

    for raw, d in raw_entities.items(): 

        rows.append({ 

            "raw_name": raw, 

            "source_list": "; ".join(f"{k}:{v}" for k, v in d["sources"].most_common()), 

            "desk_list": "; ".join(f"{k}:{v}" for k, v in d["desks"].most_common()), 

            "total_activity": d["activity"], 

            "suggested_canonical": raw, 

            "canonical_account": "", 

        }) 

    rows.sort(key=lambda r: -r["total_activity"]) 

    with open(path, "w", newline="", encoding="utf-8") as f: 

        w = csv.DictWriter(f, fieldnames=fields) 

        w.writeheader() 

        w.writerows(rows) 

 

 

def write_alias_candidates(path: str, raw_entities: dict, top_n=2000, max_pairs=1000): 

    entities = [] 

    for raw, d in raw_entities.items(): 

        mk = alias_match_key(raw) 

        if not mk: 

            continue 

        entities.append({ 

            "raw": raw, 

            "match_key": mk, 

            "sources": "; ".join(f"{k}:{v}" for k, v in d["sources"].most_common()), 

            "desks": "; ".join(f"{k}:{v}" for k, v in d["desks"].most_common()), 

            "main_desk": d["desks"].most_common(1)[0][0] if d["desks"] else "OTHER", 

            "activity": d["activity"], 

        }) 

    entities.sort(key=lambda x: -x["activity"]) 

    entities = entities[:top_n] 

 

    blocks = defaultdict(list) 

    for i, e in enumerate(entities): 

        k = e["match_key"] 

        if len(k) >= 4: 

            blocks[("prefix", k[:4])].append(i) 

        toks = [t for t in k.split() if len(t) >= 4] 

        for t in toks[:4]: 

            blocks[("tok", t[:6])].append(i) 

        acronym = "".join(t[0] for t in k.split() if t) 

        if len(acronym) >= 2: 

            blocks[("acr", acronym)].append(i) 

 

    candidate_pairs = set() 

    for ids in blocks.values(): 

        if len(ids) < 2 or len(ids) > 250: 

            continue 

        for pos, i in enumerate(ids): 

            for j in ids[pos + 1:]: 

                if i == j: 

                    continue 

                a, b = sorted((i, j)) 

                candidate_pairs.add((a, b)) 

 

    scored = [] 

    for i, j in candidate_pairs: 

        a = entities[i] 

        b = entities[j] 

        if a["raw"] == b["raw"]: 

            continue 

        if a["main_desk"] != b["main_desk"]: 

            # Still allow if very similar across desks, but lower priority. 

            cross_desk = True 

        else: 

            cross_desk = False 

        ka, kb = a["match_key"], b["match_key"] 

        ratio = SequenceMatcher(None, ka, kb).ratio() 

        contains = (len(ka) >= 3 and ka in kb) or (len(kb) >= 3 and kb in ka) 

        acronym_a = "".join(t[0] for t in ka.split() if t) 

        acronym_b = "".join(t[0] for t in kb.split() if t) 

        acronym = (ka == acronym_b or kb == acronym_a) and min(len(ka), len(kb)) >= 2 

        token_overlap = len(set(ka.split()) & set(kb.split())) 

        score = ratio 

        if contains: 

            score = max(score, 0.92) 

        if acronym: 

            score = max(score, 0.88) 

        if token_overlap >= 2: 

            score = max(score, 0.82) 

        if cross_desk: 

            score -= 0.05 

        if score >= 0.72: 

            scored.append((score, a, b)) 

 

    scored.sort(key=lambda x: (-x[0], -(x[1]["activity"] + x[2]["activity"]))) 

    scored = scored[:max_pairs] 

 

    fields = [ 

        "suggested_group_id", "entity_a", "entity_b", "main_desk_a", "main_desk_b", "activity_a", "activity_b", 

        "source_list_a", "source_list_b", "desk_list_a", "desk_list_b", "match_score", "normalized_a", "normalized_b", 

        "review_decision", "canonical_account" 

    ] 

    with open(path, "w", newline="", encoding="utf-8") as f: 

        w = csv.DictWriter(f, fieldnames=fields) 

        w.writeheader() 

        for n, (score, a, b) in enumerate(scored, start=1): 

            w.writerow({ 

                "suggested_group_id": f"G{n:04d}", 

                "entity_a": a["raw"], 

                "entity_b": b["raw"], 

                "main_desk_a": a["main_desk"], 

                "main_desk_b": b["main_desk"], 

                "activity_a": a["activity"], 

                "activity_b": b["activity"], 

                "source_list_a": a["sources"], 

                "source_list_b": b["sources"], 

                "desk_list_a": a["desks"], 

                "desk_list_b": b["desks"], 

                "match_score": round(score, 3), 

                "normalized_a": a["match_key"], 

                "normalized_b": b["match_key"], 

                "review_decision": "", 

                "canonical_account": "", 

            }) 

 

 

def write_quality_report(path: str, quality: dict): 

    with open(path, "w", encoding="utf-8") as f: 

        json.dump(quality, f, indent=2) 

 

 

def js_str_options(values): 

    return "".join(f'<option value="{html.escape(v)}">{html.escape(v)}</option>' for v in values if v) 

 

 

def write_html(path: str, score_rows: list[dict], pref_rows: list[dict], max_pref_rows=25000): 

    grade_rows = [] 

    for r in score_rows: 

        grade_rows.append({ 

            "canonical_account": r["canonical_account"], 

            "desk": r["desk"], 

            "account_grade": round(float(r["account_grade"]), 2), 

            "tier": tier_from_grade(r["account_grade"]), 

            "pnl_score": round(float(r["pnl_score"]), 2), 

            "adverse_selection_score": round(float(r["adverse_selection_score"]), 2), 

            "flow_score": round(float(r["flow_score"]), 2), 

            "similar_interest_score": round(float(r["similar_interest_score"]), 2), 

            "data_confidence_score": round(float(r["data_confidence_score"]), 2), 

            "pnl_trades": r["pnl_trades"], 

            "rfq_trade_count": r["rfq_trade_count"], 

            "bad_trade_rate": round(float(r["bad_trade_rate"]), 4) if r.get("bad_trade_rate") is not None else None, 

            "avg_weighted_pnl_per_mm": round(float(r["avg_weighted_pnl_per_mm"]), 2) if r.get("avg_weighted_pnl_per_mm") is not None else None, 

            "top_raw_names": r["top_raw_names"], 

            "top_books": r["top_books"], 

            "top_platforms": r["top_platforms"], 

        }) 

    pref_rows_for_html = [] 

    for r in pref_rows[:max_pref_rows]: 

        pref_rows_for_html.append({ 

            "canonical_account": r["canonical_account"], 

            "desk": r["desk"], 

            "reference_desk": r["reference_desk"], 

            "client_flow": r["client_flow"], 

            "ticker": r["ticker"], 

            "sector_l2": r["sector_l2"], 

            "sector_l3": r["sector_l3"], 

            "rating_bucket": r["rating_bucket"], 

            "tenor_bucket": r["tenor_bucket"], 

            "total_interest_weight": round(float(r["total_interest_weight"]), 4), 

            "pnl_trades": r["pnl_trades"], 

            "rfq_trades": r["rfq_trades"], 

            "avg_pnl_per_mm": round(float(r["avg_pnl_per_mm"]), 2) if r["avg_pnl_per_mm"] not in (None, "") else None, 

            "example_cusips": r["example_cusips"], 

        }) 

 

    desks = sorted(set(r["desk"] for r in grade_rows)) 

    sectors = sorted(set(r["sector_l2"] for r in pref_rows_for_html if r["sector_l2"])) 

    sectors3 = sorted(set(r["sector_l3"] for r in pref_rows_for_html if r["sector_l3"])) 

    ratings = sorted(set(r["rating_bucket"] for r in pref_rows_for_html if r["rating_bucket"])) 

    tenors = ["0-2Y", "2-5Y", "5-10Y", "10-15Y", "15-25Y", "25Y+", "Unknown"] 

 

    html_text = f"""<!doctype html> 

<html lang="en"> 

<head> 

  <meta charset="utf-8" /> 

  <meta name="viewport" content="width=device-width, initial-scale=1" /> 

  <title>Account Grader</title> 

  <style> 

    :root {{ 

      --navy: #0f172a; 

      --navy2: #111827; 

      --blue: #2563eb; 

      --bg: #f3f4f6; 

      --card: #ffffff; 

      --border: #d1d5db; 

      --text: #111827; 

      --muted: #6b7280; 

      --green: #166534; 

      --red: #991b1b; 

      --amber: #92400e; 

    }} 

    * {{ box-sizing: border-box; }} 

    body {{ margin: 0; font-family: Arial, Helvetica, sans-serif; background: var(--bg); color: var(--text); }} 

    .layout {{ display: grid; grid-template-columns: 250px 1fr; min-height: 100vh; }} 

    aside {{ background: var(--navy); color: white; padding: 24px 18px; }} 

    aside h1 {{ font-size: 20px; margin: 0 0 8px; }} 

    aside p {{ color: #cbd5e1; font-size: 13px; line-height: 1.4; }} 

    main {{ padding: 24px; }} 

    .hero {{ display: flex; align-items: flex-start; justify-content: space-between; gap: 18px; margin-bottom: 18px; }} 

    .hero h2 {{ margin: 0; font-size: 28px; }} 

    .hero p {{ margin: 6px 0 0; color: var(--muted); }} 

    .cards {{ display: grid; grid-template-columns: repeat(4, minmax(160px, 1fr)); gap: 12px; margin-bottom: 18px; }} 

    .card {{ background: var(--card); border: 1px solid var(--border); border-radius: 0; padding: 16px; box-shadow: 0 1px 2px rgba(0,0,0,.04); }} 

    .metric {{ font-size: 28px; font-weight: 700; }} 

    .label {{ color: var(--muted); font-size: 12px; text-transform: uppercase; letter-spacing: .04em; }} 

    .section {{ background: var(--card); border: 1px solid var(--border); margin-bottom: 18px; }} 

    .section-header {{ padding: 16px; border-bottom: 1px solid var(--border); display: flex; align-items: center; justify-content: space-between; gap: 12px; }} 

    .section-header h3 {{ margin: 0; font-size: 18px; }} 

    .section-body {{ padding: 16px; }} 

    .filters {{ display: grid; grid-template-columns: repeat(5, minmax(120px, 1fr)); gap: 10px; align-items: end; }} 

    label {{ display: block; font-size: 12px; font-weight: 700; color: #374151; margin-bottom: 5px; }} 

    input, select, button {{ width: 100%; border: 1px solid var(--border); padding: 9px 10px; font-size: 14px; border-radius: 0; background: white; }} 

    button {{ background: var(--blue); color: white; border-color: var(--blue); cursor: pointer; font-weight: 700; }} 

    table {{ border-collapse: collapse; width: 100%; font-size: 13px; }} 

    th, td {{ border-bottom: 1px solid #e5e7eb; padding: 9px 8px; text-align: left; vertical-align: top; }} 

    th {{ background: #f9fafb; color: #374151; position: sticky; top: 0; z-index: 1; }} 

    .table-wrap {{ overflow: auto; max-height: 620px; border: 1px solid #e5e7eb; }} 

    .grade {{ font-weight: 700; }} 

    .good {{ color: var(--green); }} 

    .bad {{ color: var(--red); }} 

    .weak {{ color: var(--amber); }} 

    .muted {{ color: var(--muted); }} 

    .pill {{ display: inline-block; border: 1px solid #cbd5e1; padding: 2px 6px; font-size: 11px; margin-right: 4px; background: #f8fafc; }} 

    .note {{ font-size: 12px; color: var(--muted); line-height: 1.45; }} 

    @media (max-width: 900px) {{ .layout {{ grid-template-columns: 1fr; }} aside {{ display: none; }} .cards, .filters {{ grid-template-columns: 1fr 1fr; }} }} 

  </style> 

</head> 

<body> 

  <div class="layout"> 

    <aside> 

      <h1>AI Tool Shed</h1> 

      <p>Account Grader</p> 

      <p>Designed to sit behind the Relative Value Screener: click a cheap bond, then find accounts that historically buy similar bonds and do not pick off the firm.</p> 

    </aside> 

    <main> 

      <div class="hero"> 

        <div> 

          <h2>Account Grader</h2> 

          <p>Scores one row per account using post-trade P&amp;L, adverse selection, flow, and similar-bond interest.</p> 

        </div> 

      </div> 

 

      <div class="cards"> 

        <div class="card"><div class="metric" id="metricAccounts">0</div><div class="label">Accounts</div></div> 

        <div class="card"><div class="metric" id="metricPriority">0</div><div class="label">Priority / good</div></div> 

        <div class="card"><div class="metric" id="metricBad">0</div><div class="label">Weak / bad</div></div> 

        <div class="card"><div class="metric" id="metricPrefs">0</div><div class="label">Bond preference rows</div></div> 

      </div> 

 

      <div class="section"> 

        <div class="section-header"> 

          <h3>Relative Value Screener click simulation</h3> 

          <span class="note">Use this to mimic clicking a cheap bond and ranking likely buyers.</span> 

        </div> 

        <div class="section-body"> 

          <div class="filters"> 

            <div><label>Desk</label><select id="fitDesk"><option value="">Any</option>{js_str_options(desks)}</select></div> 

            <div><label>Sector L2</label><select id="fitSector"><option value="">Any</option>{js_str_options(sectors)}</select></div> 

            <div><label>Sector L3</label><select id="fitSector3"><option value="">Any</option>{js_str_options(sectors3)}</select></div> 

            <div><label>Rating</label><select id="fitRating"><option value="">Any</option>{js_str_options(ratings)}</select></div> 

            <div><label>Tenor</label><select id="fitTenor"><option value="">Any</option>{js_str_options(tenors)}</select></div> 

            <div><label>Ticker</label><input id="fitTicker" placeholder="Optional, e.g. F" /></div> 

            <div><label>Client flow</label><select id="fitFlow"><option value="Client_Bought">Client bought bond</option><option value="">Any</option><option value="Client_Sold">Client sold bond</option></select></div> 

            <div><label>Minimum account grade</label><input id="fitMinGrade" type="number" min="1" max="10" step="0.1" value="1" /></div> 

            <div><label>&nbsp;</label><button onclick="rankClients()">Rank Clients</button></div> 

          </div> 

          <p class="note">Formula: Client Recommendation Score = 0.65 * Account Grade + 0.35 * Similar Bond Fit Score. Similarity uses desk, sector, ticker, rating, tenor, client-flow direction, and prior interest weight.</p> 

          <div class="table-wrap"><table id="recommendationTable"></table></div> 

        </div> 

      </div> 

 

      <div class="section"> 

        <div class="section-header"> 

          <h3>Account grades</h3> 

          <span class="note">1 = bad account, 10 = priority account.</span> 

        </div> 

        <div class="section-body"> 

          <div class="filters"> 

            <div><label>Search account</label><input id="accountSearch" placeholder="Type account name" oninput="renderGrades()" /></div> 

            <div><label>Desk</label><select id="deskFilter" onchange="renderGrades()"><option value="">All</option>{js_str_options(desks)}</select></div> 

            <div><label>Minimum grade</label><input id="minGrade" type="number" min="1" max="10" step="0.1" value="1" oninput="renderGrades()" /></div> 

            <div><label>Show rows</label><input id="maxRows" type="number" min="25" max="1000" step="25" value="250" oninput="renderGrades()" /></div> 

            <div><label>&nbsp;</label><button onclick="renderGrades()">Apply</button></div> 

          </div> 

          <div class="table-wrap"><table id="gradesTable"></table></div> 

        </div> 

      </div> 

 

      <p class="note">Build notes: This version assumes BUYSELL is from the firm's perspective. Confirm the convention before production. True hit rate requires lost/passed RFQs; if the RFQ sheet only has trades, this uses RFQ/trade interest instead.</p> 

    </main> 

  </div> 

 

<script> 

const GRADES = {json.dumps(grade_rows)}; 

const PREFS = {json.dumps(pref_rows_for_html)}; 

const PREFS_BY_ACCOUNT = {{}}; 

for (const p of PREFS) {{ 

  const k = p.canonical_account + '|' + p.desk; 

  if (!PREFS_BY_ACCOUNT[k]) PREFS_BY_ACCOUNT[k] = []; 

  PREFS_BY_ACCOUNT[k].push(p); 

}} 

 

function gradeClass(g) {{ 

  if (g < 5) return 'bad'; 

  if (g < 7) return 'weak'; 

  return 'good'; 

}} 

function fmt(x, digits=2) {{ 

  if (x === null || x === undefined || x === '') return ''; 

  const n = Number(x); 

  if (Number.isNaN(n)) return x; 

  return n.toLocaleString(undefined, {{maximumFractionDigits: digits}}); 

}} 

function esc(s) {{ 

  return String(s ?? '').replace(/[&<>"']/g, m => ({{'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}}[m])); 

}} 

 

function renderMetrics() {{ 

  document.getElementById('metricAccounts').textContent = GRADES.length.toLocaleString(); 

  document.getElementById('metricPriority').textContent = GRADES.filter(r => r.account_grade >= 7).length.toLocaleString(); 

  document.getElementById('metricBad').textContent = GRADES.filter(r => r.account_grade < 5).length.toLocaleString(); 

  document.getElementById('metricPrefs').textContent = PREFS.length.toLocaleString(); 

}} 

 

function renderGrades() {{ 

  const q = document.getElementById('accountSearch').value.toUpperCase(); 

  const desk = document.getElementById('deskFilter').value; 

  const minGrade = Number(document.getElementById('minGrade').value || 1); 

  const maxRows = Number(document.getElementById('maxRows').value || 250); 

  let rows = GRADES.filter(r => (!q || r.canonical_account.includes(q)) && (!desk || r.desk === desk) && r.account_grade >= minGrade); 

  rows = rows.sort((a,b) => b.account_grade - a.account_grade).slice(0, maxRows); 

  let out = '<thead><tr><th>Account</th><th>Desk</th><th>Grade</th><th>Tier</th><th>P&amp;L</th><th>Adverse</th><th>Flow</th><th>Interest</th><th>PNL Trades</th><th>RFQ Trades</th><th>Bad Rate</th><th>Avg PNL / $1mm</th><th>Raw names</th></tr></thead><tbody>'; 

  for (const r of rows) {{ 

    out += `<tr> 

      <td><strong>${{esc(r.canonical_account)}}</strong><br><span class="muted">${{esc(r.top_books || '')}}</span></td> 

      <td><span class="pill">${{esc(r.desk)}}</span></td> 

      <td class="grade ${{gradeClass(r.account_grade)}}">${{fmt(r.account_grade)}}</td> 

      <td>${{esc(r.tier)}}</td> 

      <td>${{fmt(r.pnl_score)}}</td> 

      <td>${{fmt(r.adverse_selection_score)}}</td> 

      <td>${{fmt(r.flow_score)}}</td> 

      <td>${{fmt(r.similar_interest_score)}}</td> 

      <td>${{fmt(r.pnl_trades, 0)}}</td> 

      <td>${{fmt(r.rfq_trade_count, 0)}}</td> 

      <td>${{r.bad_trade_rate === null ? '' : fmt(100*r.bad_trade_rate, 1) + '%'}}</td> 

      <td>${{fmt(r.avg_weighted_pnl_per_mm)}}</td> 

      <td class="muted">${{esc(r.top_raw_names || '')}}</td> 

    </tr>`; 

  }} 

  out += '</tbody>'; 

  document.getElementById('gradesTable').innerHTML = out; 

}} 

 

function prefFit(p, inputs) {{ 

  if (inputs.desk && p.desk !== inputs.desk) return 0; 

  let pts = 0, max = 0; 

  if (inputs.desk) {{ max += 12; pts += 12; }} 

  if (inputs.flow) {{ max += 10; if (p.client_flow === inputs.flow) pts += 10; }} 

  if (inputs.ticker) {{ max += 28; if ((p.ticker || '').toUpperCase() === inputs.ticker) pts += 28; }} 

  if (inputs.sector) {{ max += 22; if (p.sector_l2 === inputs.sector) pts += 22; }} 

  if (inputs.sector3) {{ max += 10; if (p.sector_l3 === inputs.sector3) pts += 10; }} 

  if (inputs.rating) {{ max += 14; if (p.rating_bucket === inputs.rating) pts += 14; }} 

  if (inputs.tenor) {{ max += 14; if (p.tenor_bucket === inputs.tenor) pts += 14; }} 

  if (max === 0) return 0; 

  const match01 = pts / max; 

  const intensity01 = Math.min(1, Math.log1p(Number(p.total_interest_weight || 0)) / Math.log1p(25)); 

  return 1 + 9 * (0.82 * match01 + 0.18 * intensity01); 

}} 

 

function rankClients() {{ 

  const inputs = {{ 

    desk: document.getElementById('fitDesk').value, 

    sector: document.getElementById('fitSector').value, 

    sector3: document.getElementById('fitSector3').value, 

    rating: document.getElementById('fitRating').value, 

    tenor: document.getElementById('fitTenor').value, 

    ticker: document.getElementById('fitTicker').value.trim().toUpperCase(), 

    flow: document.getElementById('fitFlow').value, 

    minGrade: Number(document.getElementById('fitMinGrade').value || 1) 

  }}; 

  const recs = []; 

  for (const g of GRADES) {{ 

    if (inputs.desk && g.desk !== inputs.desk) continue; 

    if (g.account_grade < inputs.minGrade) continue; 

    const prefs = PREFS_BY_ACCOUNT[g.canonical_account + '|' + g.desk] || []; 

    let bestFit = 0; 

    let bestPref = null; 

    for (const p of prefs) {{ 

      const fs = prefFit(p, inputs); 

      if (fs > bestFit) {{ bestFit = fs; bestPref = p; }} 

    }} 

    if (bestFit <= 0) continue; 

    const recScore = 0.65 * g.account_grade + 0.35 * bestFit; 

    recs.push({{g, bestFit, bestPref, recScore}}); 

  }} 

  recs.sort((a,b) => b.recScore - a.recScore); 

  const rows = recs.slice(0, 50); 

  let out = '<thead><tr><th>Rank</th><th>Account</th><th>Desk</th><th>Recommendation</th><th>Account Grade</th><th>Fit</th><th>Why it matched</th><th>History</th></tr></thead><tbody>'; 

  rows.forEach((r, i) => {{ 

    const p = r.bestPref || {{}}; 

    const why = [p.ticker, p.sector_l2, p.sector_l3, p.rating_bucket, p.tenor_bucket, p.client_flow].filter(Boolean).map(x => `<span class="pill">${{esc(x)}}</span>`).join(' '); 

    out += `<tr> 

      <td>${{i+1}}</td> 

      <td><strong>${{esc(r.g.canonical_account)}}</strong><br><span class="muted">${{esc(r.g.tier)}}</span></td> 

      <td>${{esc(r.g.desk)}}</td> 

      <td class="grade ${{gradeClass(r.recScore)}}">${{fmt(r.recScore)}}</td> 

      <td>${{fmt(r.g.account_grade)}}</td> 

      <td>${{fmt(r.bestFit)}}</td> 

      <td>${{why}}</td> 

      <td>PNL trades: ${{fmt(p.pnl_trades, 0)}}<br>RFQ trades: ${{fmt(p.rfq_trades, 0)}}<br><span class="muted">CUSIPs: ${{esc(p.example_cusips || '')}}</span></td> 

    </tr>`; 

  }}); 

  out += '</tbody>'; 

  document.getElementById('recommendationTable').innerHTML = out; 

}} 

 

renderMetrics(); 

renderGrades(); 

rankClients(); 

</script> 

</body> 

</html> 

""" 

    with open(path, "w", encoding="utf-8") as f: 

        f.write(html_text) 

 

 

def main(): 

    parser = argparse.ArgumentParser(description="Build HY-only account grader CSVs and HTML from PNL, RFQ, and HY ICE files.") 

    parser.add_argument("--pnl", default=DEFAULT_PNL, help="Path to PNL xlsx") 

    parser.add_argument("--rfq", default=DEFAULT_RFQ, help="Path to RFQ/trades xlsx") 

    parser.add_argument("--ig", default=DEFAULT_IG, help="Optional IG ICE csv/folder. Leave blank for HY-only mode.") 

    parser.add_argument("--hy", default=DEFAULT_HY, help="Path to HY ICE csv or folder of ICE H0A0 csv files") 

    parser.add_argument("--alias-map", default="account_alias_map.csv", help="Optional raw_name/canonical_account CSV. If missing, raw names are used.") 

    parser.add_argument("--out-dir", default="account_grader_output", help="Output directory") 

    parser.add_argument("--sample-rows", type=int, default=0, help="Optional row limit per Excel file for testing. 0 = full file.") 

    parser.add_argument("--alias-top-n", type=int, default=2000, help="Top active raw names to compare for alias candidates.") 

    parser.add_argument("--alias-max-pairs", type=int, default=1000, help="Max alias candidate pairs to output.") 

    parser.add_argument("--html-pref-rows", type=int, default=25000, help="Max preference rows embedded into HTML.") 

    args = parser.parse_args() 

 

    started = time.time() 

    os.makedirs(args.out_dir, exist_ok=True) 

 

    for label, path in [("PNL", args.pnl), ("RFQ", args.rfq), ("HY", args.hy)]: 

        if not os.path.exists(path): 

            raise FileNotFoundError(f"{label} input/path not found: {path}") 

    if args.ig and not os.path.exists(args.ig): 

        raise FileNotFoundError(f"IG input/path not found: {args.ig}") 

 

    print("Reading HY-only bond reference...") 

    bond_ref = read_bond_reference(args.ig, args.hy) 

    print(f"Loaded {len(bond_ref):,} HY reference bonds") 

 

    aliases = load_alias_map(args.alias_map) 

    if not aliases: 

        print("No alias map loaded. First run will score raw names separately and produce alias review files.") 

 

    # 

    # Process all PNL CSVs 

    # 

    print("Processing PNL folder...") 

 

    pnl_metrics = defaultdict(new_metric) 

    pnl_prefs = defaultdict(new_pref) 

    pnl_entities = defaultdict(lambda: {"sources": Counter(), "desks": Counter(), "activity": 0}) 

    pnl_quality = {} 

    pnl_transactions = [] 

 

    for f in sorted(os.listdir(args.pnl)): 

        if not f.lower().endswith(".csv"): 

            continue 

 

        path = os.path.join(args.pnl, f) 

        print(f"  {f}") 

 

        m, p, e, q, t = process_pnl(path, bond_ref, aliases, args.sample_rows) 

 

        pnl_metrics = merge_metrics(pnl_metrics, m) 

        pnl_prefs = merge_preferences(pnl_prefs, p) 

        pnl_entities = merge_raw_entities(pnl_entities, e) 

        pnl_transactions.extend(t) 

 

    print(f"PNL account groups: {len(pnl_metrics):,}") 

 

    # 

    # Process all RFQ CSVs 

    # 

    print("Processing RFQ folder...") 

 

    rfq_metrics = defaultdict(new_metric) 

    rfq_prefs = defaultdict(new_pref) 

    rfq_entities = defaultdict(lambda: {"sources": Counter(), "desks": Counter(), "activity": 0}) 

    rfq_quality = {} 

    rfq_transactions = [] 

 

    for f in sorted(os.listdir(args.rfq)): 

        if not f.lower().endswith(".csv"): 

            continue 

 

        path = os.path.join(args.rfq, f) 

        print(f"  {f}") 

 

        m, p, e, q, t = process_rfq(path, bond_ref, aliases, args.sample_rows) 

 

        rfq_metrics = merge_metrics(rfq_metrics, m) 

        rfq_prefs = merge_preferences(rfq_prefs, p) 

        rfq_entities = merge_raw_entities(rfq_entities, e) 

        rfq_transactions.extend(t) 

 

    print(f"RFQ account groups: {len(rfq_metrics):,}") 

 

    metrics = merge_metrics(pnl_metrics, rfq_metrics) 

    preferences = merge_preferences(pnl_prefs, rfq_prefs) 

    raw_entities = merge_raw_entities(pnl_entities, rfq_entities) 

 

    print("Scoring accounts...") 

    score_rows = compute_scores(metrics) 

    grade_lookup = {(r["canonical_account"], r["desk"]): r for r in score_rows} 

    legacy_pref_rows = preferences_to_rows(preferences, grade_lookup) 

    all_transactions = pnl_transactions + rfq_transactions 

    pref_rows = preference_rows_from_transactions(all_transactions, grade_lookup, fallback_rows=legacy_pref_rows) 

 

    grades_path = os.path.join(args.out_dir, "account_grades.csv") 

    prefs_path = os.path.join(args.out_dir, "account_bond_preferences.csv") 

    tx_path = os.path.join(args.out_dir, "account_trade_history.csv") 

    alias_template_path = os.path.join(args.out_dir, "account_alias_map_template.csv") 

    alias_candidates_path = os.path.join(args.out_dir, "mpid_alias_candidates.csv") 

    html_path = os.path.join(args.out_dir, "account_grader.html") 

    quality_path = os.path.join(args.out_dir, "data_quality_report.json") 

 

    print("Writing outputs...") 

    write_account_grades(grades_path, score_rows) 

    write_preferences(prefs_path, pref_rows) 

    write_transactions(tx_path, all_transactions) 

    write_alias_template(alias_template_path, raw_entities) 

    write_alias_candidates(alias_candidates_path, raw_entities, top_n=args.alias_top_n, max_pairs=args.alias_max_pairs) 

    write_html(html_path, score_rows, pref_rows, max_pref_rows=args.html_pref_rows) 

 

    quality = { 

        **pnl_quality, 

        **rfq_quality, 

        "bond_reference_count": len(bond_ref), 

        "account_desk_rows": len(score_rows), 

        "preference_rows": len(pref_rows), 

        "transaction_rows": len(pnl_transactions) + len(rfq_transactions), 

        "raw_entity_count": len(raw_entities), 

        "formula": { 

            "weighted_post_trade_pnl": "0.20 * PNLADJ1 + 0.50 * PNLADJ5 + 0.30 * PNLADJ20", 

            "account_grade": "0.50 * PNL Score + 0.20 * Adverse Selection Score + 0.15 * Flow Score + 0.10 * Similar Bond Interest Score + 0.05 * Data Confidence Score", 

            "account_state": "Historical Quality and Current Activity are separate account-grader outputs. Historical Quality uses P&L/adverse selection/confidence; Current Activity uses last active date relative to the newest activity in the file plus flow score.", 

            "client_recommendation_score": "0.30 exact CUSIP + 0.25 same issuer + 0.20 tangible similar match (sector/rating/tenor, sector/rating, or rating/tenor) + 0.10 RFQ + 0.10 account grade + 0.05 confidence; optional user-adjusted recency overlay in HTML", 

            "important_assumption": "BUYSELL is from firm's perspective: S = firm sold/client bought; B = firm bought/client sold.", 

        }, 

        "outputs": { 

            "account_grades": grades_path, 

            "account_bond_preferences": prefs_path, 

            "account_trade_history": tx_path, 

            "account_alias_map_template": alias_template_path, 

            "mpid_alias_candidates": alias_candidates_path, 

            "account_grader_html": html_path, 

        }, 

        "runtime_seconds": round(time.time() - started, 2), 

    } 

    write_quality_report(quality_path, quality) 

 

    print("\nDone.") 

    print(f"Account grades:        {grades_path}") 

    print(f"Bond preferences:      {prefs_path}") 

    print(f"Trade history:         {tx_path}") 

    print(f"Alias map template:    {alias_template_path}") 

    print(f"Alias candidates:      {alias_candidates_path}") 

    print(f"HTML page:             {html_path}") 

    print(f"Quality report:        {quality_path}") 

    print("\nNext step: review mpid_alias_candidates.csv and account_alias_map_template.csv.") 

    print("Fill canonical_account in a file named account_alias_map.csv, then rerun this script to apply grouping.") 

 

def write_csv_chunks(path, rows, fieldnames, rows_per_file=50000): 

    """ 

    Write one or more CSV files. 

 

    If rows <= rows_per_file: 

        writes: 

            account_trade_history.csv 

 

    Otherwise writes: 

            account_trade_history_1.csv 

            account_trade_history_2.csv 

            ... 

    """ 

 

    if len(rows) <= rows_per_file: 

        with open(path, "w", newline="", encoding="utf-8") as f: 

            w = csv.DictWriter(f, fieldnames=fieldnames) 

            w.writeheader() 

            w.writerows(rows) 

        return 

 

    base, ext = os.path.splitext(path) 

 

    for chunk_num, start in enumerate(range(0, len(rows), rows_per_file), start=1): 

 

        chunk = rows[start:start + rows_per_file] 

 

        chunk_path = f"{base}_{chunk_num}{ext}" 

 

        with open(chunk_path, "w", newline="", encoding="utf-8") as f: 

            w = csv.DictWriter(f, fieldnames=fieldnames) 

            w.writeheader() 

            w.writerows(chunk) 

 

        print(f"Wrote {chunk_path} ({len(chunk):,} rows)") 

 

 

if __name__ == "__main__": 

    try: 

        main() 

    except Exception as e: 

        print(f"ERROR: {e}", file=sys.stderr) 

        sys.exit(1) 
