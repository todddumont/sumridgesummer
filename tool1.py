"""
SumRidge AI Tool — Capability 1: Bonds With Recent News
"""

import csv
import time
import anthropic
import yfinance as yf
from datetime import datetime

ANTHROPIC_API_KEY = ""   # <-- insert key

OUTPUT_FILE = "capability1_news_output.csv"
MAX_HEADLINES_PER_LEVEL = 5


def load_inventory() -> list[dict]:
    return [
        {
            "cusip": "456789DD4",
            "issuer": "Turnpike Authority Revenue",
            "ticker": None,
            "sector": "Transportation",
            "rating": "A",
            "maturity": "2035-05-01",
            "bond_type": "Muni",
        },
        {
            "cusip": "123456AA1",
            "issuer": "Ford Motor Credit",
            "ticker": "F",
            "sector": "Automotive",
            "rating": "BB+",
            "maturity": "2028-03-15",
            "bond_type": "Corporate HY",
        },
    ]


def fetch_news_level1(bond: dict) -> list[dict]:
    headlines = []
    if bond.get("ticker"):
        try:
            news = yf.Ticker(bond["ticker"]).news or []
            for item in news[:MAX_HEADLINES_PER_LEVEL]:
                headlines.append({"headline": item.get("title", ""), "source": item.get("publisher", ""), "url": item.get("link", ""), "level_searched": 1})
        except Exception as e:
            print(f"  [Level 1] ticker error: {e}")
    if not headlines:
        try:
            search_term = bond["issuer"].split()[0]
            news = yf.Ticker(search_term).news or []
            for item in news[:MAX_HEADLINES_PER_LEVEL]:
                headlines.append({"headline": item.get("title", ""), "source": item.get("publisher", ""), "url": item.get("link", ""), "level_searched": 1})
        except Exception as e:
            print(f"  [Level 1] issuer search error: {e}")
    return headlines


def fetch_news_level2(bond: dict) -> list[dict]:
    sector_proxies = {
        "Transportation": "TAN", "Automotive": "CARZ", "Energy": "XLE",
        "Technology": "XLK", "Healthcare": "XLV", "Financials": "XLF",
        "Real Estate": "VNQ", "Utilities": "XLU",
    }
    headlines = []
    proxy = sector_proxies.get(bond.get("sector", ""))
    if proxy:
        try:
            news = yf.Ticker(proxy).news or []
            for item in news[:MAX_HEADLINES_PER_LEVEL]:
                headlines.append({"headline": item.get("title", ""), "source": item.get("publisher", ""), "url": item.get("link", ""), "level_searched": 2})
        except Exception as e:
            print(f"  [Level 2] sector error: {e}")
    return headlines


def fetch_news_level3(bond: dict) -> list[dict]:
    proxy = "JNK" if "HY" in bond.get("bond_type", "") else "AGG"
    headlines = []
    try:
        news = yf.Ticker(proxy).news or []
        for item in news[:MAX_HEADLINES_PER_LEVEL]:
            headlines.append({"headline": item.get("title", ""), "source": item.get("publisher", ""), "url": item.get("link", ""), "level_searched": 3})
    except Exception as e:
        print(f"  [Level 3] macro error: {e}")
    return headlines


def classify_headlines(bond: dict, headlines: list[dict], level: int, client: anthropic.Anthropic) -> list[dict]:
    if not headlines:
        return []

    level_prompts = {
        1: f"This bond is issued by: {bond['issuer']} ({bond['bond_type']}, {bond['rating']}, matures {bond['maturity']}).\nFor each headline, does it DIRECTLY reference this issuer or its debt?\nClassify each as DIRECT_ISSUER or NOT_RELEVANT, then give ONE sentence of implication.",
        2: f"This bond is in the {bond['sector']} sector ({bond['bond_type']}, {bond['rating']}, matures {bond['maturity']}).\nFor each headline, does it have a clear implication for {bond['sector']} bonds?\nClassify each as SECTOR or NOT_RELEVANT, then give ONE sentence of implication.",
        3: f"This is a {bond['bond_type']} bond rated {bond['rating']} maturing {bond['maturity']}.\nFor each headline, does it have a specific macro/rate/credit implication for this bond type?\nClassify each as BROADER_IMPLICATION or NOT_RELEVANT, then give ONE sentence of implication.",
    }

    headline_list = "\n".join(f"{i+1}. {h['headline']}" for i, h in enumerate(headlines))
    prompt = (
        f"{level_prompts[level]}\n\nHeadlines:\n{headline_list}\n\n"
        "Respond in this exact format for each headline (one per line):\n"
        "N | CLASSIFICATION | one sentence implication\n"
        "Example: 1 | DIRECT_ISSUER | Company announced covenant breach on 2035 notes."
    )

    try:
        message = client.messages.create(
            model="claude-opus-4-6",
            max_tokens=1024,
            messages=[{"role": "user", "content": prompt}]
        )
        response_text = message.content[0].text
    except Exception as e:
        print(f"  [Claude] API error: {e}")
        return []

    relevant = []
    for line in response_text.strip().split("\n"):
        parts = [p.strip() for p in line.split("|")]
        if len(parts) < 3:
            continue
        try:
            idx = int(parts[0]) - 1
            classification = parts[1].strip().upper()
            implication = parts[2].strip()
            if classification not in ("NOT_RELEVANT", "") and idx < len(headlines):
                h = headlines[idx].copy()
                h["classification"] = classification
                h["implication"] = implication
                relevant.append(h)
        except (ValueError, IndexError):
            continue

    return relevant


def build_rows(bond: dict, headlines: list[dict]) -> list[dict]:
    return [{
        "cusip": bond["cusip"], "issuer": bond["issuer"], "sector": bond["sector"],
        "rating": bond["rating"], "maturity": bond["maturity"], "bond_type": bond["bond_type"],
        "headline": h.get("headline", ""), "source": h.get("source", ""), "url": h.get("url", ""),
        "news_level": h.get("level_searched", ""), "classification": h.get("classification", ""),
        "implication": h.get("implication", ""), "run_date": datetime.today().strftime("%Y-%m-%d"),
    } for h in headlines]


def run_capability1():
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    inventory = load_inventory()
    results = []

    for bond in inventory:
        print(f"\nProcessing: {bond['issuer']} ({bond['cusip']})")
        matched = False

        print("  Searching Level 1: Direct Issuer...")
        l1 = classify_headlines(bond, fetch_news_level1(bond), 1, client)
        if l1:
            print(f"  → {len(l1)} hit(s) at Level 1")
            results.extend(build_rows(bond, l1))
            matched = True

        if not matched:
            print("  Searching Level 2: Sector...")
            l2 = classify_headlines(bond, fetch_news_level2(bond), 2, client)
            if l2:
                print(f"  → {len(l2)} hit(s) at Level 2")
                results.extend(build_rows(bond, l2))
                matched = True

        if not matched:
            print("  Searching Level 3: Broader Market...")
            l3 = classify_headlines(bond, fetch_news_level3(bond), 3, client)
            if l3:
                print(f"  → {len(l3)} hit(s) at Level 3")
                results.extend(build_rows(bond, l3))
                matched = True

        if not matched:
            print("  → No relevant news found")
            results.append({
                "cusip": bond["cusip"], "issuer": bond["issuer"], "sector": bond["sector"],
                "rating": bond["rating"], "maturity": bond["maturity"], "bond_type": bond["bond_type"],
                "headline": "NO RELEVANT NEWS FOUND", "source": "", "url": "",
                "news_level": "", "classification": "NOT_RELEVANT", "implication": "",
                "run_date": datetime.today().strftime("%Y-%m-%d"),
            })

        time.sleep(0.5)

    fieldnames = ["cusip", "issuer", "sector", "rating", "maturity", "bond_type",
                  "headline", "source", "url", "news_level", "classification", "implication", "run_date"]
    with open(OUTPUT_FILE, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(results)

    print(f"\n✓ Done — {len(results)} rows written to {OUTPUT_FILE}")


if __name__ == "__main__":
    run_capability1()
