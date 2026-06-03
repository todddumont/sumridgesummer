"""
SumRidge AI Tool - Tool 1: Inventory Intelligence
Capability 1: Bonds With Recent News

Scalable version:
- Does NOT search news one CUSIP at a time.
- Groups inventory by issuer/ticker, sector, and broader market group.
- Searches each group once.
- Uses Claude, when available, to verify relevance and write one implication.
- Maps matched news back to affected books and CUSIPs.
- Writes grouped output to capability1_news_output.csv.

Important:
- Without ANTHROPIC_API_KEY, the tool will NOT treat candidate headlines as relevant.
- This prevents false matches while the Claude relevance engine is unavailable.
"""

import csv
import os
import time
from datetime import datetime

import anthropic
import yfinance as yf


ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY", "")

OUTPUT_FILE = "capability1_news_output.csv"
MAX_HEADLINES_PER_GROUP = 5
MAX_SAMPLE_CUSIPS = 8


# ============================================================
# 1. Sample inventory
# Later this should be replaced with internal inventory/API data.
# ============================================================

def load_inventory() -> list[dict]:
    return [
        {
            "cusip": "456789DD4",
            "book": "Muni Inventory",
            "issuer": "Turnpike Authority Revenue",
            "ticker": "",
            "sector": "Transportation",
            "rating": "A",
            "maturity": "2035-05-01",
            "bond_type": "Muni",
        },
        {
            "cusip": "456789DD5",
            "book": "Muni Inventory",
            "issuer": "Turnpike Authority Revenue",
            "ticker": "",
            "sector": "Transportation",
            "rating": "A",
            "maturity": "2038-05-01",
            "bond_type": "Muni",
        },
        {
            "cusip": "678901FF6",
            "book": "Muni Inventory",
            "issuer": "Airport Revenue Bonds",
            "ticker": "",
            "sector": "Transportation",
            "rating": "A",
            "maturity": "2040-01-01",
            "bond_type": "Muni",
        },
        {
            "cusip": "345678CC3",
            "book": "Muni Inventory",
            "issuer": "Regional Healthcare System",
            "ticker": "",
            "sector": "Healthcare",
            "rating": "BBB",
            "maturity": "2037-11-01",
            "bond_type": "Muni",
        },
        {
            "cusip": "123456AA1",
            "book": "Corporate HY",
            "issuer": "Ford Motor Credit",
            "ticker": "F",
            "sector": "Automotive",
            "rating": "BB+",
            "maturity": "2028-03-15",
            "bond_type": "Corporate HY",
        },
        {
            "cusip": "123456AA2",
            "book": "Corporate HY",
            "issuer": "Ford Motor Credit",
            "ticker": "F",
            "sector": "Automotive",
            "rating": "BB+",
            "maturity": "2030-06-15",
            "bond_type": "Corporate HY",
        },
        {
            "cusip": "999999IG1",
            "book": "Corporate IG",
            "issuer": "Apple Inc",
            "ticker": "AAPL",
            "sector": "Technology",
            "rating": "AA+",
            "maturity": "2031-08-05",
            "bond_type": "Corporate IG",
        },
    ]


# ============================================================
# 2. Helpers
# ============================================================

def clean_text(value) -> str:
    if value is None:
        return ""

    return str(value).strip()


def unique_join(values: list[str], separator: str = ", ") -> str:
    cleaned = []

    for value in values:
        value = clean_text(value)

        if value and value not in cleaned:
            cleaned.append(value)

    return separator.join(cleaned)


def sample_values(values: list[str], max_count: int = MAX_SAMPLE_CUSIPS) -> str:
    cleaned = []

    for value in values:
        value = clean_text(value)

        if value and value not in cleaned:
            cleaned.append(value)

    if len(cleaned) <= max_count:
        return ", ".join(cleaned)

    shown = cleaned[:max_count]
    remaining = len(cleaned) - max_count

    return ", ".join(shown) + f", +{remaining} more"


def normalize_inventory(inventory: list[dict]) -> list[dict]:
    normalized = []

    for bond in inventory:
        normalized.append(
            {
                "cusip": clean_text(bond.get("cusip")),
                "book": clean_text(bond.get("book")),
                "issuer": clean_text(bond.get("issuer")),
                "ticker": clean_text(bond.get("ticker")).upper(),
                "sector": clean_text(bond.get("sector")),
                "rating": clean_text(bond.get("rating")),
                "maturity": clean_text(bond.get("maturity")),
                "bond_type": clean_text(bond.get("bond_type")),
            }
        )

    return normalized


# ============================================================
# 3. Create search groups
# This is the important scalability change.
# ============================================================

def create_issuer_groups(inventory: list[dict]) -> list[dict]:
    group_map = {}

    for bond in inventory:
        issuer = bond["issuer"]
        ticker = bond["ticker"]

        if ticker:
            group_key = f"ticker:{ticker}"
            group_name = issuer
            search_ticker = ticker
        else:
            group_key = f"issuer:{issuer.lower()}"
            group_name = issuer
            search_ticker = ""

        if group_key not in group_map:
            group_map[group_key] = {
                "group_type": "Issuer",
                "news_level": "DIRECT_ISSUER",
                "group_key": group_key,
                "group_name": group_name,
                "search_ticker": search_ticker,
                "search_term": group_name,
                "bonds": [],
            }

        group_map[group_key]["bonds"].append(bond)

    return list(group_map.values())


def create_sector_groups(inventory: list[dict]) -> list[dict]:
    sector_proxies = {
        "Transportation": "IYT",
        "Automotive": "CARZ",
        "Energy": "XLE",
        "Technology": "XLK",
        "Healthcare": "XLV",
        "Financials": "XLF",
        "Real Estate": "VNQ",
        "Utilities": "XLU",
    }

    group_map = {}

    for bond in inventory:
        sector = bond["sector"] or "Unknown"
        group_key = f"sector:{sector.lower()}"

        if group_key not in group_map:
            group_map[group_key] = {
                "group_type": "Sector",
                "news_level": "SECTOR",
                "group_key": group_key,
                "group_name": sector,
                "search_ticker": sector_proxies.get(sector, ""),
                "search_term": sector,
                "bonds": [],
            }

        group_map[group_key]["bonds"].append(bond)

    return list(group_map.values())


def create_broader_market_groups(inventory: list[dict]) -> list[dict]:
    group_map = {}

    for bond in inventory:
        bond_type = bond["bond_type"] or "Unknown"

        if "HY" in bond_type.upper():
            group_name = "High Yield Corporates"
            search_ticker = "JNK"
        elif "IG" in bond_type.upper():
            group_name = "Investment Grade Corporates"
            search_ticker = "LQD"
        elif "MUNI" in bond_type.upper():
            group_name = "Municipal Bonds"
            search_ticker = "MUB"
        else:
            group_name = "Broad Bond Market"
            search_ticker = "AGG"

        group_key = f"broader:{group_name.lower()}"

        if group_key not in group_map:
            group_map[group_key] = {
                "group_type": "Broader Market",
                "news_level": "BROADER_IMPLICATION",
                "group_key": group_key,
                "group_name": group_name,
                "search_ticker": search_ticker,
                "search_term": group_name,
                "bonds": [],
            }

        group_map[group_key]["bonds"].append(bond)

    return list(group_map.values())


def create_search_groups(inventory: list[dict]) -> list[dict]:
    issuer_groups = create_issuer_groups(inventory)
    sector_groups = create_sector_groups(inventory)
    broader_groups = create_broader_market_groups(inventory)

    return issuer_groups + sector_groups + broader_groups


# ============================================================
# 4. News retrieval
# ============================================================

def fetch_news_for_group(group: dict) -> list[dict]:
    """
    Pulls Yahoo/yfinance headlines for one search group.

    For ticker groups, use the ticker.
    For non-ticker groups, yfinance may be weak. This still gives the system
    a working placeholder until a better approved news source is connected.
    """

    headlines = []
    search_ticker = group.get("search_ticker", "")
    search_term = group.get("search_term", "")

    ticker_to_search = search_ticker

    if not ticker_to_search and search_term:
        ticker_to_search = search_term.split()[0]

    if not ticker_to_search:
        return headlines

    try:
        news = yf.Ticker(ticker_to_search).news or []

        for item in news[:MAX_HEADLINES_PER_GROUP]:
            headline = clean_text(item.get("title", ""))
            source = clean_text(item.get("publisher", ""))
            url = clean_text(item.get("link", ""))

            if headline:
                headlines.append(
                    {
                        "headline": headline,
                        "source": source,
                        "url": url,
                    }
                )

    except Exception as error:
        print(f"  [News] error for {group['group_name']}: {error}")

    return headlines


# ============================================================
# 5. Claude classification
# ============================================================

def create_anthropic_client() -> anthropic.Anthropic | None:
    if not ANTHROPIC_API_KEY:
        print("No ANTHROPIC_API_KEY found. Running without Claude verification.")
        return None

    return anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)


def classify_without_claude(group: dict, headlines: list[dict]) -> list[dict]:
    """
    Strict fallback mode.

    Without Claude, do not treat candidate headlines as relevant.
    This prevents false matches when the relevance engine is unavailable.

    The dashboard/backend connection can still be tested because the script
    will produce a "NO RELEVANT NEWS FOUND" summary row if nothing is verified.
    """

    return []


def classify_with_claude(
    group: dict,
    headlines: list[dict],
    client: anthropic.Anthropic,
) -> list[dict]:
    if not headlines:
        return []

    news_level = group["news_level"]
    group_type = group["group_type"]
    group_name = group["group_name"]
    affected_count = len(group["bonds"])
    books = unique_join([bond["book"] for bond in group["bonds"]])
    sectors = unique_join([bond["sector"] for bond in group["bonds"]])
    bond_types = unique_join([bond["bond_type"] for bond in group["bonds"]])
    sample_cusips_text = sample_values([bond["cusip"] for bond in group["bonds"]])

    if news_level == "DIRECT_ISSUER":
        relevance_instruction = (
            "Classify a headline as DIRECT_ISSUER only if it directly relates to "
            "the issuer, borrower, obligor, ticker, debt, business, financials, "
            "rating, litigation, or closely related entity. Otherwise use NOT_RELEVANT."
        )
    elif news_level == "SECTOR":
        relevance_instruction = (
            "Classify a headline as SECTOR only if it clearly affects the sector "
            "of the affected bonds. Otherwise use NOT_RELEVANT."
        )
    else:
        relevance_instruction = (
            "Classify a headline as BROADER_IMPLICATION only if it has a clear "
            "macro, rates, credit, municipal, fund-flow, tax, policy, or spread "
            "implication for this bond group. Otherwise use NOT_RELEVANT."
        )

    headline_list = "\n".join(
        f"{index + 1}. {headline['headline']}"
        for index, headline in enumerate(headlines)
    )

    prompt = (
        "You are reviewing news relevance for a fixed income inventory tool.\n\n"
        "Your job is only to decide whether each headline is relevant to the bond group.\n"
        "Do not make a buy or sell recommendation.\n"
        "Do not call anything attractive, cheap, rich, suitable, or a good idea.\n"
        "Do not invent facts.\n"
        "If the connection is weak, use NOT_RELEVANT.\n\n"
        f"Group type: {group_type}\n"
        f"Target relevance label: {news_level}\n"
        f"Group name: {group_name}\n"
        f"Books affected: {books}\n"
        f"Sectors affected: {sectors}\n"
        f"Bond types affected: {bond_types}\n"
        f"Affected bond count: {affected_count}\n"
        f"Sample CUSIPs: {sample_cusips_text}\n\n"
        f"{relevance_instruction}\n\n"
        f"Headlines:\n{headline_list}\n\n"
        "Respond in this exact format for each headline, one per line:\n"
        "N | CLASSIFICATION | one sentence implication\n\n"
        "Allowed classifications:\n"
        "DIRECT_ISSUER, SECTOR, BROADER_IMPLICATION, NOT_RELEVANT\n\n"
        "Example:\n"
        "1 | SECTOR | The headline may matter because transportation revenue credits can be sensitive to traffic trends."
    )

    try:
        message = client.messages.create(
            model="claude-opus-4-6",
            max_tokens=1200,
            messages=[
                {
                    "role": "user",
                    "content": prompt,
                }
            ],
        )

        response_text = message.content[0].text

    except Exception as error:
        print(f"  [Claude] API error for {group_name}: {error}")
        return []

    classified = []

    for line in response_text.strip().split("\n"):
        parts = [part.strip() for part in line.split("|")]

        if len(parts) < 3:
            continue

        try:
            index = int(parts[0]) - 1
            classification = parts[1].upper()
            implication = parts[2]

            if index < 0 or index >= len(headlines):
                continue

            if classification == "NOT_RELEVANT":
                continue

            if classification not in ("DIRECT_ISSUER", "SECTOR", "BROADER_IMPLICATION"):
                continue

            item = headlines[index].copy()
            item["classification"] = classification
            item["implication"] = implication
            classified.append(item)

        except ValueError:
            continue

    return classified


def classify_headlines_for_group(
    group: dict,
    headlines: list[dict],
    client: anthropic.Anthropic | None,
) -> list[dict]:
    if not headlines:
        return []

    if client is None:
        return classify_without_claude(group, headlines)

    return classify_with_claude(group, headlines, client)


# ============================================================
# 6. Output rows
# ============================================================

def build_group_output_rows(group: dict, classified_headlines: list[dict]) -> list[dict]:
    rows = []

    books = unique_join([bond["book"] for bond in group["bonds"]])
    sectors = unique_join([bond["sector"] for bond in group["bonds"]])
    bond_types = unique_join([bond["bond_type"] for bond in group["bonds"]])
    ratings = unique_join([bond["rating"] for bond in group["bonds"]])
    sample_cusips_text = sample_values([bond["cusip"] for bond in group["bonds"]])
    affected_bond_count = len(group["bonds"])

    for headline in classified_headlines:
        rows.append(
            {
                "book": books,
                "news_level": headline.get("classification", group["news_level"]),
                "group_type": group["group_type"],
                "issuer_sector_or_group": group["group_name"],
                "affected_bond_count": affected_bond_count,
                "sample_cusips": sample_cusips_text,
                "sector": sectors,
                "bond_type": bond_types,
                "rating": ratings,
                "headline": headline.get("headline", ""),
                "implication": headline.get("implication", ""),
                "source": headline.get("source", ""),
                "url": headline.get("url", ""),
                "run_date": datetime.today().strftime("%Y-%m-%d"),
            }
        )

    return rows


def build_no_news_rows(search_groups: list[dict]) -> list[dict]:
    """
    If no relevant headlines are found anywhere, return one summary row.
    """

    total_bonds = 0
    books = []
    sectors = []
    cusips = []

    for group in search_groups:
        if group["group_type"] == "Issuer":
            for bond in group["bonds"]:
                total_bonds += 1
                books.append(bond["book"])
                sectors.append(bond["sector"])
                cusips.append(bond["cusip"])

    return [
        {
            "book": unique_join(books),
            "news_level": "NOT_RELEVANT",
            "group_type": "Full Inventory",
            "issuer_sector_or_group": "All searched groups",
            "affected_bond_count": total_bonds,
            "sample_cusips": sample_values(cusips),
            "sector": unique_join(sectors),
            "bond_type": "",
            "rating": "",
            "headline": "NO RELEVANT NEWS FOUND",
            "implication": "No relevant direct issuer, sector, or broader market news was verified by the current prototype search.",
            "source": "",
            "url": "",
            "run_date": datetime.today().strftime("%Y-%m-%d"),
        }
    ]


# ============================================================
# 7. Main capability
# ============================================================

def run_capability1() -> None:
    print("")
    print("Running Tool 1 - Capability 1: Bonds With Recent News")

    client = create_anthropic_client()

    inventory = normalize_inventory(load_inventory())
    search_groups = create_search_groups(inventory)

    print(f"Loaded {len(inventory)} bonds.")
    print(f"Created {len(search_groups)} search groups.")
    print("This avoids searching one time per CUSIP.")

    results = []

    for group in search_groups:
        print("")
        print(
            f"Searching group: {group['group_type']} | "
            f"{group['group_name']} | "
            f"{len(group['bonds'])} affected bond(s)"
        )

        headlines = fetch_news_for_group(group)
        print(f"  Candidate headlines found: {len(headlines)}")

        classified_headlines = classify_headlines_for_group(group, headlines, client)

        if classified_headlines:
            print(f"  Relevant headline(s): {len(classified_headlines)}")
            results.extend(build_group_output_rows(group, classified_headlines))
        else:
            print("  No relevant headlines for this group.")

        time.sleep(0.25)

    if not results:
        results = build_no_news_rows(search_groups)

    fieldnames = [
        "book",
        "news_level",
        "group_type",
        "issuer_sector_or_group",
        "affected_bond_count",
        "sample_cusips",
        "sector",
        "bond_type",
        "rating",
        "headline",
        "implication",
        "source",
        "url",
        "run_date",
    ]

    with open(OUTPUT_FILE, "w", newline="", encoding="utf-8") as file:
        writer = csv.DictWriter(file, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(results)

    print("")
    print(f"Done - {len(results)} grouped row(s) written to {OUTPUT_FILE}")


if __name__ == "__main__":
    run_capability1()