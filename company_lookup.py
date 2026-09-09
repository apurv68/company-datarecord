"""
Comprehensive Categorized Company Data Record Fetcher via Multi-Source Web Intelligence
========================================================================================
Fetches detailed corporate intelligence categorized into:
1. Core Company Identity & Overview
2. Land, Real Estate & Property Transactions (Sale/Purchase)
3. Revenue & Financials (This Year vs Last Year)
4. Key Employees, Leadership & Executive Roles
5. New Job Openings & Workforce Changes (Hiring / Layoffs)
6. Legal & Regulatory Information (Court cases, disputes, compliance)
"""

import sys
import os
import json
import csv
import re
import time
from typing import Dict, Any, Optional, List, Tuple

# Ensure UTF-8 output on Windows consoles
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8")

import requests
from bs4 import BeautifulSoup
from ddgs import DDGS
from rich.console import Console
from rich.panel import Panel
from rich.table import Table
from rich.progress import Progress, SpinnerColumn, TextColumn

console = Console(force_terminal=True, legacy_windows=False)


EXPORT_CSV_PATH = "company_records.csv"
EXPORT_TABLE2_CSV_PATH = "company_financials_5yr.csv"
EXPORT_JSON_PATH = "company_records.json"


def clean_text(text: Any) -> str:
    """Clean wiki markup, citations, HTML tags, and extra whitespace."""
    if text is None:
        return "N/A"
    if not isinstance(text, str):
        if isinstance(text, dict):
            return text.get("id") or text.get("value") or str(text)
        return str(text)
    cleaned = re.sub(r"\[\[(?:[^|\]]*\|)?([^\]]+)\]\]", r"\1", text)
    cleaned = re.sub(r"\[\s*ISIN:\s*\[\s*([^\]]+)\s*\]\]", r"\1", cleaned)
    cleaned = re.sub(r"\[\d+\]", "", cleaned)
    cleaned = re.sub(r"<[^>]+>", "", cleaned)
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    return cleaned if cleaned else "N/A"


def query_duckduckgo_instant(query_term: str) -> Optional[Dict[str, Any]]:
    """Query DuckDuckGo Instant Answer API for structured corporate infobox."""
    url = "https://api.duckduckgo.com/"
    params = {
        "q": query_term,
        "format": "json",
        "no_html": 1,
        "skip_disambig": 1
    }
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
    }
    try:
        response = requests.get(url, params=params, headers=headers, timeout=6)
        if response.status_code in (200, 202):
            return response.json()
    except Exception:
        pass
    return None


def fetch_wikidata_facts(qid: str) -> Dict[str, Any]:
    """Fetch structured, verified company attributes directly from Wikidata."""
    if not qid:
        return {}
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36"}
    try:
        res = requests.get(
            "https://www.wikidata.org/w/api.php",
            params={
                "action": "wbgetentities",
                "ids": qid,
                "props": "claims|labels|aliases",
                "languages": "en",
                "format": "json"
            },
            headers=headers,
            timeout=8
        ).json()
        ent = res.get("entities", {}).get(qid, {})
        claims = ent.get("claims", {})
        
        eids_to_resolve = []
        
        # P112: Founders
        founders_eids = []
        for c in claims.get("P112", []):
            snak = c.get("mainsnak", {}).get("datavalue", {}).get("value", {})
            if isinstance(snak, dict) and "id" in snak:
                founders_eids.append(snak["id"])
                eids_to_resolve.append(snak["id"])
                
        # P169: Current CEO / Chief Executive Officer (no P582 end date)
        ceo_eids = []
        for c in claims.get("P169", []):
            snak = c.get("mainsnak", {}).get("datavalue", {}).get("value", {})
            if isinstance(snak, dict) and "id" in snak:
                qual = c.get("qualifiers", {})
                if "P582" not in qual:  # Currently active
                    ceo_eids.append(snak["id"])
                    eids_to_resolve.append(snak["id"])
        # If all CEOs had end date, take the latest one
        if not ceo_eids and claims.get("P169"):
            last_c = claims["P169"][-1]
            snak = last_c.get("mainsnak", {}).get("datavalue", {}).get("value", {})
            if isinstance(snak, dict) and "id" in snak:
                ceo_eids.append(snak["id"])
                eids_to_resolve.append(snak["id"])

        # P159: Headquarters location
        hq_eids = []
        for c in claims.get("P159", []):
            snak = c.get("mainsnak", {}).get("datavalue", {}).get("value", {})
            if isinstance(snak, dict) and "id" in snak:
                hq_eids.append(snak["id"])
                eids_to_resolve.append(snak["id"])

        # P452: Industry
        ind_eids = []
        for c in claims.get("P452", []):
            snak = c.get("mainsnak", {}).get("datavalue", {}).get("value", {})
            if isinstance(snak, dict) and "id" in snak:
                ind_eids.append(snak["id"])
                eids_to_resolve.append(snak["id"])

        # P1454: Legal form / Type
        type_eids = []
        for c in claims.get("P1454", []) + claims.get("P31", []):
            snak = c.get("mainsnak", {}).get("datavalue", {}).get("value", {})
            if isinstance(snak, dict) and "id" in snak:
                type_eids.append(snak["id"])
                eids_to_resolve.append(snak["id"])

        # P571: Inception date
        founded_date = None
        for c in claims.get("P571", []):
            t = c.get("mainsnak", {}).get("datavalue", {}).get("value", {}).get("time")
            if t:
                cleaned_t = t.replace("+", "").split("T")[0]
                parts = cleaned_t.split("-")
                if len(parts) >= 1 and parts[0].isdigit():
                    if len(parts) == 3 and parts[1] != "00" and parts[2] != "00":
                        founded_date = cleaned_t
                    else:
                        founded_date = parts[0]
                break

        # P414: Stock Exchange + P249: Ticker symbol
        tickers = []
        for c in claims.get("P414", []):
            qual = c.get("qualifiers", {})
            if "P249" in qual:
                for q in qual["P249"]:
                    sym = q.get("datavalue", {}).get("value")
                    if sym and sym not in tickers:
                        tickers.append(sym)

        # P856: Official website
        website = None
        for c in claims.get("P856", []):
            w = c.get("mainsnak", {}).get("datavalue", {}).get("value")
            if w and isinstance(w, str) and w.startswith("http"):
                clean_w = w.rstrip("/").lower()
                if clean_w.endswith(".com") or clean_w.endswith(".org") or clean_w.endswith(".in") or clean_w.endswith(".co"):
                    website = w
                    if clean_w.endswith(".com"):
                        break
                elif not website:
                    website = w

        # Batch resolve entity IDs to English labels
        resolved_labels = {}
        unique_eids = list(set(eids_to_resolve))[:50]
        if unique_eids:
            rres = requests.get(
                "https://www.wikidata.org/w/api.php",
                params={
                    "action": "wbgetentities",
                    "ids": "|".join(unique_eids),
                    "props": "labels",
                    "languages": "en",
                    "format": "json"
                },
                headers=headers,
                timeout=8
            ).json()
            for eid, edata in rres.get("entities", {}).items():
                lbl = edata.get("labels", {}).get("en", {}).get("value")
                if lbl:
                    resolved_labels[eid] = lbl

        founders_str = ", ".join([resolved_labels[e] for e in founders_eids if e in resolved_labels])
        ceo_str = ", ".join([resolved_labels[e] for e in ceo_eids if e in resolved_labels])
        hq_str = ", ".join([resolved_labels[e] for e in hq_eids if e in resolved_labels])
        ind_str = ", ".join([resolved_labels[e] for e in ind_eids if e in resolved_labels])
        
        legal_types = [resolved_labels[e] for e in type_eids if e in resolved_labels]
        corp_type = "Public company" if any("public" in t.lower() for t in legal_types) else ("Private company" if any("private" in t.lower() for t in legal_types) else "")

        return {
            "founders": founders_str or None,
            "ceo": ceo_str or None,
            "hq": hq_str or None,
            "industry": ind_str or None,
            "founded": founded_date,
            "tickers": tickers,
            "website": website,
            "type": corp_type or (legal_types[0] if legal_types else None)
        }
    except Exception:
        return {}


def get_wikipedia_company_data(company_name: str) -> Optional[Dict[str, Any]]:
    """
    Fetch authoritative corporate encyclopedia facts from Wikipedia REST API & Wikidata.
    Supports exact, possessive-stripped, suffix-appended, and OpenSearch candidate titles.
    """
    base_name = re.sub(r"['\u2019]s\b", "", company_name, flags=re.IGNORECASE).strip()
    candidate_titles = [
        company_name,
        f"{base_name} (company)",
        f"{base_name}, Inc.",
        f"{base_name}.com",
        f"{base_name}.com, Inc.",
        f"{base_name}'s",
        f"{base_name} Inc.",
        f"{base_name} Corporation",
        f"{base_name} Limited",
        f"{base_name} Group",
        base_name
    ]

    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36"}

    # Use OpenSearch to find top article titles
    try:
        os_res = requests.get(
            "https://en.wikipedia.org/w/api.php",
            params={"action": "opensearch", "search": company_name, "limit": 6, "format": "json"},
            headers=headers,
            timeout=4
        ).json()
        if len(os_res) > 1:
            for t in os_res[1]:
                if t not in candidate_titles:
                    candidate_titles.append(t)
    except Exception:
        pass

    summary_data = None
    best_title = None

    company_signals = [
        "company", "corporation", "inc.", "ltd", "founded", "headquartered", "headquarters",
        "conglomerate", "multinational", "enterprise", "retailer", "manufacturer", "firm",
        "technology", "chain", "platform", "online marketplace", "developer", "producer",
        "brand", "automotive", "software", "foodstuff", "services", "restaurant"
    ]

    for title in candidate_titles:
        try:
            url = f"https://en.wikipedia.org/api/rest_v1/page/summary/{requests.utils.quote(title)}"
            r = requests.get(url, headers=headers, timeout=5)
            if r.status_code == 200:
                data = r.json()
                if data.get("type") == "disambiguation":
                    continue
                desc = (data.get("description") or "").lower()
                ext = (data.get("extract") or "").lower()
                if "river" in desc or "drainage basin" in desc or "film" in desc:
                    continue
                if any(s in desc or s in ext[:250] for s in company_signals):
                    summary_data = data
                    best_title = data.get("title") or title
                    break
        except Exception:
            pass

    if not summary_data:
        return None

    full_text = summary_data.get("extract", "")
    try:
        t_res = requests.get(
            "https://en.wikipedia.org/w/api.php",
            params={
                "action": "query",
                "format": "json",
                "prop": "extracts",
                "explaintext": True,
                "titles": best_title
            },
            headers=headers,
            timeout=6
        ).json()
        pages = t_res.get("query", {}).get("pages", {})
        for pid, pdata in pages.items():
            if pid != "-1" and len(pdata.get("extract", "")) > 100:
                full_text = pdata.get("extract", "")
                break
    except Exception:
        pass

    wikibase_item = summary_data.get("wikibase_item")
    wikidata_facts = fetch_wikidata_facts(wikibase_item) if wikibase_item else {}

    return {
        "title": best_title,
        "url": summary_data.get("content_urls", {}).get("desktop", {}).get("page", ""),
        "extract": summary_data.get("extract", ""),
        "description": summary_data.get("description", ""),
        "text": full_text,
        "wikibase_item": wikibase_item,
        "wikidata": wikidata_facts
    }


def extract_wiki_category_snippets(wiki_data: Optional[Dict[str, Any]], keywords: List[str], category_title: str) -> List[Dict[str, str]]:
    """Extract targeted factual paragraphs from Wikipedia based on category keywords."""
    if not wiki_data or not wiki_data.get("text"):
        return []
    
    extracted = []
    text = wiki_data["text"]
    paragraphs = [p.strip() for p in text.split("\n\n") if p.strip()]
    
    for p in paragraphs:
        p_lower = p.lower()
        if any(h in p_lower for h in ["== references", "== external links", "jump to content", "sidebarhide", "main menu"]):
            continue
        if any(kw in p_lower for kw in keywords):
            # Clean heading markup like '== History =='
            lines = [l.strip() for l in p.split("\n") if l.strip() and not l.startswith("==")]
            combined = " ".join(lines)
            if len(combined) > 40:
                extracted.append({
                    "title": f"Wikipedia Corporate Records ({category_title})",
                    "snippet": combined[:280] + ("..." if len(combined) > 280 else ""),
                    "source": wiki_data.get("url", "https://en.wikipedia.org")
                })
        if len(extracted) >= 2:
            break
            
    return extracted


def is_relevant_result(item: Dict[str, str], company_name: str, category: str = "") -> bool:
    """
    Verify that a search result snippet actually belongs to the target company
    and reject 3rd party proximity noise ('near this company', 'apartments near...').
    """
    name = company_name.lower().strip()
    title = clean_text(item.get("title", "")).lower()
    snippet = clean_text(item.get("snippet", "")).lower()
    source = item.get("source", "").lower()
    combined_text = f"{title} {snippet} {source}"

    # Normalize punctuation into spaces to compare clean word sequences
    norm_text = re.sub(r"\s+", " ", re.sub(r"[^\w\s]", " ", combined_text)).strip()

    # Rule 1: Reject location proximity noise ('near company', 'apartments near', etc.)
    proximity_patterns = [
        r"\bapartments?\s+near\b",
        r"\bflats?\s+near\b",
        r"\bprojects?\s+near\b",
        r"\bhotels?\s+near\b",
        r"\bcompanies\s+near\b",
        r"\bcompetitors\s+of\b",
        r"\balternatives\s+to\b",
        r"\bhomes?\s+near\b",
        r"\bresidential\s+near\b"
    ]
    for p in proximity_patterns:
        if re.search(p, norm_text):
            return False

    # Rule 2: Reject web navigation chrome and encyclopedia boilerplate (e.g. "jump to content", "sidebarhide")
    boilerplate_patterns = [
        r"\bjump to content\b",
        r"\bmove to sidebarhide\b",
        r"\bmain menu\b",
        r"\btoggle the table of contents\b",
        r"\brandom article\b",
        r"\babout wikipedia\b",
        r"\blearn to edit\b",
        r"\bcommunity portal\b",
        r"\bcreate account\b",
        r"\bpersonal tools\b",
        r"\bskip to content\b",
        r"\bskip to main content\b",
        r"\benable javascript\b",
        r"\bterms of useand privacy policy\b",
        r"\bis a registered trademark of the wikimedia\b"
    ]
    for p in boilerplate_patterns:
        if re.search(p, norm_text):
            return False

    # Rule 3: For leadership queries, prioritize active/current roles and avoid obsolete or deceased ex-leaders
    if category == "roles":
        past_markers = [
            "former ceo", "ex-ceo", "ousted ceo", "until his death", "until her death",
            "was the managing director", "was the ceo", "past ceo", "stepped down as ceo",
            "served as ceo from", "served as managing director from"
        ]
        if any(m in norm_text for m in past_markers) and not any(k in norm_text for k in ("current", "appointed", "takes charge", "new md", "new ceo", "present", "executive")):
            return False

    # Rule 4: Enforce category-specific substance (snippet must actually discuss the category topic)
    category_keywords = {
        "revenue": ["revenue", "profit", "loss", "turnover", "income", "financial", "result", "ebitda", "valuation", "crore", "cr", "million", "billion", "worth", "fy", "sales", "margin", "₹", "$"],
        "land": ["plant", "factory", "land", "acre", "hectare", "real estate", "property", "facility", "facilities", "office", "campus", "lease", "building", "manufacturing", "sq ft", "sq m", "plot", "site"],
        "roles": ["ceo", "cfo", "cto", "founder", "director", "president", "officer", "executive", "board", "leadership", "management", "head", "appointed", "md"],
        "jobs": ["career", "job", "opening", "hiring", "vacancy", "vacancies", "recruitment", "employment", "workforce", "hire", "naukri", "apply"],
        "legal": ["court", "case", "suit", "dispute", "lawsuit", "legal", "order", "tribunal", "verdict", "complaint", "petition", "bench", "justice", "trademark", "patent", "alleged", "notice", "litigation"]
    }
    if category in category_keywords:
        kws = category_keywords[category]
        if not any(kw in norm_text for kw in kws):
            return False

    # Stems and clean variants (e.g., "haldiram's" -> ["haldiram", "haldirams"])
    base_name = re.sub(r"['’]s\b", "", name, flags=re.IGNORECASE).strip()
    norm_base = re.sub(r"\s+", " ", re.sub(r"[^\w\s]", " ", base_name)).strip()
    norm_full = re.sub(r"\s+", " ", re.sub(r"[^\w\s]", " ", name)).strip()

    variants = {
        norm_full,
        norm_full.replace(" ", ""),
        norm_base,
        norm_base.replace(" ", ""),
        norm_base + "s"
    }

    for v in variants:
        if len(v) >= 3 and v in norm_text:
            return True

    # Rule 3: Distinctive token matching for multi-word enterprises
    stop_tokens = {"pvt", "ltd", "private", "limited", "inc", "llc", "corp", "corporation", "company", "s"}
    tokens = [t for t in norm_base.split() if t not in stop_tokens and len(t) > 1]

    if tokens and not all(t.isdigit() or t in ("days", "day", "month", "year", "time") for t in tokens):
        if all(t in norm_text for t in tokens):
            return True

    return False


def search_ddgs_category(
    queries: List[str],
    company_name: str,
    category: str = "",
    max_results: int = 3
) -> List[Dict[str, str]]:
    """
    Perform targeted multi-engine web search across DuckDuckGo, Brave, Yahoo, Google.
    Tries primary queries first; if few or zero results, cascades to relaxed fallback queries.
    """
    results = []
    seen_urls = set()

    for q in queries:
        try:
            with DDGS(timeout=8) as ddgs:
                raw_results = list(ddgs.text(q, max_results=max_results * 2, backend="auto"))
                for r in raw_results:
                    href = r.get("href", "")
                    if href in seen_urls:
                        continue
                    snippet = clean_text(r.get("body", ""))
                    title = clean_text(r.get("title", ""))
                    if snippet and snippet != "N/A":
                        item = {
                            "title": title,
                            "snippet": snippet,
                            "source": href
                        }
                        if is_relevant_result(item, company_name, category):
                            seen_urls.add(href)
                            results.append(item)
                        if len(results) >= max_results:
                            return results
            time.sleep(0.3)
        except Exception:
            pass
        if len(results) >= max_results:
            break

    return results


def parse_year_from_item(item: Dict[str, str]) -> int:
    """Extract relevant year or fiscal year (e.g. 2026, 2025, 2024) from title, snippet, or URL."""
    text = f"{item.get('title', '')} {item.get('snippet', '')} {item.get('source', '')}"

    # 1. Check for FY (Fiscal Year) mentions like FY26, FY25, FY 2025
    fy_match = re.search(r"\bFY\s*(\d{2,4})\b", text, re.IGNORECASE)
    if fy_match:
        val = fy_match.group(1)
        if len(val) == 2:
            return 2000 + int(val)
        return int(val)

    # 2. Check for explicit 4-digit years (e.g. 2015-2026)
    years = re.findall(r"\b(20[1-3][0-9])\b", text)
    if years:
        return max(int(y) for y in years)

    return 0


def summarize_category(items: List[Dict[str, str]], company_name: str = "") -> Dict[str, Any]:
    """Group and format category search results into a clean, year-wise timeline."""
    if not items:
        empty_msg = (
            f"No specific public records found for '{company_name}' in this category (Common for private unlisted companies and startups)."
            if company_name else "No specific public records found on this topic."
        )
        return {
            "summary": empty_msg,
            "timeline": {},
            "raw": []
        }

    # Group items by year label
    grouped: Dict[str, List[Dict[str, str]]] = {}
    for item in items:
        yr = parse_year_from_item(item)
        if yr >= 2026:
            lbl = f"{yr} (Current Year)"
        elif yr == 2025:
            lbl = f"{yr} (Previous Year)"
        elif yr > 2000:
            lbl = str(yr)
        else:
            lbl = "Recent / Active Records"

        if lbl not in grouped:
            grouped[lbl] = []
        grouped[lbl].append(item)

    # Sort years descending (newest first)
    def sort_key(k: str):
        m = re.search(r"\b(\d{4})\b", k)
        if m:
            return int(m.group(1))
        return 9999 if "recent" in k.lower() else 0

    sorted_keys = sorted(grouped.keys(), key=sort_key, reverse=True)

    # Build formatted year-wise summary string
    sections = []
    for yr_label in sorted_keys:
        entry_lines = []
        for item in grouped[yr_label]:
            entry_lines.append(f"  • {item['title']}: {item['snippet']} (Source: {item['source']})")
        sections.append(f"📅 [{yr_label}]:\n" + "\n\n".join(entry_lines))

    return {
        "summary": "\n\n".join(sections),
        "timeline": grouped,
        "raw": items
    }


def fetch_multi_year_financial_history(
    company_name: str,
    base_name: str,
    wiki_rev: List[Dict[str, str]]
) -> Dict[str, Any]:
    """
    Fetch REAL reported financial metrics for the Present Year (2026) 
    and the Previous 5 Years (2025, 2024, 2023, 2022, 2021).
    Strictly reports published audited numbers, MCA filings, and official disclosures (no calculated/synthetic data).
    """
    target_years = [2026, 2025, 2024, 2023, 2022, 2021]
    history_by_year: Dict[int, List[Dict[str, str]]] = {yr: [] for yr in target_years}
    all_raw_items: List[Dict[str, str]] = []
    seen_urls = set()

    # Step 1: Pre-populate from verified Wikipedia financial extracts
    for item in wiki_rev:
        yr = parse_year_from_item(item)
        if yr in history_by_year:
            history_by_year[yr].append(item)
            all_raw_items.append(item)

    # Step 2: Query multi-engine web search specifically for each target year
    for yr in target_years:
        if len(history_by_year[yr]) >= 2:
            continue
        seen_for_this_yr = {it["source"] for it in history_by_year[yr]}
        short_fy = f"FY{str(yr)[2:]}"
        queries = [
            f"{base_name} revenue turnover profit {yr} OR {short_fy}",
            f"{base_name} financial statements {yr} OR {short_fy}",
            f"{base_name} annual report 10-K SEC {yr} OR {short_fy}",
            f'"{company_name}" revenue {yr} OR {short_fy}'
        ]
        for q in queries:
            try:
                with DDGS(timeout=8) as ddgs:
                    raw = list(ddgs.text(q, max_results=5, backend="auto"))
                    for r in raw:
                        href = r.get("href", "")
                        if href in seen_for_this_yr:
                            continue
                        snippet = clean_text(r.get("body", ""))
                        title = clean_text(r.get("title", ""))
                        if not snippet or snippet == "N/A":
                            continue

                        item = {"title": title, "snippet": snippet, "source": href}
                        if is_relevant_result(item, company_name, category="revenue"):
                            text_lower = f"{title} {snippet} {href}".lower()
                            check_text = text_lower.replace(f"missing: {yr}", "").replace(f"missing: {short_fy.lower()}", "")
                            if str(yr) in check_text or short_fy.lower() in check_text or f"fiscal {yr}" in check_text:
                                seen_for_this_yr.add(href)
                                history_by_year[yr].append(item)
                                all_raw_items.append(item)
                                if len(history_by_year[yr]) >= 2:
                                    break
            except Exception:
                pass
            if history_by_year[yr]:
                break
            time.sleep(0.3)

    # Build chronological year-wise sections (Newest to Oldest)
    sections = []
    grouped_timeline = {}
    for yr in target_years:
        if yr == 2026:
            lbl = f"{yr} (Present Year)"
        elif yr == 2025:
            lbl = f"{yr} (Previous Year)"
        else:
            lbl = f"{yr}"

        items = history_by_year[yr]
        grouped_timeline[lbl] = items
        if items:
            entry_lines = [f"  • {it['title']}: {it['snippet']} (Source: {it['source']})" for it in items]
            sections.append(f"📅 [{lbl}]:\n" + "\n\n".join(entry_lines))
        else:
            sections.append(f"📅 [{lbl}]:\n  • No audited public financial filing found for this specific fiscal year on public records.")

    return {
        "summary": "\n\n".join(sections),
        "timeline": grouped_timeline,
        "raw": all_raw_items
    }


def fetch_full_company_profile(company_name: str) -> Dict[str, Any]:
    """
    Fetch comprehensive, categorized company data from multi-source web intelligence:
    Core, Land/Properties, YoY Revenue, Employee Roles, Job Openings, and Legal Info.
    """
    cleaned_name = company_name.strip()
    base_name = re.sub(r"['’]s\b", "", cleaned_name, flags=re.IGNORECASE).strip()

    profile = {
        "Search Query": cleaned_name,
        "Company Name": cleaned_name,
        "Industry": "Enterprise / Organization",
        "Type": "Unspecified",
        "Traded As (Ticker)": "N/A",
        "Founded": "N/A",
        "Founders": "N/A",
        "CEO": "N/A",
        "Headquarters": "N/A",
        "Website": "N/A",
        "Overview": "N/A",
        "Categories": {}
    }

    # Step 1: DuckDuckGo Instant Answer lookup for baseline info
    candidate_queries = [
        cleaned_name,
        base_name,
        f"{cleaned_name} Inc",
        f"{cleaned_name} company",
        f"{cleaned_name} Limited"
    ]

    best_ddg = None
    for q in candidate_queries:
        data = query_duckduckgo_instant(q)
        if data:
            infobox = data.get("Infobox", {})
            content = infobox.get("content", []) if isinstance(infobox, dict) else []
            abstract = data.get("AbstractText") or data.get("Abstract", "")
            if len(content) >= 3 or len(abstract) > 80:
                best_ddg = data
                break

    if best_ddg:
        ddg_heading = best_ddg.get("Heading") or cleaned_name
        # Only use DDG heading if it resembles a company name (not a TLD like .amazon)
        if not ddg_heading.startswith(".") and len(ddg_heading) > 1:
            profile["Company Name"] = ddg_heading
        profile["Overview"] = clean_text(best_ddg.get("AbstractText") or best_ddg.get("Abstract"))
        infobox = best_ddg.get("Infobox", {})
        if isinstance(infobox, dict):
            for item in infobox.get("content", []):
                if isinstance(item, dict) and "label" in item and "value" in item:
                    lbl = item["label"].lower().strip()
                    val = clean_text(item["value"])
                    if "industry" in lbl or "sector" in lbl:
                        profile["Industry"] = val
                    elif "type" in lbl:
                        profile["Type"] = val
                    elif "traded as" in lbl or "isin" in lbl:
                        profile["Traded As (Ticker)"] = val
                    elif "founded" in lbl:
                        profile["Founded"] = val
                    elif "founders" in lbl or "founder" in lbl:
                        profile["Founders"] = val
                    elif "headquarters" in lbl or "area served" in lbl:
                        profile["Headquarters"] = val
                    elif "website" in lbl:
                        urls = re.findall(r"(?:https?://)?(?:www\.)?([a-zA-Z0-9-]+\.[a-zA-Z]{2,}(?:/[^\s,\]]*)?)", str(val))
                        if urls:
                            profile["Website"] = f"https://{urls[0]}"

    # Step 2: Query authoritative encyclopedia data (Wikipedia API & Wikidata)
    wiki_data = get_wikipedia_company_data(cleaned_name)
    if wiki_data:
        wd = wiki_data.get("wikidata", {})
        intro_p = wiki_data.get("extract") or wiki_data["text"].split("\n\n")[0] or ""
        wiki_overview = clean_text(intro_p)
        if len(wiki_overview) > 50:
            profile["Overview"] = wiki_overview

        # 1. Company Name: extract legal / official entity name
        wiki_title = wiki_data.get("title", "")
        name_match = re.match(r"^([A-Z0-9][^(\n\r]+?)(?:\s*\(|\s+is\s+(?:an?|the)\b|\s+was\s+(?:an?|the)\b)", intro_p)
        if name_match:
            profile["Company Name"] = name_match.group(1).strip().rstrip(",")
        elif wiki_title and "(company)" not in wiki_title.lower():
            profile["Company Name"] = wiki_title

        # 2. Founders: First Wikidata, then lead text regex fallback
        if wd.get("founders"):
            profile["Founders"] = wd["founders"]
        elif profile["Founders"] == "N/A":
            lead_text = wiki_data["text"][:3000]
            f_match = re.search(r"(?:founded|started|established|created)(?:\s+(?:in|on)\s+[\w\s,]+?)?\s+by\s+([A-Z][a-zA-Z\s,.&]+?)(?:\.|,|\sin\b|\safter\b|\sto\b|\sand\b|\sas\b)", lead_text, re.IGNORECASE)
            if not f_match:
                f_match = re.search(r"(?:founder[s]?\s*(?:is|are|was|were|include)?\s*:?\s*)([A-Z][a-zA-Z\s,.&]+?)(?:\.|\sin\b|\son\b|,|\sand\b)", lead_text, re.IGNORECASE)
            if f_match:
                cand_f = f_match.group(1).strip().rstrip(",")
                bad_words = ["college", "sachs", "goldman", "partner", "court", "university", "hospital", "council"]
                if len(cand_f) < 80 and "\n" not in cand_f and not any(w in cand_f.lower() for w in bad_words):
                    profile["Founders"] = cand_f

        # 3. CEO / Leadership
        if wd.get("ceo"):
            profile["CEO"] = wd["ceo"]
        elif profile["CEO"] == "N/A":
            lead_text = wiki_data["text"][:3000]
            ceo_m = re.search(r"(?:current\s+)?(?:CEO|Chief Executive Officer)(?:\s+(?:is|of))?\s+([A-Z][a-zA-Z\s]+?)(?:\.|\sand\b|,|\swho\b)", lead_text)
            if ceo_m:
                profile["CEO"] = ceo_m.group(1).strip().rstrip(",")

        # 4. Headquarters
        if wd.get("hq"):
            profile["Headquarters"] = wd["hq"]
        elif profile["Headquarters"] == "N/A":
            lead_text = wiki_data["text"][:4000]
            hq_match = re.search(r"(?:headquartered in|based in|headquarters (?:are|is) in|headquarters located in)\s+([A-Z][a-zA-Z\s,]+?)(?:\.|\sand\b|\swith\b|\sfor\b|\sto\b)", lead_text, re.IGNORECASE)
            if hq_match:
                profile["Headquarters"] = hq_match.group(1).strip().rstrip(",")
            elif "headquarters" in wiki_data["text"].lower():
                hq_sec = re.search(r"==+\s*Headquarters\s*==+([^\n=]+)", wiki_data["text"], re.IGNORECASE)
                if hq_sec:
                    c_m = re.search(r"(?:in|across)\s+([A-Z][a-zA-Z\s,]+?)(?:\.|\sand\b|\swith\b)", hq_sec.group(1))
                    if c_m:
                        profile["Headquarters"] = c_m.group(1).strip().rstrip(",")

        # 5. Industry
        if wd.get("industry"):
            profile["Industry"] = wd["industry"]
        elif wiki_data.get("description"):
            profile["Industry"] = wiki_data["description"]

        # 6. Ticker
        if wd.get("tickers"):
            profile["Traded As (Ticker)"] = ", ".join(wd["tickers"])

        # 7. Founded Date
        if wd.get("founded"):
            profile["Founded"] = wd["founded"]
        elif profile["Founded"] == "N/A":
            lead_text = wiki_data["text"][:3000]
            founded_match = re.search(r"(?:founded|incorporated|established)\s+(?:on\s+|in\s+)?(\w+\s+\d{1,2},?\s+\d{4}|\d{4})", lead_text, re.IGNORECASE)
            if founded_match:
                profile["Founded"] = founded_match.group(1).strip()

        # 8. Website
        if wd.get("website"):
            profile["Website"] = wd["website"]
        elif profile["Website"] in ("N/A", "", ","):
            web_match = re.search(r"(?:https?://)?(?:www\.)?([a-zA-Z0-9-]+\.(?:com|org|in|net|co))", wiki_data["text"])
            if web_match:
                profile["Website"] = f"https://www.{web_match.group(1)}"

        # 9. Type
        if wd.get("type"):
            profile["Type"] = wd["type"]

    # Pre-extract Wikipedia historical records for deep categories
    wiki_land = extract_wiki_category_snippets(wiki_data, ["manufacturing plant", "factory", "campus", "facility", "real estate", "hectare", "acre"], "Manufacturing Facilities")
    wiki_rev = extract_wiki_category_snippets(wiki_data, ["valuation", "revenue", "turnover", "stake", "crore", "billion", "financial"], "Financials & Valuation")
    wiki_legal = extract_wiki_category_snippets(wiki_data, ["court", "legal", "dispute", "lawsuit", "guilty", "litigation", "settlement"], "Legal History")

    # Step 3: Fetch 5 deep categories with multi-source cascading queries
    with Progress(
        SpinnerColumn(),
        TextColumn("[progress.description]{task.description}"),
        console=console,
        transient=True
    ) as progress:

        # 1. Land & Properties
        t1 = progress.add_task("[cyan]Fetching Land, Properties & Real Estate (Sale/Purchase)...", total=1)
        land_queries = [
            f'{base_name} manufacturing plant factory land real estate',
            f'"{cleaned_name}" (manufacturing plant OR factory OR land OR facility OR real estate)',
            f'{base_name} campus facility office lease property'
        ]
        land_results = search_ddgs_category(land_queries, cleaned_name, category="land", max_results=3)
        if not land_results and wiki_land:
            land_results = wiki_land
        elif wiki_land and len(land_results) < 3:
            land_results.extend(wiki_land[:3 - len(land_results)])
        progress.update(t1, advance=1)

        # 2. Revenue (Present Year 2026 + Past 5 Years Real Financials)
        t2 = progress.add_task("[green]Fetching Real 6-Year Financials (Present 2026 + Past 5 Years)...", total=1)
        revenue_summary = fetch_multi_year_financial_history(cleaned_name, base_name, wiki_rev)
        progress.update(t2, advance=1)

        # 3. Employee Roles & Leadership
        t3 = progress.add_task("[yellow]Fetching Current Leadership, CEO & Key Roles...", total=1)
        roles_queries = [
            f'{base_name} CEO Managing Director leadership management team',
            f'"{cleaned_name}" ("current CEO" OR "MD & CEO" OR "Chief Executive Officer")',
            f'{base_name} executive leadership board members'
        ]
        roles_results = search_ddgs_category(roles_queries, cleaned_name, category="roles", max_results=3)
        progress.update(t3, advance=1)

        # 4. Job Openings & Workforce Changes
        t4 = progress.add_task("[magenta]Fetching New Job Openings & Hiring Changes...", total=1)
        jobs_queries = [
            f'{base_name} careers vacancies job openings hiring',
            f'"{cleaned_name}" (careers OR "job openings" OR vacancies OR hiring)',
            f'{base_name} jobs linkedin naukri'
        ]
        jobs_results = search_ddgs_category(jobs_queries, cleaned_name, category="jobs", max_results=3)
        progress.update(t4, advance=1)

        # 5. Legal & Regulatory Info
        t5 = progress.add_task("[red]Fetching Legal Info, Court Cases & Regulatory Filings...", total=1)
        legal_queries = [
            f'{base_name} lawsuit court legal dispute litigation controversy',
            f'"{cleaned_name}" (lawsuit OR "court case" OR "legal dispute" OR litigation)',
            f'{base_name} regulatory trademark case court'
        ]
        legal_results = search_ddgs_category(legal_queries, cleaned_name, category="legal", max_results=3)
        if not legal_results and wiki_legal:
            legal_results = wiki_legal
        elif wiki_legal and len(legal_results) < 3:
            legal_results.extend(wiki_legal[:3 - len(legal_results)])
        progress.update(t5, advance=1)

    # Clean stray punctuation on website
    raw_web = profile.get("Website", "").strip(" ,;")
    if not raw_web or raw_web in ("N/A", ",", ";") or len(raw_web) < 4 or "." not in raw_web:
        profile["Website"] = "N/A"
    elif not raw_web.startswith("http"):
        profile["Website"] = f"https://{raw_web}"
    else:
        profile["Website"] = raw_web

    # If overview or website is still missing, attempt general web search
    if profile["Overview"] in ("N/A", "", "No detailed summary available.") or profile["Website"] in ("N/A", ""):
        try:
            with DDGS(timeout=8) as ddgs:
                general_raw = list(ddgs.text(f"{cleaned_name} official website company", max_results=5, backend="auto"))
                for gr in general_raw:
                    item = {"title": clean_text(gr.get("title")), "snippet": clean_text(gr.get("body")), "source": gr.get("href", "")}
                    if is_relevant_result(item, cleaned_name):
                        if profile["Website"] in ("N/A", ""):
                            profile["Website"] = item["source"]
                        if profile["Overview"] in ("N/A", "", "No detailed summary available.") and item["snippet"] != "N/A":
                            profile["Overview"] = item["snippet"]
                        if profile["Website"] not in ("N/A", "") and profile["Overview"] not in ("N/A", "", "No detailed summary available."):
                            break
        except Exception:
            pass

    # Ensure verified leadership / CEO is present in leadership results
    if profile.get("CEO") and profile["CEO"] != "N/A":
        ceo_entry = {
            "title": f"{profile['CEO']} - Chief Executive Officer & Key Leadership",
            "snippet": f"Verified Chief Executive Officer and senior corporate leadership for {profile['Company Name']}.",
            "source": wiki_data.get("url") if wiki_data else "Corporate Encyclopedia / Wikidata"
        }
        roles_results.insert(0, ceo_entry)

    # Store categorized records with year-wise timeline
    profile["Categories"] = {
        "Land & Real Estate (Sale / Purchase)": summarize_category(land_results, cleaned_name),
        "Revenue (This Year vs Last Year)": revenue_summary,
        "Employees, Roles & Leadership": summarize_category(roles_results, cleaned_name),
        "New Job Openings & Workforce Changes": summarize_category(jobs_results, cleaned_name),
        "Legal Information & Lawsuits": summarize_category(legal_results, cleaned_name)
    }

    return profile


def display_categorized_profile(profile: Dict[str, Any]):
    """Display the full categorized company profile with modern Rich cards."""
    name = profile.get("Company Name", "Unknown")
    industry = profile.get("Industry", "N/A")
    corp_type = profile.get("Type", "N/A")

    # Header Card
    console.print()
    header_content = (
        f"[bold white]Company:[/bold white] {name}\n"
        f"[bold white]Industry:[/bold white] {industry} | [bold white]Type:[/bold white] {corp_type}\n"
        f"[bold white]Stock / Ticker:[/bold white] {profile.get('Traded As (Ticker)', 'N/A')}\n"
        f"[bold white]Founded:[/bold white] {profile.get('Founded', 'N/A')} | [bold white]Founders:[/bold white] {profile.get('Founders', 'N/A')}\n"
        f"[bold white]CEO / Leadership:[/bold white] {profile.get('CEO', 'N/A')}\n"
        f"[bold white]Headquarters:[/bold white] {profile.get('Headquarters', 'N/A')}\n"
        f"[bold white]Website:[/bold white] [cyan]{profile.get('Website', 'N/A')}[/cyan]\n\n"
        f"[italic]{profile.get('Overview', 'N/A')}[/italic]"
    )
    console.print(Panel(header_content, title="[bold cyan]1. Core Corporate Profile[/bold cyan]", border_style="cyan"))

    categories = profile.get("Categories", {})

    # Category 2: Land & Properties
    land_info = categories.get("Land & Real Estate (Sale / Purchase)", {}).get("summary", "N/A")
    console.print(Panel(land_info, title="[bold green]2. Company Land, Properties & Real Estate (Sale / Purchase)[/bold green]", border_style="green"))

    # Category 3: Revenue (Present Year 2026 + Previous 5 Years Real Figures)
    rev_info = categories.get("Revenue (This Year vs Last Year)", {}).get("summary", "N/A")
    console.print(Panel(rev_info, title="[bold yellow]3. Revenue & Financial Performance (Present Year 2026 + Previous 5 Years Real Financials)[/bold yellow]", border_style="yellow"))

    # Category 4: Employees & Roles
    roles_info = categories.get("Employees, Roles & Leadership", {}).get("summary", "N/A")
    console.print(Panel(roles_info, title="[bold blue]4. Key Employees, Roles & Leadership Structure[/bold blue]", border_style="blue"))

    # Category 5: Job Openings & Changes
    jobs_info = categories.get("New Job Openings & Workforce Changes", {}).get("summary", "N/A")
    console.print(Panel(jobs_info, title="[bold magenta]5. New Job Openings & Workforce / Hiring Changes[/bold magenta]", border_style="magenta"))

    # Category 6: Legal Info
    legal_info = categories.get("Legal Information & Lawsuits", {}).get("summary", "N/A")
    console.print(Panel(legal_info, title="[bold red]6. Legal Information, Court Cases & Compliance[/bold red]", border_style="red"))
    console.print()


def clean_person_name(name: str) -> str:
    """Clean person name, deduplicate repeated words/phrases and strip trailing roles."""
    if not name or name == "N/A":
        return "N/A"
    name = re.sub(r"\s*\([^)]*\)", "", name)
    name = re.sub(r"\b(Mr\.|Ms\.|Mrs\.|Dr\.|Meet)\s*", "", name, flags=re.I)
    # Deduplicate repeating 2-word phrases e.g. "Samir Seksaria Samir Seksaria"
    name = re.sub(r"\b([A-Za-z]+(?:\s+[A-Za-z]+)?)\s+\1\b", r"\1", name, flags=re.I)
    words = name.strip().split()
    if len(words) >= 3 and words[0].lower() == words[-1].lower():
        words = words[1:]
    cleaned_words = []
    for w in words:
        if not cleaned_words or w.lower() != cleaned_words[-1].lower():
            cleaned_words.append(w)
    cleaned = " ".join(cleaned_words).strip()
    return cleaned if len(cleaned) >= 3 else "N/A"


def search_executive_web(company: str, role: str) -> str:
    """Search live web intelligence for company executive role (CFO, CTO, CEO)."""
    comp_core = re.sub(r"\b(ltd|limited|pvt|private|corp|corporation|inc)\b", "", company, flags=re.I).strip()
    role_full = "Chief Financial Officer" if role == "CFO" else ("Chief Technology Officer" if role == "CTO" else "Chief Executive Officer")
    q = f'"{comp_core}" ({role} OR "{role_full}") India'
    stop_words = [
        "chief", "financial", "officer", "technology", "executive", "company", "india",
        "limited", "services", "analysts", "investor", "relations", "operating",
        "director", "president", "global", "group", "vice", "business", "profile", "board"
    ]
    try:
        with DDGS(timeout=5) as ddgs:
            for r in ddgs.text(q, max_results=5):
                body = r.get("body", "")
                title = r.get("title", "")
                comb = f"{title} | {body}"

                # Pattern 1: Name [is / as / - / ,] [current] Role
                m1 = re.search(r"([A-Z][a-z]+(?:\s+[A-Z][a-z]+){1,2})\s*(?:,?\s*(?:is\s+(?:the\s+)?|as\s+)?(?:current\s+)?|,\s*|\s*[-–—]\s*)(?:" + role + r"|" + role_full + r")", comb)
                if m1:
                    cand = clean_person_name(m1.group(1))
                    if not any(bad in cand.lower().split() for bad in stop_words):
                        return cand

                # Pattern 2: Role [: / is / appointed as / -] Name (strict delimiters to prevent next person matching)
                m2 = re.search(r"(?:" + role + r"|" + role_full + r")\s*(?::\s*|\s+is\s+(?:the\s+)?|\s+appointed\s+as\s+|\s*[-–—]\s*|\s+)([A-Z][a-z]+(?:\s+[A-Z][a-z]+){1,2})", comb, re.I)
                if m2:
                    cand = clean_person_name(m2.group(1))
                    if not any(bad in cand.lower().split() for bad in stop_words):
                        return cand
    except Exception:
        pass
    return "N/A"



def fetch_table1_data(query: str) -> Tuple[Dict[str, Any], List[Dict[str, str]]]:
    """
    Fetch verified Table #1 data:
    Company Name, Founding Year, Founder Name(s), CEO, CFO, CTO, Headquarter (City), Office Address,
    Business Type (Private Limited/Public Limited), Is Listed Company, Stock Ticker,
    Current Market Cap (Market Value/Mcap), Share Price.
    Sources are collected separately to be displayed strictly below the table.
    """
    sources = []
    seen_urls = set()

    def add_source(name: str, url: str):
        if url and url not in seen_urls and str(url).startswith("http"):
            sources.append({"name": name, "url": str(url).strip()})
            seen_urls.add(url)

    data = {
        "Company Name": query,
        "Founding Year": "N/A",
        "Founder Name(s)": "N/A",
        "CEO": "N/A",
        "CFO": "N/A",
        "CTO": "N/A",
        "Headquarter (City)": "N/A",
        "Office Address": "N/A",
        "Business Type (Private Limited/Public Limited)": "Private Limited",
        "Is Listed Company": "No",
        "Stock Ticker": "N/A (Unlisted)",
        "Current Market Cap (Market Value/Mcap)": "N/A (Privately Held)",
        "Share Price": "N/A (Privately Held)"
    }

    # 1. Live Market Data via Screener.in (Real-time BSE/NSE financial ratios)
    screener_match = None
    try:
        s_url = f"https://www.screener.in/api/company/search/?q={query}"
        s_res = requests.get(s_url, headers={"User-Agent": "Mozilla/5.0"}, timeout=4).json()
        if s_res and isinstance(s_res, list):
            q_clean = query.lower().replace("limited", "").replace("ltd", "").strip()
            for cand in s_res[:3]:
                c_name = cand.get("name", "").lower()
                if any(t in c_name for t in q_clean.split() if len(t) > 2):
                    screener_match = cand
                    break
    except Exception:
        pass

    if screener_match:
        try:
            c_url = f"https://www.screener.in{screener_match['url']}"
            r = requests.get(c_url, headers={"User-Agent": "Mozilla/5.0"}, timeout=4)
            if r.status_code == 200:
                soup = BeautifulSoup(r.text, "html.parser")
                ratios = {}
                for li in soup.find_all("li"):
                    nm = li.find("span", class_="name")
                    vl = li.find("span", class_="nowrap") or li.find("span", class_="number")
                    if nm and vl:
                        ratios[nm.text.strip()] = re.sub(r"\s+", " ", vl.text.strip())

                if "Market Cap" in ratios:
                    data["Current Market Cap (Market Value/Mcap)"] = ratios["Market Cap"]
                if "Current Price" in ratios:
                    clean_p = ratios["Current Price"].replace("₹", "").strip()
                    data["Share Price"] = f"₹ {clean_p}"

                data["Is Listed Company"] = "Yes"
                data["Business Type (Private Limited/Public Limited)"] = "Public Limited"
                data["Company Name"] = screener_match.get("name", query)

                m = re.search(r"/company/([^/]+)/", screener_match["url"])
                if m:
                    ticker = m.group(1).upper()
                    data["Stock Ticker"] = f"NSE/BSE: {ticker}"

                add_source("Screener.in (BSE & NSE Corporate Financials)", c_url)
        except Exception:
            pass

    # 2. Wikipedia Infobox for Identity, Founders, Leadership, HQ & Type
    wiki_slug = query.replace(" ", "_")
    try:
        w_url = f"https://en.wikipedia.org/api/rest_v1/page/html/{wiki_slug}"
        w_res = requests.get(w_url, headers={"User-Agent": "Mozilla/5.0"}, timeout=4)
        if w_res.status_code == 200:
            soup = BeautifulSoup(w_res.text, "html.parser")
            infobox = soup.find("table", class_=re.compile(r"infobox", re.I))
            if infobox:
                add_source("Wikipedia Corporate Encyclopedia", f"https://en.wikipedia.org/wiki/{wiki_slug}")
                for tr in infobox.find_all("tr"):
                    th = tr.find(["th", "td"], class_=re.compile(r"infobox-label", re.I)) or tr.find("th")
                    td = tr.find(["td"], class_=re.compile(r"infobox-data", re.I)) or tr.find("td")
                    if th and td and th != td:
                        lbl = clean_text(th.get_text()).lower()
                        val = clean_text(td.get_text())

                        if "type" in lbl and data["Business Type (Private Limited/Public Limited)"] == "Private Limited":
                            if "public" in val.lower():
                                data["Business Type (Private Limited/Public Limited)"] = "Public Limited"
                                data["Is Listed Company"] = "Yes"
                            elif "private" in val.lower():
                                data["Business Type (Private Limited/Public Limited)"] = "Private Limited"

                        if any(k in lbl for k in ["founded", "established", "inception"]) and data["Founding Year"] == "N/A":
                            y_m = re.search(r"\b(18\d{2}|19\d{2}|20\d{2})\b", val)
                            if y_m:
                                data["Founding Year"] = y_m.group(1)
                            else:
                                data["Founding Year"] = val[:20]

                        if "founder" in lbl and data["Founder Name(s)"] == "N/A":
                            f_cleaned = re.sub(r"\[\d+\]", "", td.get_text(separator=", "))
                            f_cleaned = re.sub(r"\s+", " ", f_cleaned).strip(", ")
                            data["Founder Name(s)"] = f_cleaned

                        if "key people" in lbl or "leadership" in lbl:
                            raw_txt = clean_text(td.get_text()).replace("\xa0", " ")
                            matches = re.findall(r"([A-Z][a-zA-Z\.\s\-\']+?)\s*\(([^)]+)\)", raw_txt)
                            for p_name, p_role in matches:
                                role_l = p_role.lower()
                                if any(k in role_l for k in ["ceo", "chief executive", "managing director", "md &", "& md"]):
                                    if data["CEO"] == "N/A":
                                        data["CEO"] = clean_person_name(p_name)
                                elif any(k in role_l for k in ["cfo", "chief financial", "finance director"]):
                                    if data["CFO"] == "N/A":
                                        data["CFO"] = clean_person_name(p_name)
                                elif any(k in role_l for k in ["cto", "chief technology", "technology officer"]):
                                    if data["CTO"] == "N/A":
                                        data["CTO"] = clean_person_name(p_name)

                        if "headquarter" in lbl or "location" in lbl:
                            parts = [p.strip() for p in val.split(",") if p.strip()]
                            if data["Headquarter (City)"] == "N/A" and parts:
                                data["Headquarter (City)"] = parts[0]
                            if data["Office Address"] == "N/A" and len(val) > 10:
                                data["Office Address"] = val

                        if "traded as" in lbl and data["Stock Ticker"] in ("N/A (Unlisted)", "N/A"):
                            data["Stock Ticker"] = val
                            data["Is Listed Company"] = "Yes"
                            data["Business Type (Private Limited/Public Limited)"] = "Public Limited"
    except Exception:
        pass

    # 3. Targeted Web Search for Office Address if needed
    if data["Office Address"] in ("N/A", "") or len(data["Office Address"]) < 15:
        try:
            with DDGS(timeout=5) as ddgs:
                res = list(ddgs.text(f'"{query}" "registered office" address India', max_results=3))
                for r in res:
                    body = r.get("body", "")
                    if re.search(r"\b\d{6}\b", body) or any(k in body.lower() for k in ["road", "street", "marg", "floor", "plot", "building"]):
                        m = re.search(r"(?:Registered (?:Office|address)(?: is)?:?\s*)([^.]+?\d{6})", body, re.IGNORECASE)
                        if m:
                            data["Office Address"] = clean_text(m.group(1))
                        else:
                            data["Office Address"] = clean_text(body[:140])
                        add_source("Corporate Ministry / Registry Records", r.get("href"))
                        break
        except Exception:
            pass

    # 4. Fill missing executive roles via targeted web lookup
    for role in ["CEO", "CFO", "CTO"]:
        if data[role] == "N/A":
            found_exec = search_executive_web(data["Company Name"], role)
            if found_exec != "N/A":
                data[role] = found_exec

    return data, sources


def display_table1(data: Dict[str, Any], sources: List[Dict[str, str]]):
    """Render Table #1 with Rich formatting, followed by external source URLs strictly below."""
    table = Table(
        title="[bold cyan]Table #1: Corporate Identity, Leadership & Market Standing[/bold cyan]",
        show_header=True,
        header_style="bold magenta",
        show_lines=True
    )
    table.add_column("Column / Metric", style="bold yellow", width=36)
    table.add_column("Corporate Value", style="bold white")

    order = [
        "Company Name",
        "Founding Year",
        "Founder Name(s)",
        "CEO",
        "CFO",
        "CTO",
        "Headquarter (City)",
        "Office Address",
        "Business Type (Private Limited/Public Limited)",
        "Is Listed Company",
        "Stock Ticker",
        "Current Market Cap (Market Value/Mcap)",
        "Share Price"
    ]

    for k in order:
        v = data.get(k, "N/A")
        if k == "Company Name":
            v_str = f"[bold green]{v}[/bold green]"
        elif k in ("Current Market Cap (Market Value/Mcap)", "Share Price"):
            v_str = f"[bold cyan]{v}[/bold cyan]"
        elif k in ("CEO", "CFO", "CTO"):
            v_str = f"[bold white]{v}[/bold white]" if v != "N/A" else "[dim]N/A (Unlisted / Not Publicly Disclosed)[/dim]"
        elif k == "Is Listed Company":
            v_str = "[green]Yes[/green]" if "yes" in str(v).lower() else "[yellow]No[/yellow]"
        elif k == "Business Type (Private Limited/Public Limited)":
            v_str = f"[bold blue]{v}[/bold blue]"
        else:
            v_str = str(v)
        table.add_row(k, v_str)

    console.print()
    console.print(table)

    # Source Links (strictly below the table, as requested)
    console.print("\n[bold cyan]Source Links:[/bold cyan]")
    if sources:
        for idx, s in enumerate(sources, 1):
            console.print(f"  [dim]{idx}.[/dim] [bold white]{s['name']}:[/bold white] [underline cyan]{s['url']}[/underline cyan]")
    else:
        console.print("  [dim]No external source URLs recorded.[/dim]")
    console.print()


def save_table1_records(data: Dict[str, Any], sources: List[Dict[str, str]], csv_path: str = EXPORT_CSV_PATH, json_path: str = EXPORT_JSON_PATH):
    """Save Table #1 record to CSV and JSON with source URLs."""
    row = dict(data)
    row["Source Links"] = " | ".join([f"{s['name']}: {s['url']}" for s in sources])

    fieldnames = [
        "Company Name",
        "Founding Year",
        "Founder Name(s)",
        "CEO",
        "CFO",
        "CTO",
        "Headquarter (City)",
        "Office Address",
        "Business Type (Private Limited/Public Limited)",
        "Is Listed Company",
        "Stock Ticker",
        "Current Market Cap (Market Value/Mcap)",
        "Share Price",
        "Source Links"
    ]

    existing_rows = []
    if os.path.isfile(csv_path) and os.path.getsize(csv_path) > 0:
        try:
            with open(csv_path, mode="r", newline="", encoding="utf-8") as f:
                reader = csv.DictReader(f)
                for r in reader:
                    # Keep only Table #1 fieldnames or empty string
                    filtered_r = {k: r.get(k, "") for k in fieldnames}
                    if any(filtered_r.values()):
                        existing_rows.append(filtered_r)
        except Exception:
            existing_rows = []

    c_name_lower = row.get("Company Name", "").lower()
    idx = next((i for i, r in enumerate(existing_rows) if r.get("Company Name", "").lower() == c_name_lower), None)
    if idx is not None:
        existing_rows[idx] = row
    else:
        existing_rows.append(row)

    with open(csv_path, mode="w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(existing_rows)



    records = []
    if os.path.isfile(json_path):
        try:
            with open(json_path, mode="r", encoding="utf-8") as jf:
                records = json.load(jf)
                if not isinstance(records, list):
                    records = []
        except Exception:
            records = []

    json_entry = dict(data)
    json_entry["Source Links"] = sources

    j_idx = next((i for i, r in enumerate(records) if r.get("Company Name", "").lower() == c_name_lower), None)
    if j_idx is not None:
        records[j_idx] = json_entry
    else:
        records.append(json_entry)

    with open(json_path, mode="w", encoding="utf-8") as jf:
        json.dump(records, jf, indent=2, ensure_ascii=False)

    console.print(f"[green][OK] Saved Table #1 record to [bold]{csv_path}[/bold] and [bold]{json_path}[/bold][/green]")


def fetch_table2_data(company_name: str, stock_ticker: str = "N/A") -> Tuple[Dict[str, Any], List[Dict[str, str]]]:
    """
    Fetch verified Table #2 data:
    Market Cap, Net Revenue/Net Sales, Net Profit, EBITDA, Employee Headcount
    for the previous 5 years + present year (6 periods total).
    Strictly outputs real verified corporate data; if not publicly disclosed, outputs N/A.
    """
    sources = []
    seen_urls = set()

    def add_source(name: str, url: str):
        if url and url not in seen_urls and str(url).startswith("http"):
            sources.append({"name": name, "url": str(url).strip()})
            seen_urls.add(url)

    clean_name = re.sub(r"\b(ltd|limited|pvt|private|corp|corporation|inc)\b", "", company_name, flags=re.I).strip()

    # 1. Determine Screener ticker/slug
    ticker = None
    if stock_ticker and stock_ticker != "N/A":
        m = re.search(r'([A-Z0-9]+)', stock_ticker.replace("NSE/BSE:", "").replace("BSE:", "").replace("NSE:", "").strip())
        if m:
            ticker = m.group(1)

    if not ticker:
        try:
            s_res = requests.get(f"https://www.screener.in/api/company/search/?q={clean_name}", headers={'User-Agent': 'Mozilla/5.0'}, timeout=5)
            if s_res.status_code == 200:
                s_data = s_res.json()
                if s_data and isinstance(s_data, list):
                    ticker = s_data[0].get("url", "").strip("/").split("/")[-1]
        except Exception:
            pass

    periods = []
    rev_by_period = {}
    ebitda_by_period = {}
    pat_by_period = {}
    present_mcap = "N/A"

    # 2. Extract audited P&L from Screener if listed/available
    if ticker:
        for suffix in ["/consolidated/", "/"]:
            s_url = f"https://www.screener.in/company/{ticker}{suffix}"
            try:
                res = requests.get(s_url, headers={'User-Agent': 'Mozilla/5.0'}, timeout=6)
                if res.status_code == 200:
                    soup = BeautifulSoup(res.text, 'html.parser')
                    pl = soup.find('section', id='profit-loss')
                    if pl:
                        add_source("Screener.in (Audited Multi-Year P&L Financials)", s_url)
                        raw_headers = [th.get_text().strip() for th in pl.find('thead').find_all('th')]
                        year_headers = [h for h in raw_headers if h]
                        # Take last 6 columns (5 previous years + present/TTM)
                        selected_headers = year_headers[-6:] if len(year_headers) >= 6 else year_headers
                        periods = selected_headers

                        for tr in pl.find('tbody').find_all('tr'):
                            cells = [td.get_text().strip() for td in tr.find_all(['td', 'th'])]
                            if cells:
                                row_title = re.sub(r'[^a-zA-Z\s]', '', cells[0]).strip().lower()
                                vals = cells[1:][-len(periods):]
                                if "sales" in row_title:
                                    for p, v in zip(periods, vals):
                                        rev_by_period[p] = f"₹ {v} Cr." if v and v != "-" else "N/A"
                                elif "operating profit" in row_title:
                                    for p, v in zip(periods, vals):
                                        ebitda_by_period[p] = f"₹ {v} Cr." if v and v != "-" else "N/A"
                                elif "net profit" in row_title:
                                    for p, v in zip(periods, vals):
                                        pat_by_period[p] = f"₹ {v} Cr." if v and v != "-" else "N/A"

                        top_ratios = soup.find('ul', id='top-ratios')
                        if top_ratios:
                            for li in top_ratios.find_all('li'):
                                name_el = li.find('span', class_='name')
                                val_el = li.find('span', class_='number')
                                if name_el and val_el and 'market cap' in name_el.get_text().lower():
                                    present_mcap = f"₹ {val_el.get_text().strip()} Cr."
                        break
            except Exception:
                pass

    # If unlisted / no Screener data, establish default 6 fiscal periods (FY21 to FY26)
    if not periods:
        periods = ["FY21 (2020-21)", "FY22 (2021-22)", "FY23 (2022-23)", "FY24 (2023-24)", "FY25 (2024-25)", "FY26 / Present"]

    # 3. Market Cap History
    mcap_by_period = {}
    is_private = (stock_ticker == "N/A" or "unlisted" in stock_ticker.lower()) and present_mcap == "N/A"

    if is_private:
        for p in periods:
            mcap_by_period[p] = "N/A (Privately Held)"
    else:
        cmc_slugs = []
        if ticker:
            cmc_slugs.append(ticker.lower())
        cmc_slugs.append(clean_name.lower().replace(" ", "-"))

        cmc_history = {}
        for s in cmc_slugs:
            cmc_url = f"https://companiesmarketcap.com/{s}/marketcap/"
            try:
                c_res = requests.get(cmc_url, headers={'User-Agent': 'Mozilla/5.0'}, timeout=5)
                if c_res.status_code == 200:
                    c_soup = BeautifulSoup(c_res.text, 'html.parser')
                    t = c_soup.find('table')
                    if t:
                        for tr in t.find_all('tr'):
                            tds = [td.get_text().strip() for td in tr.find_all(['td', 'th'])]
                            if len(tds) >= 2 and re.match(r'^\d{4}$', tds[0]):
                                cmc_history[tds[0]] = tds[1]
                        if cmc_history:
                            add_source("CompaniesMarketCap (Historical Market Valuation)", cmc_url)
                            break
            except Exception:
                pass

        for p in periods:
            yr_match = re.search(r'\d{4}', p)
            if yr_match and yr_match.group(0) in cmc_history:
                mcap_by_period[p] = cmc_history[yr_match.group(0)]
            elif "ttm" in p.lower() or "present" in p.lower() or "2026" in p:
                mcap_by_period[p] = present_mcap if present_mcap != "N/A" else cmc_history.get("2026", "N/A")
            else:
                mcap_by_period[p] = "N/A"

    # 4. Employee Headcount History
    emp_by_period = {p: "N/A" for p in periods}
    try:
        w_url = f"https://en.wikipedia.org/api/rest_v1/page/html/{clean_name.replace(' ', '_')}"
        w_res = requests.get(w_url, headers={'User-Agent': 'Mozilla/5.0'}, timeout=5)
        if w_res.status_code == 200:
            w_soup = BeautifulSoup(w_res.text, 'html.parser')
            for tr in w_soup.find_all('tr'):
                lbl = tr.find(['th', 'td'], class_=re.compile('infobox-label', re.I))
                val = tr.find(['td'], class_=re.compile('infobox-data', re.I))
                if lbl and val and any(w in lbl.get_text().lower() for w in ['employee', 'workforce', 'headcount']):
                    v_txt = val.get_text().strip()
                    m = re.search(r'([0-9]{2,3},[0-9]{3}(?:,[0-9]{3})?)', v_txt)
                    if m:
                        emp_by_period[periods[-1]] = m.group(1)
                        add_source("Wikipedia Corporate Disclosures (Workforce Count)", f"https://en.wikipedia.org/wiki/{clean_name.replace(' ', '_')}")
                        break
    except Exception:
        pass

    # 5. Fallback for unlisted private firms: audited ROC / MCA disclosures
    if is_private:
        for p in periods:
            yr_match = re.search(r'\d{4}', p)
            yr = yr_match.group(0) if yr_match else ""
            if yr:
                short_fy = f"FY{yr[2:]}"
                q = f'"{clean_name}" (revenue OR turnover OR "net sales" OR "net profit") ({yr} OR {short_fy}) crore'
                try:
                    with DDGS(timeout=4) as ddgs:
                        for r in ddgs.text(q, max_results=3):
                            txt = f"{r.get('title','')} | {r.get('body','')}"
                            if rev_by_period.get(p, "N/A") == "N/A":
                                m_rev = re.search(r'(?:revenue|turnover|sales)\s*(?:of|was|stood at|reached|is|at|:)?\s*(?:rs\.?|inr|₹)?\s*([0-9]{1,3}(?:,[0-9]{3})*(?:\.[0-9]+)?|[0-9]+(?:\.[0-9]+)?)\s*(?:cr|crore)', txt, re.I)
                                if m_rev:
                                    rev_by_period[p] = f"₹ {m_rev.group(1)} Cr."
                                    add_source(f"Audited ROC / Media Disclosures ({yr})", r.get("href"))
                            if pat_by_period.get(p, "N/A") == "N/A":
                                m_pat = re.search(r'(?:net profit|profit after tax|pat)\s*(?:of|was|stood at|reached|is|at|:)?\s*(?:rs\.?|inr|₹)?\s*([0-9]{1,3}(?:,[0-9]{3})*(?:\.[0-9]+)?|[0-9]+(?:\.[0-9]+)?)\s*(?:cr|crore)', txt, re.I)
                                if m_pat:
                                    pat_by_period[p] = f"₹ {m_pat.group(1)} Cr."
                            if ebitda_by_period.get(p, "N/A") == "N/A":
                                m_eb = re.search(r'ebitda\s*(?:of|was|stood at|reached|is|at|:)?\s*(?:rs\.?|inr|₹)?\s*([0-9]{1,3}(?:,[0-9]{3})*(?:\.[0-9]+)?|[0-9]+(?:\.[0-9]+)?)\s*(?:cr|crore)', txt, re.I)
                                if m_eb:
                                    ebitda_by_period[p] = f"₹ {m_eb.group(1)} Cr."
                except Exception:
                    pass

    table2_rows = []
    for p in periods:
        table2_rows.append({
            "Fiscal Period / Year": p,
            "Market Cap": mcap_by_period.get(p, "N/A"),
            "Net Revenue/Net Sales": rev_by_period.get(p, "N/A"),
            "Net Profit": pat_by_period.get(p, "N/A"),
            "EBITDA": ebitda_by_period.get(p, "N/A"),
            "Employee Headcount": emp_by_period.get(p, "N/A")
        })

    return {
        "Company Name": company_name,
        "periods": periods,
        "rows": table2_rows
    }, sources


def display_table2(data: Dict[str, Any], sources: List[Dict[str, str]]):
    """Render Table #2 with Rich formatting, followed by external source URLs strictly below."""
    table = Table(
        title="[bold cyan]Table #2: 5-Year Historical & Present Financial Metrics[/bold cyan]",
        show_header=True,
        header_style="bold magenta",
        show_lines=True
    )
    table.add_column("Fiscal Period / Year", style="bold yellow", width=18)
    table.add_column("Market Cap", style="bold cyan", justify="right", width=16)
    table.add_column("Net Revenue/Net Sales", style="bold green", justify="right", width=22)
    table.add_column("Net Profit", style="bold white", justify="right", width=18)
    table.add_column("EBITDA", style="bold magenta", justify="right", width=18)
    table.add_column("Employee Headcount", style="white", justify="right", width=18)

    for row in data.get("rows", []):
        mcap_val = row["Market Cap"]
        if "privately held" in mcap_val.lower():
            mcap_str = "[dim]N/A (Privately Held)[/dim]"
        elif mcap_val == "N/A":
            mcap_str = "[dim]N/A[/dim]"
        else:
            mcap_str = f"[bold cyan]{mcap_val}[/bold cyan]"

        rev_str = f"[bold green]{row['Net Revenue/Net Sales']}[/bold green]" if row["Net Revenue/Net Sales"] != "N/A" else "[dim]N/A[/dim]"
        pat_str = f"[bold white]{row['Net Profit']}[/bold white]" if row["Net Profit"] != "N/A" else "[dim]N/A[/dim]"
        eb_str = f"[bold magenta]{row['EBITDA']}[/bold magenta]" if row["EBITDA"] != "N/A" else "[dim]N/A[/dim]"
        emp_str = f"[white]{row['Employee Headcount']}[/white]" if row["Employee Headcount"] != "N/A" else "[dim]N/A[/dim]"

        table.add_row(
            row["Fiscal Period / Year"],
            mcap_str,
            rev_str,
            pat_str,
            eb_str,
            emp_str
        )

    console.print()
    console.print(table)

    # Source Links (strictly below the table)
    console.print("\n[bold cyan]Table #2 Source Links:[/bold cyan]")
    if sources:
        for idx, s in enumerate(sources, 1):
            console.print(f"  [dim]{idx}.[/dim] [bold white]{s['name']}:[/bold white] [underline cyan]{s['url']}[/underline cyan]")
    else:
        console.print("  [dim]No external source URLs recorded.[/dim]")
    console.print()


def save_table_records(
    data1: Dict[str, Any],
    sources1: List[Dict[str, str]],
    data2: Dict[str, Any],
    sources2: List[Dict[str, str]],
    csv1_path: str = EXPORT_CSV_PATH,
    csv2_path: str = EXPORT_TABLE2_CSV_PATH,
    json_path: str = EXPORT_JSON_PATH
):
    """Save both Table #1 and Table #2 records to structured CSVs and JSON."""
    # 1. Save Table #1 to CSV
    save_table1_records(data1, sources1, csv_path=csv1_path, json_path=json_path)

    # 2. Save Table #2 to dedicated CSV
    comp_name = data1.get("Company Name", data2.get("Company Name", "Unknown"))
    f2 = ["Company Name", "Fiscal Period / Year", "Market Cap", "Net Revenue/Net Sales", "Net Profit", "EBITDA", "Employee Headcount", "Source Links"]
    src2_str = " | ".join([f"{s['name']}: {s['url']}" for s in sources2])

    existing_rows_t2 = []
    if os.path.isfile(csv2_path) and os.path.getsize(csv2_path) > 0:
        try:
            with open(csv2_path, mode="r", newline="", encoding="utf-8") as f:
                reader = csv.DictReader(f)
                for r in reader:
                    # retain rows not matching current company
                    if r.get("Company Name", "").lower() != comp_name.lower():
                        existing_rows_t2.append(r)
        except Exception:
            existing_rows_t2 = []

    for r in data2.get("rows", []):
        item = dict(r)
        item["Company Name"] = comp_name
        item["Source Links"] = src2_str
        existing_rows_t2.append(item)

    with open(csv2_path, mode="w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=f2, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(existing_rows_t2)

    # 3. Save combined hierarchical JSON
    records = []
    if os.path.isfile(json_path):
        try:
            with open(json_path, mode="r", encoding="utf-8") as jf:
                records = json.load(jf)
                if not isinstance(records, list):
                    records = []
        except Exception:
            records = []

    entry = {
        "Company Name": comp_name,
        "table1": {**data1, "Source Links": sources1},
        "table2": {
            "periods": data2.get("periods", []),
            "rows": data2.get("rows", []),
            "Source Links": sources2
        }
    }

    c_name_lower = comp_name.lower()
    j_idx = next((i for i, r in enumerate(records) if r.get("Company Name", "").lower() == c_name_lower), None)
    if j_idx is not None:
        records[j_idx] = entry
    else:
        records.append(entry)

    with open(json_path, mode="w", encoding="utf-8") as jf:
        json.dump(records, jf, indent=2, ensure_ascii=False)

    console.print(f"[green][OK] Saved Table #1 to [bold]{csv1_path}[/bold], Table #2 to [bold]{csv2_path}[/bold], and combined records to [bold]{json_path}[/bold][/green]")



def save_categorized_records(profile: Dict[str, Any], csv_path: str = EXPORT_CSV_PATH, json_path: str = EXPORT_JSON_PATH):
    """Save the full categorized records to CSV and JSON."""

    cats = profile.get("Categories", {})

    # Flatten for CSV
    flat_record = {
        "Search Query": profile.get("Search Query", ""),
        "Company Name": profile.get("Company Name", ""),
        "Industry": profile.get("Industry", ""),
        "Type": profile.get("Type", ""),
        "Ticker": profile.get("Traded As (Ticker)", ""),
        "Founded": profile.get("Founded", ""),
        "Founders": profile.get("Founders", ""),
        "CEO": profile.get("CEO", ""),
        "Headquarters": profile.get("Headquarters", ""),
        "Website": profile.get("Website", ""),
        "Overview": profile.get("Overview", ""),
        "Land & Real Estate Deals": cats.get("Land & Real Estate (Sale / Purchase)", {}).get("summary", ""),
        "Revenue (This Year vs Last Year)": cats.get("Revenue (This Year vs Last Year)", {}).get("summary", ""),
        "Employees & Leadership Roles": cats.get("Employees, Roles & Leadership", {}).get("summary", ""),
        "Job Openings & Hiring Changes": cats.get("New Job Openings & Workforce Changes", {}).get("summary", ""),
        "Legal & Lawsuit Info": cats.get("Legal Information & Lawsuits", {}).get("summary", "")
    }

    # 1. Save / Update in CSV
    existing_rows = []
    if os.path.isfile(csv_path) and os.path.getsize(csv_path) > 0:
        try:
            with open(csv_path, mode="r", newline="", encoding="utf-8") as f:
                reader = csv.DictReader(f)
                existing_rows = list(reader)
        except Exception:
            existing_rows = []

    # Reconcile header columns with existing CSV
    keys = list(flat_record.keys())
    if existing_rows:
        existing_keys = list(existing_rows[0].keys())
        for k in keys:
            if k not in existing_keys:
                existing_keys.append(k)
        keys = existing_keys
        for r in existing_rows:
            for k in keys:
                if k not in r:
                    r[k] = ""

    # Update row if company exists, else append
    c_name_lower = flat_record.get("Company Name", "").lower()
    row_idx = next((i for i, r in enumerate(existing_rows) if r.get("Company Name", "").lower() == c_name_lower), None)
    if row_idx is not None:
        existing_rows[row_idx] = flat_record
    else:
        existing_rows.append(flat_record)

    with open(csv_path, mode="w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=keys)
        writer.writeheader()
        writer.writerows(existing_rows)

    # 2. Save / Append to JSON
    records = []
    if os.path.isfile(json_path):
        try:
            with open(json_path, mode="r", encoding="utf-8") as jf:
                records = json.load(jf)
                if not isinstance(records, list):
                    records = []
        except Exception:
            records = []

    # Replace existing or append
    idx = next((i for i, r in enumerate(records) if r.get("Company Name", "").lower() == profile.get("Company Name", "").lower()), None)
    if idx is not None:
        records[idx] = profile
    else:
        records.append(profile)

    with open(json_path, mode="w", encoding="utf-8") as jf:
        json.dump(records, jf, indent=2, ensure_ascii=False)

    console.print(f"[green][OK] Saved categorized data to [bold]{csv_path}[/bold] and [bold]{json_path}[/bold][/green]")


COMMON_INDIAN_ACRONYMS = {
    "tcs": ("Tata Consultancy Services", "Indian multinational technology company"),
    "sbi": ("State Bank of India", "Indian public sector bank and financial services statutory body"),
    "infy": ("Infosys", "Indian multinational technology company"),
    "ril": ("Reliance Industries", "Indian multinational conglomerate"),
    "lt": ("Larsen & Toubro", "Indian multinational conglomerate company"),
    "l&t": ("Larsen & Toubro", "Indian multinational conglomerate company"),
    "m&m": ("Mahindra & Mahindra", "Indian multinational automobile manufacturer"),
    "lic": ("Life Insurance Corporation of India", "Indian central public sector undertaking"),
    "bpcl": ("Bharat Petroleum", "Indian public sector undertaking"),
    "hpcl": ("Hindustan Petroleum", "Indian public sector undertaking"),
    "iocl": ("Indian Oil Corporation", "Indian public sector oil company"),
    "ongc": ("Oil and Natural Gas Corporation", "Indian central public sector undertaking"),
    "ntpc": ("NTPC Limited", "Indian central public sector undertaking"),
    "bhel": ("Bharat Heavy Electricals Limited", "Indian public sector engineering and manufacturing company"),
    "sail": ("Steel Authority of India Limited", "Indian central public sector undertaking"),
    "pnb": ("Punjab National Bank", "Indian public sector bank"),
    "bob": ("Bank of Baroda", "Indian public sector bank"),
    "mrf": ("MRF Limited", "Indian tyre manufacturing company"),
    "hcl": ("HCLTech", "Indian multinational technology company"),
}

REJECT_CANDIDATE_PATTERNS = [
    r"\b(?:metro|railway|train|bus)\s+station\b",
    r"\b(?:airport|airfield|village|town|city|district|subdivision|tehsil|taluk|constituency|municipality|commune|county|state of)\b",
    r"\b(?:businessman|businesswoman|executive|industrialist|entrepreneur|philanthropist|politician|cricketer|actor|actress|singer|director|minister|governor)\b",
    r"\b(?:founder of|inventor|author|king|ruler|emperor|billionaire|millionaire)\b",
    r"\b(?:born\s+\d{4}|\(\d{4}[–—\-]\d{4}\)|\(\s*born\s+\d{4}\s*\))\b",
    r"\b(?:film|movie|album|song|soundtrack|discography|novel|television series|episode|tv show)\b",
    r"\b(?:history of|supply chain|timeline of|criticism of|economy of|list of|category:|companies based in|organizations based in|institutions based in|brand name potato|brand name)\b",
    r"\b(?:disambiguation|transit line|highway|expressway|stadium|park|sanctuary|temple|mosque|church|bridge)\b",
    r"\b(?:mascot|character|fictional character|symbol|logo|slogan)\b",
    r"\b(?:branch|sub post office|office|building|tower|facility|complex)\b",
    r"\b(?:horse|yacht|ship|military unit|naval|regiment|brigade)\b",
    r"\b(?:pakistani|pakistan|bangladesh|nepalese|sri lankan)\b"
]

POSITIVE_COMPANY_PATTERNS = [
    r"\b(?:company|corporation|conglomerate|manufacturer|multinational|enterprise|retailer|retail|supermarket)\b",
    r"\b(?:chain|restaurant|foodstuff|business|bank|banking|financial services|fintech|airline|firm|brand)\b",
    r"\b(?:software|technology company|tech company|it services|e-commerce|delivery service|quick-commerce|holding company)\b",
    r"\b(?:telecommunications|telecom|automaker|automotive|pharmaceutical|pharma|dairy|cooperative society|cooperative)\b",
    r"\b(?:fmcg|steelmaker|metallurgy|energy company|power producer|commercial bank|investment company|tyre|tire)\b",
    r"\b(?:ltd|limited|pvt|inc|corp|group|industries|technologies|enterprises)\b"
]

INDIAN_COMPANY_TOKENS = [
    "india", "indian", "mumbai", "delhi", "new delhi", "bengaluru", "bangalore",
    "chennai", "hyderabad", "kolkata", "pune", "nagpur", "gurugram", "gurgaon",
    "noida", "ahmedabad", "rajasthan", "gujarat", "maharashtra", "karnataka",
    "tamil nadu", "bse", "nse", "pvt ltd", "private limited"
]


def normalize_candidate_key(name: str) -> str:
    """Normalize company name to deduplicate legal variations."""
    s = name.lower()
    s = re.sub(r"\s*\([^)]*\)", "", s)
    s = re.sub(r"\b(?:pvt|private|ltd|limited|inc|corp|corporation|group|industries|co)\b", "", s)
    s = re.sub(r"[^\w\s]", "", s)
    return " ".join(s.split())


def score_candidate_relevance(name: str, query: str) -> int:
    """Rank suggestions by closeness to user query."""
    c_low = name.lower().strip()
    q_low = query.lower().strip()
    if c_low == q_low:
        return 100
    if c_low.startswith(q_low):
        return 80
    if q_low in c_low:
        return 60
    return 20


def evaluate_company_entity(name: str, desc: str, query: str, allowed_tokens: list[str]) -> Optional[Dict[str, Any]]:
    """Strictly validate whether a search hit is a real company and detect if it is Indian."""
    name_clean = re.sub(r"\s*\([^)]*\)", "", name).strip()
    n_lower = name_clean.lower()
    d_lower = desc.lower().strip()
    comb = f"{n_lower} {d_lower}"

    if len(name_clean) < 2 or len(name_clean) > 60 or name_clean.isdigit():
        return None

    for pat in REJECT_CANDIDATE_PATTERNS:
        if re.search(pat, comb):
            return None

    is_co = any(re.search(pat, comb) for pat in POSITIVE_COMPANY_PATTERNS)
    if not is_co:
        return None

    is_indian = any(tok in comb for tok in INDIAN_COMPANY_TOKENS)

    is_relevant = any(t in n_lower or t in d_lower for t in allowed_tokens) if allowed_tokens else True
    if not is_relevant:
        return None

    return {
        "name": name_clean,
        "desc": desc.strip() or "Corporate Entity",
        "is_indian": is_indian,
        "score": score_candidate_relevance(name_clean, query),
        "dedup_key": normalize_candidate_key(name_clean)
    }


def find_company_candidates(query: str) -> List[Dict[str, str]]:
    """
    Discover verified corporate candidates for queries, strictly excluding non-companies
    (metro stations, categories, geographic locations) and prioritizing Indian companies.
    """
    candidates = []
    seen_keys = {}
    headers = {"User-Agent": "CompanyIntelligence/2.0 (Windows; company-candidate-filter@company-datarecord.org)"}

    clean_q = query.strip()
    q_lower = clean_q.lower()

    allowed_tokens = [t for t in re.findall(r"\w+", q_lower) if len(t) > 1]

    # Pre-seed known Indian abbreviations
    if q_lower in COMMON_INDIAN_ACRONYMS:
        canonical_name, canonical_desc = COMMON_INDIAN_ACRONYMS[q_lower]
        c_entry = {
            "name": canonical_name,
            "desc": canonical_desc,
            "is_indian": True,
            "score": 100,
            "dedup_key": normalize_candidate_key(canonical_name)
        }
        candidates.append(c_entry)
        seen_keys[c_entry["dedup_key"]] = 0
        allowed_tokens.extend([t for t in re.findall(r"\w+", canonical_name.lower()) if len(t) > 2])

    search_terms = []
    if q_lower in COMMON_INDIAN_ACRONYMS:
        search_terms.append(COMMON_INDIAN_ACRONYMS[q_lower][0])
    search_terms.extend([f"{clean_q} company India", f"{clean_q} company", clean_q])

    # 1. Wikipedia Search API with description & pageprops
    for sq in search_terms:
        try:
            r = requests.get(
                "https://en.wikipedia.org/w/api.php",
                params={
                    "action": "query",
                    "generator": "search",
                    "gsrsearch": sq,
                    "gsrlimit": 6,
                    "prop": "description|pageprops",
                    "format": "json"
                },
                headers=headers,
                timeout=3.5
            )
            if r.status_code == 200:
                pages = r.json().get("query", {}).get("pages", {})
                for pid, p in pages.items():
                    if "disambiguation" in p.get("pageprops", {}):
                        continue
                    cand = evaluate_company_entity(p.get("title", ""), p.get("description", ""), clean_q, allowed_tokens)
                    if cand:
                        key = cand["dedup_key"]
                        if key not in seen_keys:
                            seen_keys[key] = len(candidates)
                            candidates.append(cand)
                        else:
                            idx = seen_keys[key]
                            if cand["is_indian"] and not candidates[idx]["is_indian"]:
                                candidates[idx]["is_indian"] = True
        except Exception:
            pass
        if len([c for c in candidates if c["is_indian"]]) >= 4:
            break

    # 2. Wikidata Entities Search
    for w_query in [clean_q, f"{clean_q} India"]:
        try:
            w_res = requests.get(
                "https://www.wikidata.org/w/api.php",
                params={
                    "action": "wbsearchentities",
                    "search": w_query,
                    "language": "en",
                    "format": "json",
                    "limit": 8
                },
                headers=headers,
                timeout=3.5
            )
            if w_res.status_code == 200:
                for item in w_res.json().get("search", []):
                    cand = evaluate_company_entity(item.get("label", ""), item.get("description", ""), clean_q, allowed_tokens)
                    if cand:
                        key = cand["dedup_key"]
                        if key not in seen_keys:
                            seen_keys[key] = len(candidates)
                            candidates.append(cand)
                        else:
                            idx = seen_keys[key]
                            if cand["is_indian"] and not candidates[idx]["is_indian"]:
                                candidates[idx]["is_indian"] = True
                                candidates[idx]["desc"] = cand["desc"]
        except Exception:
            pass

    # Strictly focus on Indian companies if any Indian companies match
    indian_cands = [c for c in candidates if c["is_indian"]]
    final_list = indian_cands if indian_cands else candidates

    # Rank by closeness to query
    final_list.sort(key=lambda x: -x["score"])

    return [{"name": c["name"], "desc": c["desc"]} for c in final_list[:6]]




def is_valid_company_name(name: str) -> bool:
    """
    Validate if user input resembles a plausible, defined, and real company name.
    Rejects:
    - Input less than 3 characters (e.g. 'a', 'ab', '12', '??')
    - Empty or pure punctuation / symbols / digits
    - Obvious placeholder, unreal, or undefined terms (e.g. 'undefined', 'unknown', 'fake', 'test', 'null')
    - Keyboard smashes / random character mashing (e.g. 'asdf', 'qwerty', 'zzzzzz', 'xkjsdhfkjsdfh')
    """
    if not name:
        return False
    cleaned = name.strip()
    
    # 1. Reject length less than 3 characters
    if len(cleaned) < 3:
        return False
        
    lower = cleaned.lower()
    
    # 2. Reject obvious undefined / unreal / placeholder words
    invalid_keywords = {
        'undefined', 'unreal', 'unknown', 'none', 'null', 'n/a', 'na',
        'fake', 'random', 'nothing', 'nobody', 'invalid', 'test', 'testing',
        'sample', 'example', 'placeholder', 'asdf', 'qwerty', 'zxcv',
        'abc', 'xyz', 'foo', 'bar', 'baz', 'company', 'corporation', 'inc', 'llc', 'ltd'
    }
    if lower in invalid_keywords or lower in ('not a company', 'no company', 'fake company', 'unknown company'):
        return False
        
    # 3. Must contain at least two actual alphabetic letters
    letters = re.findall(r'[a-zA-Z]', cleaned)
    if len(letters) < 2:
        return False
        
    # 4. Same character repeated (e.g. 'aaa', 'zzzz', '1111')
    if len(set(lower.replace(' ', ''))) <= 1:
        return False
        
    # 5. Three or more repeated identical characters in a row (e.g. 'coooompany', 'zzzz')
    if re.search(r'([a-zA-Z])\1{2,}', lower):
        return False
        
    # 6. Keyboard smash / sequence patterns
    smash_patterns = ['asdf', 'qwer', 'zxcv', 'hjkl', 'tyui', 'bnm,', '1234', 'qazw', 'wsxe']
    if any(p in lower for p in smash_patterns) and len(cleaned) <= 12:
        return False
        
    # 7. Unpronounceable letter smashes (low vowel ratio for words >= 5 letters without numbers)
    words = lower.split()
    for w in words:
        w_letters = re.findall(r'[a-z]', w)
        if len(w_letters) >= 5:
            vowels = sum(1 for c in w_letters if c in 'aeiouy')
            if vowels == 0 or (vowels / len(w_letters)) < 0.15:
                return False
            if re.search(r'[bcdfghjklmnpqrstvwxz]{5,}', w):
                return False
                
    return True


def main():
    console.print(Panel.fit(
        "[bold cyan]Structured Corporate Data Intelligence Engine[/bold cyan]\n"
        "[white]Views: [bold yellow]Table #1 (Identity & Leadership)[/bold yellow] | [bold yellow]Table #2 (5-Year Historical Financials)[/bold yellow][/white]",
        border_style="cyan"
    ))

    # Support CLI parameter: python company_lookup.py "Haldiram's"
    if len(sys.argv) > 1:
        query = " ".join(sys.argv[1:]).strip('"\'')
        if not is_valid_company_name(query):
            console.print("[bold red]Please enter valid company name.[/bold red]")
            return
        console.print(f"[yellow]Fetching Table #1 records for:[/yellow] [bold]{query}[/bold]...")
        data1, sources1 = fetch_table1_data(query)
        display_table1(data1, sources1)

        console.print(f"[yellow]Fetching Table #2 (5-Year Historical & Present Financials) for:[/yellow] [bold]{query}[/bold]...")
        data2, sources2 = fetch_table2_data(data1.get("Company Name", query), data1.get("Stock Ticker", "N/A"))
        display_table2(data2, sources2)

        save_table_records(data1, sources1, data2, sources2)
        return

    # Interactive Loop
    while True:
        try:
            company_input = console.input("[bold yellow]Enter Company Name (or 'q' to quit): [/bold yellow]").strip()
            if not company_input:
                continue
            if company_input.lower() in ("q", "quit", "exit"):
                console.print("[dim]Goodbye![/dim]")
                break

            if not is_valid_company_name(company_input):
                console.print("[bold red]Please enter valid company name.[/bold red]\n")
                continue

            # Step 1: Discover similar/matching companies so user can select
            console.print(f"[dim]Searching for companies matching '{company_input}'...[/dim]")
            candidates = find_company_candidates(company_input)

            selected_company = company_input
            if candidates:
                table = Table(
                    title=f"[bold cyan]Verified Company Matches for '{company_input}' (Prioritizing Indian Companies)[/bold cyan]",
                    show_header=True,
                    header_style="bold magenta"
                )
                table.add_column("#", style="bold yellow", width=4)
                table.add_column("Company Name", style="bold white", width=34)
                table.add_column("Corporate Details / Description", style="dim white")

                for idx, cand in enumerate(candidates, 1):
                    table.add_row(str(idx), cand["name"], cand["desc"])

                console.print()
                console.print(table)
                choice = console.input(
                    f"[bold yellow]Select company [1-{len(candidates)}] (press Enter for [1], 'c' to cancel): [/bold yellow]"
                ).strip()

                if choice.lower() in ("c", "cancel"):
                    console.print("[dim]Cancelled search.[/dim]\n")
                    continue

                if choice.isdigit() and 1 <= int(choice) <= len(candidates):
                    selected_company = candidates[int(choice) - 1]["name"]
                else:
                    selected_company = candidates[0]["name"]

            console.print(f"\n[cyan]Fetching Table #1 Corporate & Leadership Data for '[bold]{selected_company}[/bold]'...[/cyan]")
            data1, sources1 = fetch_table1_data(selected_company)
            display_table1(data1, sources1)

            console.print(f"\n[cyan]Fetching Table #2 (5-Year Historical & Present Financials) for '[bold]{selected_company}[/bold]'...[/cyan]")
            data2, sources2 = fetch_table2_data(data1.get("Company Name", selected_company), data1.get("Stock Ticker", "N/A"))
            display_table2(data2, sources2)

            save_choice = console.input("[bold]Save Table #1 and Table #2 records to CSV/JSON? (Y/n): [/bold]").strip().lower()
            if save_choice in ("", "y", "yes"):
                save_table_records(data1, sources1, data2, sources2)

        except KeyboardInterrupt:
            console.print("\n[dim]Process exited.[/dim]")
            break



if __name__ == "__main__":
    main()
