"""
scraper.py
Scrapes the Microsoft Excel known issues page and returns structured data.
"""

import requests
from bs4 import BeautifulSoup
import json
import logging
from datetime import datetime

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
logger = logging.getLogger(__name__)

SOURCE_URL = (
    "https://support.microsoft.com/en-us/office/"
    "fixes-or-workarounds-for-recent-issues-in-excel-for-windows-"
    "49d932ce-0240-49cf-94df-1587d9d97093"
)

BASE_URL = "https://support.microsoft.com"

TARGET_SECTIONS = [
    "Excel crashes and slow performance issues",
    "Excel features and add-ins issues",
    "Known issues, changed functionality, and blocked or discontinued features",
]

STATUS_TAGS = [
    "[FIXED]",
    "[INVESTIGATING]",
    "[WORKAROUND]",
    "[BY DESIGN]",
    "[RESOLVED]",
    "[UPDATED]",
    "[SUSPENDED]",
]

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.5",
}


def extract_status(text: str):
    """Extract status tag and clean title from raw list item text."""
    text = text.strip()
    upper = text.upper()
    for tag in STATUS_TAGS:
        if upper.startswith(tag.upper()):
            status = tag.strip("[]")
            title = text[len(tag):].strip().lstrip("\u00a0").strip()
            return status, title
    return "KNOWN ISSUE", text


def resolve_url(href: str) -> str:
    """Make relative URLs absolute."""
    if not href:
        return SOURCE_URL
    if href.startswith("http"):
        return href
    if href.startswith("/"):
        return BASE_URL + href
    return BASE_URL + "/en-us/office/" + href


def scrape() -> dict:
    """
    Scrape the Microsoft Excel issues page.
    Returns a dict with metadata and sections list.
    """
    logger.info(f"Fetching: {SOURCE_URL}")
    response = requests.get(SOURCE_URL, headers=HEADERS, timeout=30)
    response.raise_for_status()
    logger.info(f"HTTP {response.status_code} — page size {len(response.text):,} chars")

    soup = BeautifulSoup(response.text, "html.parser")
    main = soup.find("main") or soup

    sections = []

    for h2 in main.find_all("h2"):
        section_title = h2.get_text(strip=True)
        if section_title not in TARGET_SECTIONS:
            continue

        issues = []
        sibling = h2.find_next_sibling()
        while sibling and sibling.name != "h2":
            if sibling.name == "ul":
                for li in sibling.find_all("li"):
                    raw_text = li.get_text(separator=" ", strip=True)
                    link_tag = li.find("a")
                    href = resolve_url(link_tag.get("href", "") if link_tag else "")
                    status, title = extract_status(raw_text)
                    issues.append({
                        "status": status,
                        "title": title,
                        "url": href,
                    })
            sibling = sibling.find_next_sibling()

        if issues:
            sections.append({
                "heading": section_title,
                "issues": issues,
            })
            logger.info(f"  Section '{section_title}': {len(issues)} issues")

    total = sum(len(s["issues"]) for s in sections)
    logger.info(f"Total issues scraped: {total}")

    return {
        "source_url": SOURCE_URL,
        "scraped_at": datetime.utcnow().strftime("%Y-%m-%d %H:%M UTC"),
        "total_issues": total,
        "sections": sections,
    }


if __name__ == "__main__":
    data = scrape()
    with open("issues_extracted.json", "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
    print(f"\nDone — {data['total_issues']} issues saved to issues_extracted.json")