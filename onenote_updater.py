"""
onenote_updater.py
Authenticates with Microsoft Graph and replaces the target OneNote page
with freshly scraped Excel known-issues content, mirroring the webpage layout.
"""

import os
import json
import logging
import requests
from datetime import datetime

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
logger = logging.getLogger(__name__)

# ── Graph endpoints ────────────────────────────────────────────────────────────
GRAPH_BASE = "https://graph.microsoft.com/v1.0"

# ── Status emoji map ──────────────────────────────────────────────────────────
STATUS_EMOJI = {
    "FIXED":        "✅",
    "INVESTIGATING": "🔴",
    "WORKAROUND":   "🟡",
    "BY DESIGN":    "🔵",
    "RESOLVED":     "✅",
    "UPDATED":      "🔄",
    "SUSPENDED":    "⏸️",
    "KNOWN ISSUE":  "⚠️",
}


def get_access_token() -> str:
    """
    Get access token using client credentials (app registration)
    or refresh token flow depending on available env vars.
    """
    tenant_id    = os.environ["TENANT_ID"]
    client_id    = os.environ["CLIENT_ID"]
    refresh_token = os.environ.get("REFRESH_TOKEN")
    client_secret = os.environ.get("CLIENT_SECRET")

    token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"

    if refresh_token:
        # Delegated flow using stored refresh token (personal account)
        logger.info("Authenticating via refresh token...")
        data = {
            "grant_type":    "refresh_token",
            "client_id":     client_id,
            "refresh_token": refresh_token,
            "scope":         "https://graph.microsoft.com/Notes.ReadWrite offline_access",
        }
        if client_secret:
            data["client_secret"] = client_secret

    elif client_secret:
        # App-only flow (requires admin consent)
        logger.info("Authenticating via client credentials...")
        data = {
            "grant_type":    "client_credentials",
            "client_id":     client_id,
            "client_secret": client_secret,
            "scope":         "https://graph.microsoft.com/.default",
        }
    else:
        raise EnvironmentError(
            "Set either REFRESH_TOKEN or CLIENT_SECRET environment variable."
        )

    resp = requests.post(token_url, data=data, timeout=30)
    resp.raise_for_status()
    token_data = resp.json()

    if "access_token" not in token_data:
        raise RuntimeError(f"Token error: {token_data}")

    logger.info("Access token obtained successfully.")
    return token_data["access_token"]


def graph_get(token: str, path: str) -> dict:
    """GET request to Microsoft Graph."""
    resp = requests.get(
        f"{GRAPH_BASE}{path}",
        headers={"Authorization": f"Bearer {token}"},
        timeout=30,
    )
    resp.raise_for_status()
    return resp.json()


def find_page(token: str, notebook_name: str, section_name: str, page_name: str) -> str:
    """
    Find the OneNote page ID by notebook → section → page name.
    Returns the page ID string.
    """
    # 1. Find notebook
    logger.info(f"Looking for notebook: '{notebook_name}'")
    nb_data = graph_get(token, "/me/onenote/notebooks")
    notebook = next(
        (n for n in nb_data["value"] if n["displayName"].lower() == notebook_name.lower()),
        None,
    )
    if not notebook:
        available = [n["displayName"] for n in nb_data["value"]]
        raise ValueError(f"Notebook '{notebook_name}' not found. Available: {available}")
    logger.info(f"  Found notebook: {notebook['displayName']} ({notebook['id']})")

    # 2. Find section
    logger.info(f"Looking for section: '{section_name}'")
    sec_data = graph_get(token, f"/me/onenote/notebooks/{notebook['id']}/sections")
    section = next(
        (s for s in sec_data["value"] if s["displayName"].lower() == section_name.lower()),
        None,
    )
    if not section:
        available = [s["displayName"] for s in sec_data["value"]]
        raise ValueError(f"Section '{section_name}' not found. Available: {available}")
    logger.info(f"  Found section: {section['displayName']} ({section['id']})")

    # 3. Find page
    logger.info(f"Looking for page: '{page_name}'")
    pg_data = graph_get(token, f"/me/onenote/sections/{section['id']}/pages")
    page = next(
        (p for p in pg_data["value"] if p["title"].lower() == page_name.lower()),
        None,
    )
    if not page:
        available = [p["title"] for p in pg_data["value"]]
        raise ValueError(f"Page '{page_name}' not found. Available: {available}")
    logger.info(f"  Found page: {page['title']} ({page['id']})")

    return page["id"]


def build_html(data: dict) -> str:
    """
    Build OneNote-compatible HTML that mirrors the Microsoft support page layout.
    Sections as H2 headings, issues as bulleted lists with status tags and links.
    """
    scraped_at = data.get("scraped_at", "Unknown")
    source_url = data.get("source_url", "")
    total      = data.get("total_issues", 0)

    lines = [
        '<!DOCTYPE html>',
        '<html>',
        '<head><meta charset="utf-8"/></head>',
        '<body>',

        # Page header
        f'<h1>Fixes or workarounds for recent issues in Excel for Windows</h1>',
        f'<p><strong>Source:</strong> <a href="{source_url}">{source_url}</a></p>',
        f'<p><strong>Last updated:</strong> {scraped_at} &nbsp;|&nbsp; '
        f'<strong>Total issues:</strong> {total}</p>',
        '<hr/>',
    ]

    for section in data.get("sections", []):
        heading = section["heading"]
        issues  = section["issues"]

        lines.append(f'<h2>{heading}</h2>')
        lines.append('<ul>')

        for issue in issues:
            status  = issue["status"]
            title   = issue["title"]
            url     = issue["url"]
            emoji   = STATUS_EMOJI.get(status, "⚠️")

            # Escape any HTML special chars in title
            safe_title = (
                title
                .replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;")
                .replace('"', "&quot;")
            )

            lines.append(
                f'<li>{emoji} <strong>[{status}]</strong> '
                f'<a href="{url}">{safe_title}</a></li>'
            )

        lines.append('</ul>')

    lines += ['</body>', '</html>']
    return "\n".join(lines)


def replace_page_content(token: str, page_id: str, html_content: str) -> None:
    """
    Replace the entire content of a OneNote page using the PATCH endpoint.
    Uses 'replace' action on the body element.
    """
    logger.info(f"Replacing content of page {page_id}...")

    patch_url = f"{GRAPH_BASE}/me/onenote/pages/{page_id}/content"

    # Extract just the body content for the PATCH command
    from bs4 import BeautifulSoup
    soup = BeautifulSoup(html_content, "html.parser")
    body = soup.find("body")
    body_html = body.decode_contents() if body else html_content

    patch_commands = [
        {
            "target": "body",
            "action": "replace",
            "content": body_html,
        }
    ]

    resp = requests.patch(
        patch_url,
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type":  "application/json",
        },
        data=json.dumps(patch_commands),
        timeout=30,
    )

    if resp.status_code == 204:
        logger.info("✅ Page content replaced successfully.")
    else:
        logger.error(f"PATCH failed: {resp.status_code} — {resp.text}")
        resp.raise_for_status()


def update_onenote(data: dict) -> None:
    """
    Main entry point: authenticate, find the page, replace its content.
    Reads config from environment variables:
        TENANT_ID, CLIENT_ID, REFRESH_TOKEN (or CLIENT_SECRET)
        ONENOTE_NOTEBOOK, ONENOTE_SECTION, ONENOTE_PAGE
    """
    notebook_name = os.environ.get("ONENOTE_NOTEBOOK", "peter farrar")
    section_name  = os.environ.get("ONENOTE_SECTION", "")
    page_name     = os.environ.get("ONENOTE_PAGE",     "microsoft-test")

    if not section_name:
        raise EnvironmentError("Set ONENOTE_SECTION environment variable.")

    token   = get_access_token()
    page_id = find_page(token, notebook_name, section_name, page_name)
    html    = build_html(data)
    replace_page_content(token, page_id, html)


if __name__ == "__main__":
    # Quick local test — load previously scraped data
    with open("issues_extracted.json", "r", encoding="utf-8") as f:
        data = json.load(f)
    update_onenote(data)