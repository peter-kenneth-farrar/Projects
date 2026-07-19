# Microsoft Excel Issues → OneNote Updater

Automatically scrapes Microsoft's Excel known issues page and updates a OneNote page monthly, mirroring the original webpage layout.

## What It Does

1. Scrapes [Microsoft Excel Known Issues](https://support.microsoft.com/en-us/office/fixes-or-workarounds-for-recent-issues-in-excel-for-windows-49d932ce-0240-49cf-94df-1587d9d97093)
2. Extracts all issues across 3 sections with their status tags
3. Replaces your OneNote page with the latest content
4. Runs automatically on the **1st of every month at 08:00 UTC**

## OneNote Page Layout

Mirrors the Microsoft support page exactly:

```
Fixes or workarounds for recent issues in Excel for Windows
Source: https://support.microsoft.com/...
Last updated: 2026-07-01 08:00 UTC | Total issues: 40
─────────────────────────────────────────────────────────

Excel crashes and slow performance issues
  • ✅ [FIXED]         Word, Excel, or Outlook might stop responding...
  • 🔴 [INVESTIGATING] Error "Something Went Wrong [1001]"...
  • 🟡 [WORKAROUND]    Excel stops responding when using 3DxWare...

Excel features and add-ins issues
  • 🔴 [INVESTIGATING] Error "No results were found"...
  • ✅ [FIXED]         When trying to sign in to an Office app...

Known issues, changed functionality, and blocked or discontinued features
  • ⚠️  Flash, Silverlight, and Shockwave controls blocked...
```

## Status Icons

| Icon | Status |
|------|--------|
| ✅ | FIXED / RESOLVED |
| 🔴 | INVESTIGATING |
| 🟡 | WORKAROUND |
| 🔵 | BY DESIGN |
| 🔄 | UPDATED |
| ⏸️ | SUSPENDED |
| ⚠️ | KNOWN ISSUE |

---

## Setup Guide

### Step 1: Get a Refresh Token (One Time Only)

1. Install dependencies locally:
   ```bash
   pip install -r requirements.txt
   ```

2. Edit `get_refresh_token.py` and fill in your `CLIENT_ID` and `TENANT_ID`

3. Run it:
   ```bash
   python get_refresh_token.py
   ```

4. Follow the on-screen instructions:
   - Go to https://microsoft.com/devicelogin
   - Enter the code shown
   - Sign in with your Microsoft account
   - Copy the refresh token printed at the end

### Step 2: Add GitHub Secrets

Go to your GitHub repo → **Settings** → **Secrets and variables** → **Actions** → **New repository secret**

Add these secrets:

| Secret Name | Value |
|-------------|-------|
| `TENANT_ID` | Your Azure tenant ID |
| `CLIENT_ID` | Your Azure app client ID |
| `REFRESH_TOKEN` | Token from Step 1 |
| `ONENOTE_NOTEBOOK` | `peter farrar` |
| `ONENOTE_SECTION` | `Microsoft-Test` |
| `ONENOTE_PAGE` | `microsoft-test` |

### Step 3: Push to GitHub

```bash
git init
git remote add origin https://github.com/peter-kenneth-farrar/Projects.git
git add .
git commit -m "Initial commit - Excel issues OneNote updater"
git push -u origin main
```

### Step 4: Test It Manually

1. Go to your GitHub repo
2. Click **Actions** tab
3. Click **Monthly Excel Issues → OneNote Update**
4. Click **Run workflow** → **Run workflow**
5. Watch the logs — should complete in under 2 minutes
6. Check your OneNote **microsoft-test** page!

### Step 5: Verify Automatic Schedule

The workflow runs automatically on the **1st of every month at 08:00 UTC**.
You can verify it's scheduled under **Actions** → **Scheduled** runs.

---

## File Structure

```
project/
├── main.py                          ← Entry point
├── scraper.py                       ← Scrapes Microsoft page
├── onenote_updater.py               ← Writes to OneNote via Graph API
├── get_refresh_token.py             ← Run once to get auth token
├── requirements.txt                 ← Python dependencies
├── .gitignore                       ← Keeps tokens out of git
├── README.md                        ← This file
└── .github/
    └── workflows/
        └── monthly_update.yml       ← GitHub Actions schedule
```

## Running Locally

```bash
# Set environment variables
export TENANT_ID="your-tenant-id"
export CLIENT_ID="your-client-id"
export REFRESH_TOKEN="your-refresh-token"
export ONENOTE_NOTEBOOK="peter farrar"
export ONENOTE_SECTION="Microsoft-Test"
export ONENOTE_PAGE="microsoft-test"

# Run
python main.py
```

## Troubleshooting

| Error | Fix |
|-------|-----|
| `Notebook not found` | Check exact notebook name in OneNote |
| `Section not found` | Check exact section name — should be `Microsoft-Test` |
| `Page not found` | Check page is named exactly `microsoft-test` |
| `Token error` | Re-run `get_refresh_token.py` to get a fresh token |
| `HTTP 401` | Refresh token expired — re-run `get_refresh_token.py` |
| `HTTP 403` | Check API permissions in Azure app registration |

## Adding More Pages Later

To add more Microsoft known issues pages (e.g. Word, Teams), duplicate `scraper.py` with the new URL and add a new job in the workflow YAML.