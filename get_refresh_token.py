"""
get_refresh_token.py
Run this ONCE locally to generate a refresh token for GitHub Actions.
It opens a browser window for you to sign in with your Microsoft account.
The refresh token is then stored and used by GitHub Actions for unattended runs.

Usage:
    python get_refresh_token.py

Requirements:
    pip install msal
"""

import json
import msal

# ── Fill these in before running ──────────────────────────────────────────────
CLIENT_ID = "YOUR_CLIENT_ID_HERE"   # From Azure App Registration
TENANT_ID = "YOUR_TENANT_ID_HERE"   # From Azure App Registration
# If no app registration, use Microsoft's public client ID:
# CLIENT_ID = "d3590ed6-52b3-4102-aeff-aad2292ab01c"  # Microsoft Office
# TENANT_ID = "consumers"
# ─────────────────────────────────────────────────────────────────────────────

SCOPES = [
    "https://graph.microsoft.com/Notes.ReadWrite",
    "https://graph.microsoft.com/Notes.Read.All",
    "offline_access",
]

def main():
    print("=" * 60)
    print("OneNote Refresh Token Generator")
    print("=" * 60)
    print()

    app = msal.PublicClientApplication(
        client_id=CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
    )

    # Device code flow - works without a browser on the machine
    print("Starting device code flow...")
    print("You will be given a code to enter at https://microsoft.com/devicelogin")
    print()

    flow = app.initiate_device_flow(scopes=SCOPES)

    if "user_code" not in flow:
        raise ValueError(f"Failed to create device flow: {flow}")

    print("=" * 60)
    print(flow["message"])
    print("=" * 60)
    print()
    print("Waiting for you to sign in...")

    result = app.acquire_token_by_device_flow(flow)

    if "access_token" not in result:
        print(f"\n❌ Authentication failed: {result.get('error_description', result)}")
        return

    print("\n✅ Authentication successful!")
    print()

    refresh_token = result.get("refresh_token", "")

    if not refresh_token:
        print("⚠️  No refresh token returned.")
        print("Make sure 'offline_access' is in your scopes and")
        print("'Allow public client flows' is enabled in your app registration.")
        return

    # Save tokens to file for reference
    token_data = {
        "access_token":  result["access_token"],
        "refresh_token": refresh_token,
        "token_type":    result.get("token_type"),
        "scope":         result.get("scope"),
    }

    with open("tokens.json", "w") as f:
        json.dump(token_data, f, indent=2)

    print("=" * 60)
    print("YOUR REFRESH TOKEN (copy this into GitHub Secrets):")
    print("=" * 60)
    print()
    print(refresh_token)
    print()
    print("=" * 60)
    print()
    print("Next steps:")
    print("1. Copy the refresh token above")
    print("2. Go to: https://github.com/peter-kenneth-farrar/Projects")
    print("3. Click: Settings → Secrets and variables → Actions")
    print("4. Add these secrets one by one:")
    print()
    print(f"   Secret Name      │ Value")
    print(f"   ─────────────────┼──────────────────────────────")
    print(f"   TENANT_ID        │ {TENANT_ID}")
    print(f"   CLIENT_ID        │ {CLIENT_ID}")
    print(f"   REFRESH_TOKEN    │ <paste refresh token above>")
    print(f"   ONENOTE_NOTEBOOK │ peter farrar")
    print(f"   ONENOTE_SECTION  │ Microsoft-Test")
    print(f"   ONENOTE_PAGE     │ microsoft-test")
    print()
    print("5. Push project to GitHub")
    print("6. Go to Actions tab → Run workflow manually to test")
    print()
    print("Token also saved to tokens.json (keep this file private!)")
    print("⚠️  DO NOT commit tokens.json to GitHub!")


if __name__ == "__main__":
    main()