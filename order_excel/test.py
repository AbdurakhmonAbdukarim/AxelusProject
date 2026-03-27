# =============================================================================
# get_zoho_token.py
# Run this ONCE to exchange your code for access + refresh tokens.
#
# Usage:
#   1. Fill in CLIENT_ID, CLIENT_SECRET, CODE below
#   2. Run: python get_zoho_token.py
#   3. Copy the tokens into config.py
# =============================================================================

import requests

# ── Fill these in ────────────────────────────────────────────────────────────
CLIENT_ID     = "1000.HKSTC4EOA2K294N9605P78PTSX96YQ"       # from api-console.zoho.com Self Client
CLIENT_SECRET = "220f4a58e520e65016f3884baf19678c2dafe8fd95"   # from api-console.zoho.com Self Client
CODE          = "1000.23737a7a16f3f14475e2f04c039ff72e.3b4483de9f539e7e3610a2bad2b33f4e"            # from Generate Code tab (valid 10 min)
# ─────────────────────────────────────────────────────────────────────────────

url = "https://accounts.zoho.com/oauth/v2/token"

# Self Client does NOT need redirect_uri — that's why you got 500
params = {
    "code":          CODE,
    "client_id":     CLIENT_ID,
    "client_secret": CLIENT_SECRET,
    "grant_type":    "authorization_code",
}

print("Sending request to Zoho...")
r = requests.post(url, params=params)

print(f"Status: {r.status_code}")
print(f"Response: {r.text}")

if r.status_code == 200:
    data = r.json()
    print("\n" + "="*50)
    print("✅ SUCCESS! Copy these into config.py:")
    print("="*50)
    print(f"ZOHO_ACCESS_TOKEN  = \"{data.get('access_token')}\"")
    print(f"ZOHO_REFRESH_TOKEN = \"{data.get('refresh_token')}\"")
    print("="*50)
else:
    print("\n❌ Failed. Check your CLIENT_ID, CLIENT_SECRET and CODE.")
    print("Make sure the CODE was generated within the last 10 minutes.")