#!/usr/bin/env python3
"""Device-code login against the 'Microsoft Graph Command Line Tools' public client."""
import json
import os
import sys
import msal

KNOWN_CLIENTS = {
    "graphcli": "14d82eec-204b-4c2f-b7e8-296a70dab67e",  # Microsoft Graph Command Line Tools
    "teams":    "1fec8e78-bce4-4aaf-ab1b-5451cc387264",  # Microsoft Teams
    "office":   "d3590ed6-52b3-4102-aeff-aad2292ab01c",  # Microsoft Office
    "teamsweb": "5e3ce6c0-2b1f-4285-8d4b-75ee78787346",  # Teams Web Client
}
_sel = os.environ.get("GRAPH_CLIENT", "teams")
CLIENT_ID = KNOWN_CLIENTS.get(_sel, _sel)
AUTHORITY = "https://login.microsoftonline.com/organizations"
SCOPES = os.environ.get("GRAPH_SCOPES", "Chat.ReadWrite").split()
print(f"[auth] client={_sel} ({CLIENT_ID}) scopes={SCOPES}", file=sys.stderr)
CACHE_PATH = "token_cache.json"


def load_cache():
    cache = msal.SerializableTokenCache()
    try:
        with open(CACHE_PATH, "r") as f:
            cache.deserialize(f.read())
    except FileNotFoundError:
        pass
    return cache


def save_cache(cache):
    if cache.has_state_changed:
        with open(CACHE_PATH, "w") as f:
            f.write(cache.serialize())


def main():
    cache = load_cache()
    app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY, token_cache=cache)

    result = None
    accounts = app.get_accounts()
    if accounts:
        print(f"Found cached account: {accounts[0].get('username')}", file=sys.stderr)
        result = app.acquire_token_silent(SCOPES, account=accounts[0])

    if not result:
        flow = app.initiate_device_flow(scopes=SCOPES)
        if "user_code" not in flow:
            raise RuntimeError(f"Failed to create device flow: {json.dumps(flow, indent=2)}")
        print(flow["message"], flush=True)
        result = app.acquire_token_by_device_flow(flow)

    if "access_token" not in result:
        print(json.dumps(result, indent=2), file=sys.stderr)
        raise SystemExit("Auth failed")

    save_cache(cache)
    print(f"\n✔ Signed in. Token expires in {result.get('expires_in')}s.", file=sys.stderr)
    print(f"✔ Account: {result.get('id_token_claims', {}).get('preferred_username')}", file=sys.stderr)


if __name__ == "__main__":
    main()
