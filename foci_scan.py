#!/usr/bin/env python3
"""Redeem the cached family refresh token as each known FOCI client and dump Graph scopes."""
import base64
import json
import msal

CACHE_PATH = "token_cache.json"
AUTHORITY = "https://login.microsoftonline.com/organizations"
SCOPE = ["https://graph.microsoft.com/.default"]

FOCI_CLIENTS = {
    "Microsoft Teams":                "1fec8e78-bce4-4aaf-ab1b-5451cc387264",
    "Microsoft Office":               "d3590ed6-52b3-4102-aeff-aad2292ab01c",
    "Office 365 Management":          "00b41c95-dab0-4487-9791-b9d2c32c80f2",
    "Outlook Mobile":                 "27922004-5251-4030-b22d-91ecd9a37ea4",
    "OneDrive SyncEngine":            "ab9b8c07-8f02-4f72-87fa-80105867a763",
    "Azure CLI":                      "04b07795-8ddb-461a-bbee-02f9e1bf7b46",
    "Azure PowerShell":               "1950a258-227b-4e31-a9cf-717495945fc2",
    "Visual Studio":                  "872cd9fa-d31f-45e0-9eab-6e460a02d1f1",
    "Microsoft Edge":                 "ecd6b820-32c2-49b6-98a6-444530e5a77a",
    "Microsoft Edge (2)":             "e9c51622-460d-4d3d-952d-966a5b1da34c",
    "Microsoft Bing Search (Edge)":   "2d7f3606-b07d-41d1-b9d2-0d0c9296a6e8",
    "Microsoft To-Do":                "22098786-6e16-43cc-a27d-191a01a1e3b5",
    "Windows Search":                 "26a7ee05-5602-4d76-a7ba-eae8b7b67941",
    "Microsoft Whiteboard":           "57336123-6e14-4acc-8dcf-287b6088aa28",
    "Microsoft Authenticator":        "4813382a-8fa7-425e-ab75-3b753aab3abb",
    "Microsoft Power BI":             "c0d2a505-13b8-4ae0-aa9e-cddd5eab0b12",
    "Accounts Control UI":            "a40d7d7d-59aa-447e-a655-679a4107e548",
    "Microsoft Flow":                 "57fcbcfa-7cee-4eb1-8b25-12d2030b4ee0",
    "OneDrive iOS":                   "af124e86-4e96-495a-b70a-90f90ab96707",
    "Microsoft Stream Mobile":        "844cca35-0656-46ce-b636-13f48b0eecbd",
    "SharePoint":                     "d326c1ce-6cc6-4de2-bebc-4591e5e13ef0",
    "Microsoft Planner":              "66375f6b-983f-4c2c-9701-d680650f588f",
    "Microsoft Intune Company Portal":"9ba1a5c7-f17a-4de9-a1f1-6178c8d51223",
    "Microsoft Tunnel":               "eb539595-3fe1-474e-9c1d-feb3625d1be5",
    "Yammer iPhone":                  "a569458c-7f2b-45cb-bab9-b7dee514d112",
    "M365 Compliance Drive":          "be1918be-3fe3-4be9-b32b-b542fc27f02e",
}


def decode_jwt_scopes(token):
    try:
        payload = token.split(".")[1]
        payload += "=" * (-len(payload) % 4)
        claims = json.loads(base64.urlsafe_b64decode(payload))
        return claims.get("scp", "")
    except Exception as e:
        return f"<decode error: {e}>"


def main():
    with open(CACHE_PATH) as f:
        raw = f.read()

    hits = []
    for name, cid in FOCI_CLIENTS.items():
        cache = msal.SerializableTokenCache()
        cache.deserialize(raw)
        app = msal.PublicClientApplication(cid, authority=AUTHORITY, token_cache=cache)
        accts = app.get_accounts()
        if not accts:
            print(f"{name:35} -- no account visible")
            continue
        result = app.acquire_token_silent_with_error(SCOPE, account=accts[0])
        if not result or "access_token" not in result:
            err = (result or {}).get("error") or (result or {}).get("error_description", "")
            print(f"{name:35} -- FAIL: {str(err)[:80]}")
            continue
        scp = decode_jwt_scopes(result["access_token"])
        chat_scopes = sorted(s for s in scp.split() if s.lower().startswith("chat"))
        marker = "  <== HAS Chat.Read!" if any("read" in s.lower() for s in chat_scopes) else ""
        print(f"{name:35} -- chat scopes: {chat_scopes or '-'}{marker}")
        if marker:
            hits.append((name, cid, chat_scopes))

    print("\n=== Summary ===")
    if hits:
        for name, cid, scopes in hits:
            print(f"  {name} ({cid}): {scopes}")
    else:
        print("  No FOCI client has Chat.Read* preauthorized on Graph.")


if __name__ == "__main__":
    main()
