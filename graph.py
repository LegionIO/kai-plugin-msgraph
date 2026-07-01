#!/usr/bin/env python3
"""Acquire a token silently from the cache and call Graph."""
import json
import os
import sys
import msal
import requests

KNOWN_CLIENTS = {
    "graphcli": "14d82eec-204b-4c2f-b7e8-296a70dab67e",
    "teams":    "1fec8e78-bce4-4aaf-ab1b-5451cc387264",
    "office":   "d3590ed6-52b3-4102-aeff-aad2292ab01c",
    "teamsweb": "5e3ce6c0-2b1f-4285-8d4b-75ee78787346",
}
_sel = os.environ.get("GRAPH_CLIENT", "teams")
CLIENT_ID = KNOWN_CLIENTS.get(_sel, _sel)
AUTHORITY = "https://login.microsoftonline.com/organizations"
SCOPES = os.environ.get("GRAPH_SCOPES", "Chat.ReadWrite").split()
CACHE_PATH = "token_cache.json"
GRAPH = "https://graph.microsoft.com/v1.0"


def get_token():
    cache = msal.SerializableTokenCache()
    with open(CACHE_PATH, "r") as f:
        cache.deserialize(f.read())
    app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY, token_cache=cache)
    accounts = app.get_accounts()
    if not accounts:
        raise SystemExit("No cached account. Run auth.py first.")
    result = app.acquire_token_silent(SCOPES, account=accounts[0])
    if not result or "access_token" not in result:
        raise SystemExit(f"Silent token acquisition failed: {result}. Run auth.py again.")
    if cache.has_state_changed:
        with open(CACHE_PATH, "w") as f:
            f.write(cache.serialize())
    return result["access_token"], accounts[0]


def gget(token, path, **params):
    r = requests.get(f"{GRAPH}{path}", headers={"Authorization": f"Bearer {token}"}, params=params)
    if not r.ok:
        print(f"GET {path} -> {r.status_code}\n{r.text}", file=sys.stderr)
        r.raise_for_status()
    return r.json()


def main():
    if len(sys.argv) < 2:
        raise SystemExit("usage: graph.py <search-name-or-email>")
    query = sys.argv[1]

    token, account = get_token()
    me = gget(token, "/me")
    my_id = me["id"]
    print(f"Me: {me.get('displayName')} <{me.get('userPrincipalName')}> ({my_id})\n")

    # Find the 1:1 chat by scanning member displayNames (avoids needing User.ReadBasic.All)
    q = query.lower()
    chats = gget(token, "/me/chats", **{"$expand": "members", "$filter": "chatType eq 'oneOnOne'", "$top": "50"})
    chat_id = None
    other_id = None
    other_name = None
    for c in chats.get("value", []):
        for m in c.get("members", []):
            name = (m.get("displayName") or "").lower()
            email = (m.get("email") or "").lower()
            if m.get("userId") != my_id and (q in name or q in email):
                chat_id = c["id"]
                other_id = m.get("userId")
                other_name = m.get("displayName")
                break
        if chat_id:
            break
    if not chat_id:
        raise SystemExit(f"No 1:1 chat found matching '{query}' in your 50 most recent oneOnOne chats.")
    print(f"Target: {other_name} ({other_id})")
    print(f"Chat:   {chat_id}\n")

    # Fetch recent messages, newest first
    msgs = gget(token, f"/chats/{chat_id}/messages", **{"$top": "20"})
    latest_sent = None
    latest_received = None
    for m in msgs.get("value", []):
        frm = (m.get("from") or {}).get("user") or {}
        sender_id = frm.get("id")
        if sender_id == my_id and latest_sent is None:
            latest_sent = m
        elif sender_id == other_id and latest_received is None:
            latest_received = m
        if latest_sent and latest_received:
            break

    def show(label, m):
        if not m:
            print(f"{label}: (none in last 20 messages)")
            return
        body = (m.get("body") or {}).get("content", "")
        print(f"{label}  [{m.get('createdDateTime')}]")
        print(f"  from: {(m.get('from') or {}).get('user', {}).get('displayName')}")
        print(f"  body: {body[:500]}")
        print()

    show("Latest SENT (you → them):    ", latest_sent)
    show("Latest RECEIVED (them → you):", latest_received)


if __name__ == "__main__":
    main()
