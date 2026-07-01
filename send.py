#!/usr/bin/env python3
import sys
import json
import requests
import graph

def main():
    if len(sys.argv) < 3:
        raise SystemExit("usage: send.py <chat-id> <message>")
    chat_id, text = sys.argv[1], sys.argv[2]
    token, _ = graph.get_token()
    r = requests.post(
        f"{graph.GRAPH}/chats/{chat_id}/messages",
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json={"body": {"contentType": "text", "content": text}},
    )
    print(r.status_code)
    print(json.dumps(r.json(), indent=2))
    r.raise_for_status()

if __name__ == "__main__":
    main()
