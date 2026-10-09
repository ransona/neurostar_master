"""Standard-library client: supervise requests with an explicit heartbeat."""
import argparse
import json
import os
from pathlib import Path
import time
import urllib.error
import urllib.request
import uuid


def request(url, token, path, payload=None):
    data = None if payload is None else json.dumps(payload).encode()
    req = urllib.request.Request(url.rstrip("/") + path, data=data,
                                 headers={"Authorization": "Bearer " + token,
                                          "Content-Type": "application/json"})
    # Do not leak the credential through environment proxy settings/redirects.
    class NoRedirect(urllib.request.HTTPRedirectHandler):
        def redirect_request(self, *args, **kwargs):
            return None
    opener = urllib.request.build_opener(urllib.request.ProxyHandler({}), NoRedirect())
    with opener.open(req, timeout=2) as response:
        return json.load(response)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("action", choices=("status", "events", "heartbeat", "stop", "move"))
    parser.add_argument("--url", default="http://127.0.0.1:8765")
    parser.add_argument("--json", help="Movement JSON, preferably including a unique command_id")
    parser.add_argument("--json-file", type=Path, help="Alternative to --json; avoids Windows shell quote handling")
    parser.add_argument("--seconds", type=float, default=60, help="Heartbeat session duration, 1–300 seconds")
    args = parser.parse_args()
    token = os.environ.get("STEREODRIVE_PROBE_TOKEN")
    if not token:
        parser.error("Set STEREODRIVE_PROBE_TOKEN in this process environment.")
    if args.action == "heartbeat":
        if not 1 <= args.seconds <= 300:
            parser.error("heartbeat duration must be 1–300 seconds")
        print("Sending heartbeat every 0.5 seconds. Ctrl+C stops it and requests Stop.")
        deadline = time.monotonic() + args.seconds
        try:
            while time.monotonic() < deadline:
                try:
                    request(args.url, token, "/heartbeat", {})
                except urllib.error.HTTPError as exc:
                    if exc.code != 400:
                        raise
                    # Human may not yet have armed; heartbeat cannot arm remotely.
                time.sleep(0.5)
        finally:
            request(args.url, token, "/stop", {})
    elif args.action == "move":
        if bool(args.json) == bool(args.json_file):
            parser.error("move requires exactly one of --json or --json-file; run heartbeat separately before local ARM.")
        payload = json.loads(args.json_file.read_text(encoding="utf-8") if args.json_file else args.json)
        if not isinstance(payload, dict):
            parser.error("movement must be a JSON object")
        payload.setdefault("command_id", str(uuid.uuid4()))
        print(json.dumps(request(args.url, token, "/move", payload), indent=2))
    else:
        payload = {} if args.action == "stop" else None
        print(json.dumps(request(args.url, token, "/" + args.action, payload), indent=2))


if __name__ == "__main__":
    main()
