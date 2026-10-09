"""Token-free client for the same-computer movement probe."""
import argparse
import json
from pathlib import Path
import urllib.error
import urllib.request
import urllib.parse
import uuid


def request(url, path, payload=None):
    parsed = urllib.parse.urlsplit(url)
    if parsed.scheme != "http" or parsed.hostname not in ("127.0.0.1", "localhost") or parsed.username or parsed.password or parsed.path not in ("", "/") or parsed.query or parsed.fragment:
        raise ValueError("Use http://127.0.0.1:PORT or http://localhost:PORT on the controller computer.")
    data = None if payload is None else json.dumps(payload).encode()
    req = urllib.request.Request(url.rstrip("/") + path, data=data,
                                 headers={"Content-Type": "application/json"})
    # Keep local control off proxies and reject redirects to other hosts.
    class NoRedirect(urllib.request.HTTPRedirectHandler):
        def redirect_request(self, *args, **kwargs):
            return None
    opener = urllib.request.build_opener(urllib.request.ProxyHandler({}), NoRedirect())
    with opener.open(req, timeout=2) as response:
        return json.load(response)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("action", choices=("status", "events", "stop", "move"))
    parser.add_argument("--url", default="http://127.0.0.1:8765")
    parser.add_argument("--json", help="Movement JSON, preferably including a unique command_id")
    parser.add_argument("--json-file", type=Path, help="Alternative to --json; avoids Windows shell quote handling")
    args = parser.parse_args()
    if args.action == "move":
        if bool(args.json) == bool(args.json_file):
            parser.error("move requires exactly one of --json or --json-file")
        payload = json.loads(args.json_file.read_text(encoding="utf-8") if args.json_file else args.json)
        if not isinstance(payload, dict):
            parser.error("movement must be a JSON object")
        payload.setdefault("command_id", str(uuid.uuid4()))
        print(json.dumps(request(args.url, "/move", payload), indent=2))
    else:
        payload = {} if args.action == "stop" else None
        print(json.dumps(request(args.url, "/" + args.action, payload), indent=2))


if __name__ == "__main__":
    main()
