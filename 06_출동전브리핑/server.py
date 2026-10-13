"""브리핑 백엔드 (표준 라이브러리만 사용).

GET /api/health   → {"ok": true, ...}
GET /api/briefing → dist/briefing.json
그 외 경로        → dist/ 정적 파일
"""
import argparse
import json
from datetime import datetime
from functools import partial
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

DIST = Path(__file__).resolve().parent / "dist"


class Handler(SimpleHTTPRequestHandler):
    def end_headers(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        super().end_headers()

    def do_OPTIONS(self):
        self.send_response(204)
        self.end_headers()

    def _json(self, code, obj):
        body = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        path = self.path.split("?")[0]
        brief = DIST / "briefing.json"
        if path == "/api/health":
            meta = {}
            if brief.exists():
                meta = json.loads(brief.read_text(encoding="utf-8")).get("meta", {})
            return self._json(200, {"ok": True, "service": "ons-brief", "time": datetime.now().isoformat(timespec="seconds"),
                                    "baseTime": meta.get("baseTime"), "open": meta.get("counts", {}).get("open")})
        if path == "/api/briefing":
            if not brief.exists():
                return self._json(404, {"ok": False, "error": "briefing.json 없음 - prepare.py 먼저 실행"})
            return self._json(200, json.loads(brief.read_text(encoding="utf-8")))
        return super().do_GET()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--host", default="127.0.0.1")
    ap.add_argument("--port", type=int, default=8000)
    a = ap.parse_args()
    srv = ThreadingHTTPServer((a.host, a.port), partial(Handler, directory=str(DIST)))
    print(f"ons-brief backend: http://{a.host}:{a.port}  (health: /api/health)", flush=True)
    srv.serve_forever()


if __name__ == "__main__":
    main()
