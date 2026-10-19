"""AX·FA 과제 현황판을 사내망에 띄우는 서버 (Python 3.8 이상, 표준 라이브러리만 사용).

실행: python serve.py               → 이 PC의 모든 IP, 포트 8765
      python serve.py --port 9000   → 포트 변경
      python serve.py --host 12.82.113.36 → 특정 IP에서만 받기

data 폴더에 넣은 엑셀 중 가장 최근에 수정된 파일을 접속자 모두가 봅니다.
서버는 읽기 전용입니다. 화면의 업로드는 올린 사람의 브라우저에서만 보입니다.
"""
import argparse
import json
import os
import posixpath
import socket
import sys
from datetime import datetime
from functools import partial
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import unquote, urlsplit

ROOT = Path(__file__).resolve().parent
DATA = ROOT / "data"
DATA_EXTS = {".xlsx", ".xls", ".csv"}
# 이 경로와 index.html만 내려줍니다. .git, tools 등 나머지는 404로 막습니다.
PUBLIC_DIRS = ("/vendor/", "/data/", "/sample/")


def latest_data_file():
    # ~$로 시작하는 파일은 엑셀이 파일을 여는 동안 만드는 잠금 파일이라 제외
    files = [p for p in DATA.glob("*") if p.is_file() and p.suffix.lower() in DATA_EXTS and not p.name.startswith(("~$", "."))]
    return max(files, key=lambda p: p.stat().st_mtime, default=None)


class Handler(SimpleHTTPRequestHandler):
    extensions_map = {
        **SimpleHTTPRequestHandler.extensions_map,
        ".html": "text/html; charset=utf-8",
        ".js": "text/javascript; charset=utf-8",
        ".json": "application/json; charset=utf-8",
        ".csv": "text/csv; charset=utf-8",
        ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        ".xls": "application/vnd.ms-excel",
    }

    def allowed(self, path):
        return path in ("/", "/index.html") or path.startswith(PUBLIC_DIRS)

    def route(self):
        # '..'를 먼저 정리한 경로로 검사해야 /vendor/../.git 같은 우회가 막힘
        path = posixpath.normpath(unquote(urlsplit(self.path).path))
        if path == "/api/data":
            f = latest_data_file()
            info = {"file": f.name, "modified": datetime.fromtimestamp(f.stat().st_mtime).strftime("%Y-%m-%d %H:%M"), "size": f.stat().st_size, "mtime": f.stat().st_mtime} if f else {"file": None}
            body = json.dumps(info, ensure_ascii=False).encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            return body
        if not self.allowed(path):
            self.send_error(404)
            return None
        return False

    def do_GET(self):
        r = self.route()
        if r is False:
            return super().do_GET()
        if r:
            self.wfile.write(r)

    def do_HEAD(self):
        r = self.route()
        if r is False:
            return super().do_HEAD()

    def list_directory(self, path):
        self.send_error(404)
        return None

    def end_headers(self):
        # 데이터 파일을 바꾸면 접속자가 새로고침 없이도 1분 안에 새 파일을 받도록 캐시하지 않음
        self.send_header("Cache-Control", "no-cache")
        super().end_headers()


class Server(ThreadingHTTPServer):
    # Windows에서 SO_REUSEADDR를 켜면 이미 사용 중인 포트에도 조용히 붙어 버리므로 끔
    allow_reuse_address = os.name != "nt"
    daemon_threads = True


def lan_ips():
    ips = set()
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as s:
            s.connect(("10.255.255.255", 1))  # 실제로 패킷을 보내지 않고 기본 경로의 IP만 확인
            ips.add(s.getsockname()[0])
    except OSError:
        pass
    try:
        ips.update(i[4][0] for i in socket.getaddrinfo(socket.gethostname(), None, socket.AF_INET))
    except OSError:
        pass
    return sorted(ip for ip in ips if not ip.startswith("127."))


def main():
    ap = argparse.ArgumentParser(description="AX·FA 과제 현황판 서버")
    ap.add_argument("--host", default="0.0.0.0", help="받을 IP (기본: 이 PC의 모든 IP)")
    ap.add_argument("--port", type=int, default=int(os.environ.get("DASHBOARD_PORT", 8765)), help="포트 (기본: 8765)")
    a = ap.parse_args()
    DATA.mkdir(exist_ok=True)
    try:
        httpd = Server((a.host, a.port), partial(Handler, directory=str(ROOT)))
    except OSError as e:
        print(f"[오류] {a.host}:{a.port}에서 서버를 열지 못했습니다: {e}")
        print(f"       다른 프로그램이 포트 {a.port}를 쓰고 있으면 'python serve.py --port 8766'처럼 다른 번호로 실행하세요.")
        sys.exit(1)
    hosts = lan_ips() if a.host == "0.0.0.0" else [a.host]
    f = latest_data_file()
    print("AX·FA 과제 현황판 서버를 시작했습니다. 이 창을 닫거나 Ctrl+C를 누르면 멈춥니다.")
    print(f"  이 PC에서:       http://localhost:{a.port}")
    for ip in hosts:
        print(f"  사내 다른 PC에서: http://{ip}:{a.port}")
    print(f"  공유 데이터:      {f.name if f else '없음 (data 폴더에 엑셀을 넣으면 모두에게 표시, 그 전까지는 예시 데이터)'}")
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\n서버를 멈췄습니다.")
    finally:
        httpd.server_close()


if __name__ == "__main__":
    main()
