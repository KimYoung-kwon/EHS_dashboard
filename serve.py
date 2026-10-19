"""AX·FA 과제 현황판을 사내망에 띄우는 서버 (Python 3.8 이상, 표준 라이브러리만 사용).

실행: python serve.py                  → 이 PC의 모든 IP, 포트 8765
      python serve.py --port 9000      → 포트 변경
      python serve.py --readonly       → 보기만 가능(입력·수정·업로드 막기)

접속자 누구나 보고 입력·수정·엑셀 반영을 할 수 있습니다. 데이터는 data 폴더에 저장됩니다.
  data/store.json      현재 과제 데이터
  data/history/*.json  주간 스냅샷(주마다 마지막 저장 상태) → 화면의 '전주 대비'
  data/changes.jsonl   변경 이력(누가 언제 무엇을)
  data/backup/*.json   저장 직전 상태(최근 200개). 잘못 지웠을 때 되살리는 용도
"""
import argparse
import json
import math
import os
import posixpath
import re
import socket
import sys
import threading
import uuid
from datetime import date, datetime, timedelta
from functools import partial
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import unquote, urlsplit

ROOT = Path(__file__).resolve().parent
DATA = ROOT / "data"
STORE = DATA / "store.json"
HISTORY = DATA / "history"
BACKUP = DATA / "backup"
CHANGES = DATA / "changes.jsonl"
KEEP_BACKUPS = 200
READONLY = False
# 이 경로와 index.html만 내려줍니다. data, .git, tools 등은 404로 막습니다.
PUBLIC_DIRS = ("/assets/", "/vendor/", "/sample/")
# 주간 스냅샷에 남기는 항목: 전주 대비와 8주 추이 계산에 쓰는 값만
SLIM = ("id", "team", "group", "task", "kind", "priority", "stage", "found_fte", "secured_fte",
        "s3_date", "s3_fte", "s4_date", "s4_fte", "s5_date", "s5_fte", "re_team", "re_line", "re_etc", "act_end")
KEY_RE = re.compile(r"^[a-z][a-z0-9_]{0,39}$")
ID_RE = re.compile(r"^[A-Za-z0-9_-]{1,40}$")
MAX_BODY = 8 * 1024 * 1024
LOCK = threading.Lock()


class ApiError(Exception):
    def __init__(self, status, message):
        super().__init__(message)
        self.status = status


# ── 파일 읽기·쓰기 ───────────────────────────────────────────
def read_json(path, default):
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except FileNotFoundError:
        return default


def write_json(path, obj):
    # 임시 파일에 쓴 뒤 바꿔치기해서, 저장 도중 꺼져도 파일이 반쯤 쓰인 상태로 남지 않게 함
    tmp = path.with_name(path.name + ".tmp")
    tmp.write_text(json.dumps(obj, ensure_ascii=False), encoding="utf-8")
    os.replace(tmp, path)


def load_store():
    return read_json(STORE, {"version": 0, "tasks": [], "updated": None, "updated_by": None})


def week_of(d):
    y, w, _ = d.isocalendar()
    monday = d - timedelta(days=d.weekday())
    return f"{y}-W{w:02d}", monday, monday + timedelta(days=6)


# ── 입력값 검사 ─────────────────────────────────────────────
def clean_task(t):
    if not isinstance(t, dict):
        raise ApiError(400, "과제 데이터 형식이 올바르지 않습니다.")
    out = {}
    for k, v in t.items():
        if k == "id" or not KEY_RE.match(str(k)):
            continue
        if v is None or isinstance(v, bool):
            out[k] = v
        elif isinstance(v, (int, float)):
            out[k] = v if math.isfinite(v) else None
        elif isinstance(v, str):
            v = v.strip()[:4000]
            out[k] = v or None
        else:
            raise ApiError(400, f"'{k}' 값의 형식이 올바르지 않습니다.")
    if not out.get("team") or not out.get("task"):
        raise ApiError(400, "팀과 과제명(중분류)은 꼭 입력해야 합니다.")
    tid = t.get("id")
    out["id"] = tid if isinstance(tid, str) and ID_RE.match(tid) else uuid.uuid4().hex[:12]
    return out


def check_unique(tasks):
    seen = set()
    for t in tasks:
        key = (t["team"], t["task"])
        if key in seen:
            raise ApiError(400, f"{t['team']}에 같은 과제명 '{t['task']}'이(가) 두 번 있습니다.")
        seen.add(key)


# ── 저장 ────────────────────────────────────────────────────
def backup(st):
    # 저장 직전 상태를 남겨 두고 오래된 것부터 지움
    if not st["tasks"]:
        return
    write_json(BACKUP / f"store-{datetime.now().strftime('%Y%m%d-%H%M%S-%f')}.json", st)
    for old in sorted(BACKUP.glob("store-*.json"))[:-KEEP_BACKUPS]:
        old.unlink(missing_ok=True)


def commit(body, ip):
    if READONLY:
        raise ApiError(403, "보기 전용으로 실행 중인 서버라 저장할 수 없습니다.")
    editor = str(body.get("editor") or "").strip()[:30] or f"이름 없음({ip})"
    op = body.get("op")
    with LOCK:
        st = load_store()
        if body.get("base_version") != st["version"]:
            raise ApiError(409, "다른 사람이 먼저 수정했습니다.")
        tasks = st["tasks"]
        log = {"op": op}
        if op == "save":
            t = clean_task(body.get("task"))
            old = next((x for x in tasks if x["id"] == t["id"]), None)
            if old:
                tasks = [t if x["id"] == t["id"] else x for x in tasks]
                log["fields"] = sorted(k for k in set(old) | set(t) if old.get(k) != t.get(k) and k != "id")
            else:
                tasks = tasks + [t]
                log["op"] = "add"
            log.update(team=t["team"], task=t["task"])
        elif op == "delete":
            old = next((x for x in tasks if x["id"] == body.get("id")), None)
            if not old:
                raise ApiError(404, "삭제할 과제를 찾지 못했습니다. 이미 삭제됐을 수 있습니다.")
            tasks = [x for x in tasks if x["id"] != old["id"]]
            log.update(team=old["team"], task=old["task"])
        elif op == "import":
            teams = body.get("teams")
            if not isinstance(teams, list) or not teams or not all(isinstance(x, str) and x for x in teams):
                raise ApiError(400, "반영할 팀 목록이 없습니다.")
            new = [clean_task(x) for x in (body.get("tasks") or [])]
            if any(x["team"] not in teams for x in new):
                raise ApiError(400, "반영할 팀 목록에 없는 팀의 과제가 섞여 있습니다.")
            summary = []
            for team in teams:
                before = {x["task"]: x for x in tasks if x["team"] == team}
                after = {x["task"]: x for x in new if x["team"] == team}
                summary.append({"team": team, "added": len(after.keys() - before.keys()), "removed": len(before.keys() - after.keys()),
                                "changed": sum(1 for k in after.keys() & before.keys() if {**after[k], "id": None} != {**before[k], "id": None})})
            tasks = [x for x in tasks if x["team"] not in teams] + new
            log["teams"] = summary
        else:
            raise ApiError(400, "알 수 없는 작업입니다.")
        check_unique(tasks)
        backup(st)
        now = datetime.now()
        st = {"version": st["version"] + 1, "tasks": tasks, "updated": now.strftime("%Y-%m-%d %H:%M"), "updated_by": editor}
        write_json(STORE, st)
        week, _, end = week_of(now.date())
        write_json(HISTORY / f"{week}.json", {"week": week, "end": end.isoformat(), "saved": now.isoformat(timespec="seconds"),
                                               "tasks": [{k: x.get(k) for k in SLIM} for x in tasks]})
        with CHANGES.open("a", encoding="utf-8") as f:
            f.write(json.dumps({"at": now.strftime("%Y-%m-%d %H:%M"), "editor": editor, "ip": ip, **log}, ensure_ascii=False) + "\n")
        return st


def store_payload():
    st = load_store()
    _, monday, _ = week_of(date.today())
    return {**st, "editable": not READONLY, "week_start": monday.isoformat()}


def history_payload():
    files = sorted(HISTORY.glob("*-W*.json"))[-16:]
    return {"weeks": [read_json(p, {}) for p in files]}


def changes_payload():
    try:
        lines = CHANGES.read_text(encoding="utf-8").splitlines()[-100:]
    except FileNotFoundError:
        lines = []
    out = []
    for line in reversed(lines):
        try:
            x = json.loads(line)
            x.pop("ip", None)
            out.append(x)
        except ValueError:
            continue
    return {"changes": out}


# ── HTTP ────────────────────────────────────────────────────
class Handler(SimpleHTTPRequestHandler):
    extensions_map = {
        **SimpleHTTPRequestHandler.extensions_map,
        ".html": "text/html; charset=utf-8", ".js": "text/javascript; charset=utf-8", ".css": "text/css; charset=utf-8",
        ".json": "application/json; charset=utf-8", ".woff2": "font/woff2",
        ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    }

    def path_only(self):
        # '..'를 먼저 정리한 경로로 검사해야 /vendor/../data 같은 우회가 막힘
        return posixpath.normpath(unquote(urlsplit(self.path).path))

    def send_json(self, status, obj):
        body = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        path = self.path_only()
        api = {"/api/store": store_payload, "/api/history": history_payload, "/api/changes": changes_payload,
               "/api/version": lambda: {"version": load_store()["version"]}}.get(path)
        if api:
            return self.send_json(200, api())
        if path in ("/", "/index.html") or path.startswith(PUBLIC_DIRS):
            return super().do_GET()
        self.send_error(404)

    def do_HEAD(self):
        path = self.path_only()
        if path in ("/", "/index.html") or path.startswith(PUBLIC_DIRS):
            return super().do_HEAD()
        self.send_error(404)

    def do_POST(self):
        path = self.path_only()
        ip = self.client_address[0]
        try:
            n = int(self.headers.get("Content-Length") or 0)
            if n > MAX_BODY:
                raise ApiError(413, "보낸 데이터가 너무 큽니다(8MB 초과).")
            try:
                body = json.loads(self.rfile.read(n) or b"{}")
            except ValueError:
                raise ApiError(400, "요청 형식이 올바르지 않습니다.")
            if not isinstance(body, dict):
                raise ApiError(400, "요청 형식이 올바르지 않습니다.")
            if path == "/api/commit":
                st = commit(body, ip)
                return self.send_json(200, {**st, "editable": not READONLY, "week_start": week_of(date.today())[1].isoformat()})
            raise ApiError(404, "없는 주소입니다.")
        except ApiError as e:
            self.send_json(e.status, {"error": str(e)})

    def list_directory(self, path):
        self.send_error(404)
        return None

    def end_headers(self):
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
    ap.add_argument("--readonly", action="store_true", help="보기만 가능하게 실행")
    a = ap.parse_args()
    global READONLY
    READONLY = a.readonly
    for d in (DATA, HISTORY, BACKUP):
        d.mkdir(exist_ok=True)
    try:
        httpd = Server((a.host, a.port), partial(Handler, directory=str(ROOT)))
    except OSError as e:
        print(f"[오류] {a.host}:{a.port}에서 서버를 열지 못했습니다: {e}")
        print(f"       다른 프로그램이 포트 {a.port}를 쓰고 있으면 'py serve.py --port 8766'처럼 다른 번호로 실행하세요.")
        sys.exit(1)
    hosts = lan_ips() if a.host == "0.0.0.0" else [a.host]
    st = load_store()
    print("AX·FA 과제 현황판 서버를 시작했습니다. 이 창을 닫거나 Ctrl+C를 누르면 멈춥니다.")
    print(f"  이 PC에서:       http://localhost:{a.port}")
    for ip in hosts:
        print(f"  사내 다른 PC에서: http://{ip}:{a.port}")
    print(f"  저장된 과제:      {len(st['tasks'])}건" + ("" if st["tasks"] else " (예시 데이터 표시 중, 엑셀 업로드나 과제 추가로 시작)"))
    print(f"  편집:            {'보기 전용 (--readonly)' if READONLY else '누구나 가능 · 저장 직전 상태는 data/backup에 보관'}")
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\n서버를 멈췄습니다.")
    finally:
        httpd.server_close()


if __name__ == "__main__":
    main()
