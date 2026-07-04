from __future__ import annotations

import json
import socket
import sys
import threading
import webbrowser
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse

from ojt_exam_maker_console import (
    DEFAULT_COUNTS,
    DEFAULT_SCORES,
    QUESTION_TYPES,
    calculate_total,
    find_default_workbook,
    find_output_desktop,
    load_question_banks,
    safe_filename,
    select_questions,
    write_exam_workbook,
)


APP_TITLE = "OJT EXAM MAKER"


HTML = r"""<!doctype html>
<html lang="ko">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>OJT EXAM MAKER</title>
  <style>
    :root {
      --bg: #0b1220;
      --panel: #121c2e;
      --panel2: #18243a;
      --line: #2a3b57;
      --text: #e8eef8;
      --muted: #92a1b7;
      --accent: #20d084;
      --warn: #f6b44b;
      --danger: #f06464;
      --blue: #4aa3ff;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      background: radial-gradient(circle at 20% 0%, #16233a 0, #0b1220 34%, #070b13 100%);
      color: var(--text);
      font-family: "Malgun Gothic", "Segoe UI", Arial, sans-serif;
      height: 100vh;
      overflow: hidden;
    }
    .app { height: 100vh; display: flex; flex-direction: column; padding: 20px 24px 22px; gap: 14px; }
    .topbar { display: flex; align-items: center; gap: 16px; }
    .brand { font-size: 28px; font-weight: 800; letter-spacing: .5px; }
    .sub { color: var(--muted); font-size: 13px; margin-top: 8px; }
    .badge { margin-left: auto; padding: 8px 14px; border-radius: 6px; background: #102f22; color: #83ffc0; font-weight: 800; }
    .filebar {
      display: grid; grid-template-columns: 92px 1fr 118px; gap: 10px; align-items: center;
      padding: 12px; background: rgba(18,28,46,.96); border: 1px solid var(--line); border-radius: 8px;
    }
    .label { color: var(--muted); font-size: 12px; font-weight: 800; }
    input {
      width: 100%; border: 1px solid #253650; background: #08111f; color: var(--text);
      border-radius: 6px; padding: 10px 12px; outline: none; font-size: 14px;
    }
    input:focus { border-color: var(--blue); }
    button {
      border: 0; border-radius: 6px; padding: 10px 14px; cursor: pointer;
      font-weight: 800; font-family: inherit; background: var(--panel2); color: var(--text);
    }
    button.primary { background: var(--accent); color: #04130d; font-size: 16px; padding: 14px; }
    button:hover { filter: brightness(1.08); }
    .grid { flex: 1; min-height: 0; display: grid; grid-template-columns: 1.35fr .9fr .9fr; gap: 14px; }
    .card {
      min-height: 0; background: rgba(18,28,46,.96); border: 1px solid var(--line); border-radius: 8px;
      display: flex; flex-direction: column; overflow: hidden;
    }
    .cardHead { padding: 16px 16px 10px; border-bottom: 1px solid rgba(42,59,87,.7); }
    .title { font-size: 18px; font-weight: 800; }
    .hint { color: var(--muted); font-size: 12px; margin-top: 4px; }
    .tableWrap { overflow: auto; padding: 10px; }
    table { width: 100%; border-collapse: collapse; font-size: 13px; }
    th { position: sticky; top: 0; background: #1e2d48; color: #dbe8ff; z-index: 1; }
    th, td { padding: 9px 8px; border-bottom: 1px solid #243651; text-align: center; }
    td:first-child, th:first-child { text-align: left; }
    tr { cursor: pointer; }
    tr:hover { background: #162a43; }
    tr.selected { background: #14513f; }
    .body { padding: 16px; display: flex; flex-direction: column; gap: 14px; }
    .settingGrid { display: grid; grid-template-columns: 1fr 86px 86px; gap: 10px; align-items: center; }
    .typeName { font-size: 16px; font-weight: 800; }
    .target { display: grid; grid-template-columns: 1fr 96px; gap: 10px; align-items: center; padding-top: 6px; }
    .scoreCard { background: #10243a; border: 1px solid var(--line); border-radius: 8px; padding: 16px; }
    .scoreLabel { color: var(--muted); font-size: 12px; font-weight: 900; }
    .score { font-size: 56px; line-height: 1; font-weight: 900; color: var(--accent); margin-top: 6px; }
    .selectedBox { background: #0a1322; border: 1px solid #263954; border-radius: 8px; padding: 14px; line-height: 1.7; }
    .log {
      flex: 1; min-height: 0; overflow: auto; background: #07101d; border: 1px solid #263954; border-radius: 8px;
      padding: 12px; color: #bcd1ff; font-family: Consolas, monospace; font-size: 12px; white-space: pre-wrap;
    }
    .ok { color: var(--accent); }
    .bad { color: var(--danger); }
    .warn { color: var(--warn); }
  </style>
</head>
<body>
  <div class="app">
    <div class="topbar">
      <div>
        <div class="brand">OJT EXAM MAKER</div>
        <div class="sub">원본 시험지 양식 유지 · 랜덤 문항 추출 · 100점 검증</div>
      </div>
      <div id="badge" class="badge">READY</div>
    </div>
    <div class="filebar">
      <div class="label">문제은행</div>
      <input id="workbookPath" />
      <button onclick="reloadBanks()">새로고침</button>
    </div>
    <div class="grid">
      <section class="card">
        <div class="cardHead">
          <div class="title">공정 선택</div>
          <div class="hint">문제은행 시트별 보유 문항 현황</div>
        </div>
        <div class="tableWrap">
          <table>
            <thead><tr><th>공정명</th><th>공통</th><th>객관식</th><th>주관식</th></tr></thead>
            <tbody id="bankRows"></tbody>
          </table>
        </div>
      </section>
      <section class="card">
        <div class="cardHead">
          <div class="title">시험 조건</div>
          <div class="hint">총점이 목표 점수와 같을 때만 생성</div>
        </div>
        <div class="body">
          <div class="settingGrid">
            <div class="label">유형</div><div class="label">문항</div><div class="label">점수</div>
            <div class="typeName">공통</div><input id="count_common" value="2" oninput="updateTotal()"><input id="score_common" value="2.5" oninput="updateTotal()">
            <div class="typeName">객관식</div><input id="count_choice" value="20" oninput="updateTotal()"><input id="score_choice" value="4" oninput="updateTotal()">
            <div class="typeName">주관식</div><input id="count_subjective" value="3" oninput="updateTotal()"><input id="score_subjective" value="5" oninput="updateTotal()">
          </div>
          <div class="target"><div class="label">목표 점수</div><input id="targetScore" value="100" oninput="updateTotal()"></div>
          <div>
            <div class="label" style="margin-bottom:7px;">저장 위치</div>
            <input id="outputDir" />
          </div>
          <button class="primary" onclick="generateExam()">시험지 생성</button>
        </div>
      </section>
      <section class="card">
        <div class="cardHead">
          <div class="title">상태</div>
          <div class="hint">선택 공정 및 점수 검증</div>
        </div>
        <div class="body" style="height:100%;">
          <div class="scoreCard">
            <div class="scoreLabel">TOTAL SCORE</div>
            <div id="totalScore" class="score">100</div>
          </div>
          <div id="selectedBox" class="selectedBox">공정을 선택하세요.</div>
          <div id="log" class="log">대기 중</div>
        </div>
      </section>
    </div>
  </div>
  <script>
    let banks = [];
    let selected = 0;
    const ids = {
      "공통": ["count_common", "score_common"],
      "객관식": ["count_choice", "score_choice"],
      "주관식": ["count_subjective", "score_subjective"]
    };
    function log(msg) {
      const box = document.getElementById("log");
      box.textContent += "\n" + msg;
      box.scrollTop = box.scrollHeight;
    }
    function badge(text, mode="ready") {
      const b = document.getElementById("badge");
      b.textContent = text;
      b.style.background = mode === "error" ? "#3b1111" : mode === "busy" ? "#0d2a45" : "#102f22";
      b.style.color = mode === "error" ? "#ffb0b0" : mode === "busy" ? "#a7d7ff" : "#83ffc0";
    }
    function num(id) { return Number(document.getElementById(id).value || 0); }
    function currentPayload() {
      return {
        workbook_path: document.getElementById("workbookPath").value,
        output_dir: document.getElementById("outputDir").value,
        selected_index: selected,
        counts: {"공통": num("count_common"), "객관식": num("count_choice"), "주관식": num("count_subjective")},
        scores: {"공통": num("score_common"), "객관식": num("score_choice"), "주관식": num("score_subjective")},
        target_score: num("targetScore")
      };
    }
    function updateTotal() {
      const p = currentPayload();
      const total = p.counts["공통"] * p.scores["공통"] + p.counts["객관식"] * p.scores["객관식"] + p.counts["주관식"] * p.scores["주관식"];
      const el = document.getElementById("totalScore");
      el.textContent = Number.isInteger(total) ? total : total.toFixed(1);
      el.style.color = Math.abs(total - p.target_score) < 0.0001 ? "var(--accent)" : "var(--warn)";
    }
    function renderBanks() {
      const body = document.getElementById("bankRows");
      body.innerHTML = "";
      banks.forEach((bank, i) => {
        const tr = document.createElement("tr");
        if (i === selected) tr.className = "selected";
        tr.innerHTML = `<td>${bank.name}</td><td>${bank.counts["공통"]}</td><td>${bank.counts["객관식"]}</td><td>${bank.counts["주관식"]}</td>`;
        tr.onclick = () => { selected = i; renderBanks(); renderSelected(); };
        body.appendChild(tr);
      });
      renderSelected();
    }
    function renderSelected() {
      const box = document.getElementById("selectedBox");
      if (!banks.length) { box.textContent = "문제은행을 불러오세요."; return; }
      const b = banks[selected];
      box.innerHTML = `<b>${b.name}</b><br>보유 문항: 공통 ${b.counts["공통"]} / 객관식 ${b.counts["객관식"]} / 주관식 ${b.counts["주관식"]}`;
    }
    async function reloadBanks() {
      badge("LOADING", "busy");
      const path = document.getElementById("workbookPath").value;
      const res = await fetch("/api/banks", {method:"POST", headers:{"Content-Type":"application/json"}, body:JSON.stringify({workbook_path:path})});
      const data = await res.json();
      if (!data.ok) { badge("ERROR", "error"); log("오류: " + data.error); alert(data.error); return; }
      banks = data.banks;
      selected = 0;
      document.getElementById("workbookPath").value = data.workbook_path;
      document.getElementById("outputDir").value = data.output_dir;
      renderBanks();
      updateTotal();
      badge("READY");
      log(`문제은행 로드 완료: ${banks.length}개 공정`);
    }
    async function generateExam() {
      badge("RUNNING", "busy");
      const res = await fetch("/api/generate", {method:"POST", headers:{"Content-Type":"application/json"}, body:JSON.stringify(currentPayload())});
      const data = await res.json();
      if (!data.ok) { badge("ERROR", "error"); log("생성 실패: " + data.error); alert(data.error); return; }
      badge("DONE");
      log("생성 완료: " + data.output_path);
      alert("시험지 생성 완료\n\n" + data.output_path);
    }
    reloadBanks();
  </script>
</body>
</html>
"""


class AppState:
    def __init__(self):
        self.default_workbook = find_default_workbook()
        self.banks = []
        self.workbook_path = self.default_workbook


STATE = AppState()


def json_response(handler: BaseHTTPRequestHandler, payload: dict, status: int = 200):
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    handler.send_response(status)
    handler.send_header("Content-Type", "application/json; charset=utf-8")
    handler.send_header("Content-Length", str(len(body)))
    handler.end_headers()
    handler.wfile.write(body)


def read_json(handler: BaseHTTPRequestHandler) -> dict:
    length = int(handler.headers.get("Content-Length", "0"))
    if not length:
        return {}
    return json.loads(handler.rfile.read(length).decode("utf-8"))


def banks_payload(workbook_path: Path):
    banks = load_question_banks(workbook_path)
    STATE.banks = banks
    STATE.workbook_path = workbook_path
    output_dir = find_output_desktop(workbook_path) / "OJT_Random_Exam_Output"
    return {
        "ok": True,
        "workbook_path": str(workbook_path),
        "output_dir": str(output_dir),
        "banks": [{"name": b.display_name, "counts": b.counts} for b in banks],
    }


class Handler(BaseHTTPRequestHandler):
    def log_message(self, _format, *args):
        return

    def do_GET(self):
        parsed = urlparse(self.path)
        if parsed.path == "/":
            body = HTML.encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            return
        self.send_error(404)

    def do_POST(self):
        try:
            if self.path == "/api/banks":
                data = read_json(self)
                path_text = data.get("workbook_path") or str(STATE.default_workbook or "")
                workbook_path = Path(path_text)
                if not workbook_path.exists():
                    raise FileNotFoundError(f"문제은행 파일을 찾을 수 없습니다: {workbook_path}")
                json_response(self, banks_payload(workbook_path))
                return

            if self.path == "/api/generate":
                data = read_json(self)
                workbook_path = Path(data["workbook_path"])
                if not workbook_path.exists():
                    raise FileNotFoundError(f"문제은행 파일을 찾을 수 없습니다: {workbook_path}")
                if not STATE.banks or STATE.workbook_path != workbook_path:
                    STATE.banks = load_question_banks(workbook_path)
                    STATE.workbook_path = workbook_path

                selected_index = int(data["selected_index"])
                bank = STATE.banks[selected_index]
                counts = {kind: int(data["counts"][kind]) for kind in QUESTION_TYPES}
                scores = {kind: float(data["scores"][kind]) for kind in QUESTION_TYPES}
                target = float(data["target_score"])
                total = calculate_total(counts, scores)
                if abs(total - target) >= 0.0001:
                    raise ValueError(f"총점이 목표 점수와 다릅니다. 현재 {total:g}점 / 목표 {target:g}점")
                questions = select_questions(bank, counts, scores)
                output_dir = Path(data["output_dir"])
                output_dir.mkdir(parents=True, exist_ok=True)
                output_path = output_dir / f"OJT_시험지_{safe_filename(bank.display_name)}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsm"
                write_exam_workbook(output_path, bank.display_name, questions, counts, scores, target, template_path=workbook_path)
                json_response(self, {"ok": True, "output_path": str(output_path)})
                return

            self.send_error(404)
        except Exception as exc:
            json_response(self, {"ok": False, "error": str(exc)}, status=200)


def free_port() -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.bind(("127.0.0.1", 0))
        return sock.getsockname()[1]


def self_test():
    if STATE.default_workbook:
        payload = banks_payload(STATE.default_workbook)
        print(f"OK: {len(payload['banks'])} banks")
    else:
        print("OK: no default workbook")


def main():
    if "--self-test" in sys.argv:
        self_test()
        return
    port = free_port()
    server = ThreadingHTTPServer(("127.0.0.1", port), Handler)
    url = f"http://127.0.0.1:{port}/"
    threading.Timer(0.5, lambda: webbrowser.open(url)).start()
    server.serve_forever()


if __name__ == "__main__":
    main()
