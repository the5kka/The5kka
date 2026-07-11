from __future__ import annotations

from html import escape
from pathlib import Path


ROOT = Path(__file__).resolve().parent
OUT_DIR = ROOT / "JIIN_DNC_V4_design_gallery"
OUT_FILE = OUT_DIR / "JIIN_DNC_V4_Design_Gallery.html"


PALETTES = [
    ("MES Blue", "#1d4ed8", "#dbeafe", "#eff6ff", "#0f172a", "#60a5fa"),
    ("Simtek Teal", "#008c95", "#d7f6f3", "#f0fdfa", "#12323a", "#14b8a6"),
    ("KCC Burgundy", "#8b1232", "#f7e4eb", "#fff5f7", "#2a1018", "#d94670"),
    ("TLB Cyan", "#0369a1", "#e0f2fe", "#f0f9ff", "#143044", "#38bdf8"),
    ("Graphite Line", "#334155", "#e2e8f0", "#f8fafc", "#0f172a", "#94a3b8"),
    ("Safety Green", "#047857", "#d1fae5", "#f0fdf4", "#10281e", "#34d399"),
    ("Quality Amber", "#b45309", "#fef3c7", "#fffbeb", "#2f1f08", "#f59e0b"),
    ("Navy Steel", "#1e3a8a", "#dbe4f0", "#f5f7fb", "#111827", "#64748b"),
    ("Clean Mono", "#111827", "#e5e7eb", "#ffffff", "#111827", "#9ca3af"),
    ("Process Violet", "#6d28d9", "#ede9fe", "#faf5ff", "#21133d", "#a78bfa"),
]

MAIN_LAYOUTS = [
    ("Classic 5:5", "V3와 가장 가까운 5:5 LOT 배치, 작업자 혼선이 가장 적음", "low"),
    ("Status First", "조건 결과와 DNC 진행 상태를 LOT 입력 바로 아래에 크게 배치", "low"),
    ("Right Command Rail", "실행/초기화/작업일보 버튼을 오른쪽 세로 레일로 고정", "low"),
    ("Compact Inspection", "공통 입력과 LOT 입력 높이를 줄여 로그 영역을 넓힘", "medium"),
]

POPUP_LAYOUTS = [
    ("Two Lot Guard", "LOT 1+LOT 2에서 조건/LOT No 불일치가 바로 보이는 구조", "low"),
    ("Wizard Header", "상단에 입력-확인-실행 진행 단계를 표시하는 팝업", "medium"),
    ("Dense Verify", "신규 검증 입력을 촘촘하게 모아 현장 키보드 입력에 유리", "low"),
]

SETTING_LAYOUTS = [
    ("Folder Matrix", "공정별 폴더와 상태를 표처럼 비교", "low"),
    ("Grouped Admin", "폴더/작업일보/라이선스/백업을 그룹 카드로 분리", "medium"),
]

ALERT_LAYOUTS = [
    ("Operator Alert", "작업자용 큰 제목, 한 문장 원인, 단일 확인 버튼", "low"),
]


def html_attrs(style: str) -> str:
    return f' style="{escape(style, quote=True)}"'


def status_pill(text: str, tone: str) -> str:
    return f'<span class="pill {tone}">{escape(text)}</span>'


def lot_fields(include_extra: bool = True) -> str:
    rows = [
        ("STEP", "520", "차수", "3차"),
        ("관리번호", "SCM00040C00-000", "LOT No", "MH2650002-35"),
        ("매수", "", "조건", "519x618 / X -5.00"),
    ]
    if include_extra:
        rows.append(("추가가공", "Bond_2", "조건(조회)", "519x618_-5_5R_5up_One Zig_Bond_2"))
        rows.append(("지그(조회)", "[통합지그] 604", "", ""))
    html = []
    for left_label, left_value, right_label, right_value in rows:
        html.append('<div class="field-row">')
        html.append(f'<div class="label">{escape(left_label)}</div><div class="input">{escape(left_value)}</div>')
        if right_label:
            html.append(f'<div class="label">{escape(right_label)}</div><div class="input">{escape(right_value)}</div>')
        else:
            html.append('<div></div><div></div>')
        html.append("</div>")
    return "\n".join(html)


def render_main_mockup(variant: dict) -> str:
    layout = variant["layout"]
    right_rail = "Right Command Rail" in layout
    compact = "Compact" in layout
    status_first = "Status First" in layout
    body_class = "mock-main compact" if compact else "mock-main"
    side = (
        '<div class="side-rail"><button>DNC 실행</button><button>입력 초기화</button>'
        '<button>작업일보 반영</button><button>작업일보 열기</button></div>'
        if right_rail
        else ""
    )
    status = (
        '<div class="status-wide">'
        f'{status_pill("2LOT 조건 일치 확인", "ok")}{status_pill("DNC 진행 상태: 대기중", "wait")}'
        f'{status_pill("작업일보 반영: 대기중", "wait")}'
        "</div>"
    )
    return f"""
    <div class="{body_class}">
      <div class="topbar"><div class="brand">JIIN DNC Manager</div><div class="logo">SIMMTECH</div></div>
      <div class="tabs"><span>TLB</span><span>심텍 SPS</span><span class="active">심텍 HDI</span><span>KCC PKG</span><span>KCC HDI</span><span>설정</span></div>
      <div class="process-head"><b>심텍 HDI DNC</b><button>DNC 실행</button><button>입력 초기화</button></div>
      <div class="common-strip"><span>설비 호기: 트리밍 1호기</span><span>작업일자: 2026-07-07</span><span>조: A</span><span>근무: 주간</span><span>작업자: 오국진</span></div>
      {status if status_first else ""}
      <div class="main-body">
        <div class="lot-card"><h4>LOT 1 입력</h4>{lot_fields()}</div>
        <div class="lot-card"><h4>LOT 2 입력 (선택)</h4>{lot_fields()}</div>
        {side}
      </div>
      {"" if status_first else status}
      <div class="log">[11:01:17] 조건 조회 완료<br>[11:01:20] DNC 대기중</div>
    </div>
    """


def render_popup_mockup(variant: dict) -> str:
    layout = variant["layout"]
    guard = "Guard" in layout
    wizard = "Wizard" in layout
    dense = "Dense" in layout
    warning = (
        '<div class="guard-line danger">LOT 1 / LOT 2 조건(조회) 다름 - 동시 신규 검증 불가</div>'
        if guard
        else '<div class="guard-line">조건 마스터 등록: 확인 필요</div>'
    )
    steps = '<div class="steps"><span class="on">1 입력</span><span>2 조건 확인</span><span>3 신규 DNC</span></div>' if wizard else ""
    return f"""
    <div class="modal-preview {'dense' if dense else ''}">
      <div class="modal-title">심텍 HDI 신규 모델 검증 DNC</div>
      {steps}
      <div class="lot-switch"><span>LOT 1 입력</span><span>LOT 2 입력</span><span class="active">LOT 1 + LOT 2</span></div>
      <div class="popup-grid">
        <div class="lot-card"><h4>LOT 1 신규 입력</h4>{lot_fields()}</div>
        <div class="lot-card"><h4>LOT 2 신규 입력</h4>{lot_fields()}</div>
      </div>
      {warning}
      <div class="modal-actions"><button class="primary">신규 모델 DNC 실행</button><button>입력 초기화</button><button>닫기</button></div>
    </div>
    """


def render_settings_mockup(variant: dict) -> str:
    grouped = "Grouped" in variant["layout"]
    rows = ["TLB", "심텍 SPS", "심텍 HDI", "KCC PKG", "KCC HDI"]
    folder_rows = "".join(
        f'<div class="folder-row"><span>{escape(name)}</span><div class="path">D:/DNC/{escape(name.replace(" ", "_"))}</div><button>폴더 선택</button></div>'
        for name in rows
    )
    if grouped:
        return f"""
        <div class="settings-preview grouped">
          <div class="setting-group"><h4>DNC 조건 폴더</h4>{folder_rows}</div>
          <div class="setting-group small"><h4>작업일보 / 백업</h4><button>작업일보 선택</button><button>마스터 백업</button></div>
          <div class="setting-group small"><h4>관리</h4><button>라이선스</button><button>조건 마스터</button></div>
        </div>
        """
    return f"""
    <div class="settings-preview">
      <div class="section-title">DNC 조건 시트 폴더 설정</div>
      {folder_rows}
    </div>
    """


def render_alert_mockup(variant: dict) -> str:
    return """
    <div class="alert-stage">
      <div class="alert-box">
        <div class="alert-title">확인 필요</div>
        <div class="alert-body">LOT 1 / LOT 2 LOT No 동일<br>동시 신규 검증 불가</div>
        <button class="primary">확인</button>
      </div>
    </div>
    """


def make_variants() -> list[dict]:
    variants: list[dict] = []
    number = 1
    for layout_name, desc, risk in MAIN_LAYOUTS:
        for palette in PALETTES:
            variants.append({
                "no": number,
                "kind": "메인",
                "layout": layout_name,
                "palette": palette,
                "desc": desc,
                "risk": risk,
                "render": render_main_mockup,
            })
            number += 1
    for layout_name, desc, risk in POPUP_LAYOUTS:
        for palette in PALETTES:
            variants.append({
                "no": number,
                "kind": "신규 DNC 팝업",
                "layout": layout_name,
                "palette": palette,
                "desc": desc,
                "risk": risk,
                "render": render_popup_mockup,
            })
            number += 1
    for layout_name, desc, risk in SETTING_LAYOUTS:
        for palette in PALETTES:
            variants.append({
                "no": number,
                "kind": "설정",
                "layout": layout_name,
                "palette": palette,
                "desc": desc,
                "risk": risk,
                "render": render_settings_mockup,
            })
            number += 1
    for index, palette in enumerate(PALETTES, start=1):
        variants.append({
            "no": number,
            "kind": "알림/확인 팝업",
            "layout": f"Operator Alert {index}",
            "palette": palette,
            "desc": "오류/확인 팝업을 크게, 한 문장 중심으로 보여주는 안전한 변경",
            "risk": "low",
            "render": render_alert_mockup,
        })
        number += 1
    return variants


def render_card(variant: dict) -> str:
    palette_name, primary, light, bg, text, accent = variant["palette"]
    style = f"--primary:{primary};--light:{light};--bg:{bg};--text:{text};--accent:{accent};"
    risk_text = {"low": "적용 쉬움", "medium": "적용 중간"}.get(variant["risk"], "검토 필요")
    mockup = variant["render"](variant)
    return f"""
    <article class="card" data-kind="{escape(variant['kind'])}" data-risk="{escape(variant['risk'])}"{html_attrs(style)}>
      <header class="card-head">
        <div>
          <div class="meta">#{variant['no']:03d} · {escape(variant['kind'])} · {escape(palette_name)}</div>
          <h2>{escape(variant['layout'])}</h2>
        </div>
        <span class="risk {escape(variant['risk'])}">{escape(risk_text)}</span>
      </header>
      <p>{escape(variant['desc'])}</p>
      <div class="swatches"><span></span><span></span><span></span><span></span><span></span></div>
      {mockup}
    </article>
    """


def build_html() -> str:
    variants = make_variants()
    cards = "\n".join(render_card(variant) for variant in variants)
    return f"""<!doctype html>
<html lang="ko">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>JIIN DNC V4 Design Gallery</title>
  <style>
    * {{ box-sizing: border-box; }}
    body {{ margin: 0; font-family: "Malgun Gothic", "Segoe UI", sans-serif; background: #f3f6fb; color: #142033; }}
    .hero {{ position: sticky; top: 0; z-index: 5; background: rgba(255,255,255,.96); border-bottom: 1px solid #d8e0eb; padding: 18px 26px 14px; }}
    .hero h1 {{ margin: 0 0 6px; font-size: 24px; }}
    .hero p {{ margin: 0; color: #607085; }}
    .toolbar {{ display: flex; gap: 8px; flex-wrap: wrap; margin-top: 14px; }}
    .toolbar button {{ border: 1px solid #b7c4d6; background: #fff; padding: 8px 12px; cursor: pointer; font-weight: 700; color: #1f334d; }}
    .toolbar button.active {{ background: #1d4ed8; color: #fff; border-color: #1d4ed8; }}
    .wrap {{ padding: 22px; display: grid; grid-template-columns: repeat(auto-fill, minmax(620px, 1fr)); gap: 18px; }}
    .card {{ --primary:#1d4ed8; --light:#dbeafe; --bg:#eff6ff; --text:#0f172a; --accent:#60a5fa; background: #fff; border: 1px solid #d7e0ec; box-shadow: 0 10px 26px rgba(15,23,42,.07); padding: 16px; border-radius: 8px; }}
    .card-head {{ display: flex; justify-content: space-between; align-items: flex-start; gap: 12px; border-bottom: 1px solid #e6edf5; padding-bottom: 10px; }}
    .meta {{ font-size: 12px; color: #64748b; font-weight: 700; }}
    h2 {{ margin: 4px 0 0; font-size: 18px; color: var(--text); }}
    .card p {{ min-height: 42px; color: #415168; line-height: 1.45; margin: 12px 0; }}
    .risk {{ white-space: nowrap; border: 1px solid #cbd5e1; padding: 5px 8px; font-size: 12px; font-weight: 800; }}
    .risk.low {{ background: #dcfce7; color: #166534; }}
    .risk.medium {{ background: #fef3c7; color: #92400e; }}
    .swatches {{ display: flex; gap: 6px; margin: 0 0 12px; }}
    .swatches span {{ width: 34px; height: 14px; border: 1px solid rgba(0,0,0,.12); }}
    .swatches span:nth-child(1) {{ background: var(--primary); }}
    .swatches span:nth-child(2) {{ background: var(--light); }}
    .swatches span:nth-child(3) {{ background: var(--bg); }}
    .swatches span:nth-child(4) {{ background: var(--text); }}
    .swatches span:nth-child(5) {{ background: var(--accent); }}
    button {{ border: 1px solid #9ca8b8; background: #fff; min-height: 30px; padding: 5px 10px; font-family: inherit; }}
    button.primary, .process-head button:first-of-type, .modal-actions button:first-child {{ background: var(--primary); border-color: var(--primary); color: #fff; font-weight: 800; }}
    .mock-main, .modal-preview, .settings-preview, .alert-stage {{ background: var(--bg); border: 1px solid #b9c6d8; min-height: 360px; padding: 10px; color: var(--text); }}
    .topbar, .process-head, .modal-title {{ background: var(--light); border-bottom: 1px solid #c3d0df; padding: 10px; display: flex; align-items: center; justify-content: space-between; font-weight: 900; }}
    .brand {{ font-size: 16px; }}
    .logo {{ color: var(--primary); letter-spacing: 0; font-size: 18px; }}
    .tabs, .lot-switch {{ display: flex; gap: 0; margin: 8px 0; }}
    .tabs span, .lot-switch span {{ border: 1px solid #cbd5e1; background: #fff; padding: 7px 12px; min-width: 78px; text-align: center; font-weight: 700; }}
    .tabs .active, .lot-switch .active {{ background: var(--light); color: var(--primary); border-color: var(--primary); }}
    .process-head {{ gap: 8px; }}
    .process-head b {{ flex: 1; text-align: center; }}
    .common-strip {{ display: grid; grid-template-columns: repeat(5, 1fr); gap: 6px; margin: 8px 0; }}
    .common-strip span {{ background: #fff; border: 1px solid #d7dee9; padding: 7px; font-size: 12px; }}
    .main-body, .popup-grid {{ display: grid; grid-template-columns: 1fr 1fr; gap: 10px; align-items: start; }}
    .side-rail {{ display: grid; gap: 7px; }}
    .main-body:has(.side-rail) {{ grid-template-columns: 1fr 1fr 150px; }}
    .lot-card {{ background: #fff; border: 1px solid #c4cfdd; padding: 10px; min-height: 210px; }}
    .lot-card h4 {{ margin: -10px -10px 10px; background: var(--light); padding: 8px; text-align: center; color: var(--primary); }}
    .field-row {{ display: grid; grid-template-columns: 80px 1fr 80px 1fr; gap: 6px; align-items: center; margin-bottom: 7px; }}
    .label {{ background: #f2f5f9; padding: 6px; text-align: right; font-size: 12px; }}
    .input {{ border: 1px solid #a9b3c2; min-height: 28px; padding: 6px; color: #075ed8; background: #fff; font-size: 12px; overflow: hidden; white-space: nowrap; text-overflow: ellipsis; }}
    .status-wide {{ display: grid; grid-template-columns: repeat(3,1fr); gap: 8px; margin: 9px 0; }}
    .pill {{ display: block; padding: 9px; border: 1px solid #cdd6e3; background: #fff; text-align: center; font-weight: 800; }}
    .pill.ok {{ background: #dcfce7; color: #047857; border-color: #10b981; }}
    .pill.wait {{ background: #f8fafc; color: #64748b; }}
    .log {{ margin-top: 10px; min-height: 54px; background: #fff; border: 1px solid #d5deea; padding: 8px; font-family: Consolas, monospace; font-size: 12px; }}
    .compact .common-strip span, .compact .field-row {{ font-size: 11px; margin-bottom: 4px; }}
    .modal-preview {{ min-height: 390px; }}
    .steps {{ display: grid; grid-template-columns: repeat(3,1fr); gap: 6px; margin: 8px 0; }}
    .steps span {{ background: #fff; border: 1px solid #ccd6e3; padding: 7px; text-align: center; font-weight: 800; }}
    .steps .on {{ background: var(--primary); color: #fff; }}
    .guard-line {{ margin: 9px 0; background: #fff; border: 1px solid #ccd6e3; padding: 10px; font-weight: 900; color: var(--primary); }}
    .guard-line.danger {{ background: #fee2e2; border-color: #ef4444; color: #b91c1c; }}
    .modal-actions {{ display: flex; gap: 8px; justify-content: space-between; margin-top: 10px; }}
    .dense .field-row {{ grid-template-columns: 70px 1fr 70px 1fr; gap: 4px; margin-bottom: 5px; }}
    .settings-preview .section-title {{ text-align: center; background: var(--light); color: var(--primary); font-weight: 900; padding: 8px; margin-bottom: 10px; }}
    .folder-row {{ display: grid; grid-template-columns: 120px 1fr 120px; gap: 8px; align-items: center; margin: 8px 0; }}
    .folder-row span {{ background: var(--light); color: var(--primary); font-weight: 800; padding: 9px; text-align: center; }}
    .path {{ background: #fff; border: 1px solid #b8c2d0; min-height: 34px; padding: 8px; color: #506176; }}
    .settings-preview.grouped {{ display: grid; grid-template-columns: 1fr 220px 220px; gap: 10px; }}
    .setting-group {{ background: #fff; border: 1px solid #c7d1df; padding: 10px; }}
    .setting-group h4 {{ margin: 0 0 8px; color: var(--primary); }}
    .setting-group.small button {{ display: block; width: 100%; margin-bottom: 8px; }}
    .alert-stage {{ display: grid; place-items: center; background: color-mix(in srgb, var(--bg), #000 8%); }}
    .alert-box {{ width: 360px; background: #fff; border: 1px solid #b7c3d4; box-shadow: 0 18px 50px rgba(15,23,42,.18); text-align: center; }}
    .alert-title {{ color: #dc2626; font-size: 22px; font-weight: 900; padding: 18px 16px 8px; }}
    .alert-body {{ padding: 12px 20px 20px; line-height: 1.6; font-weight: 800; }}
    .alert-box button {{ margin: 0 0 18px; width: 140px; }}
    @media (max-width: 760px) {{
      .wrap {{ grid-template-columns: 1fr; padding: 12px; }}
      .main-body, .popup-grid, .common-strip, .settings-preview.grouped {{ grid-template-columns: 1fr; }}
      .field-row, .folder-row {{ grid-template-columns: 84px 1fr; }}
      .field-row .label:nth-of-type(2), .field-row .input:nth-of-type(2) {{ display: none; }}
    }}
  </style>
</head>
<body>
  <section class="hero">
    <h1>JIIN DNC Manager V4 디자인 후보 100</h1>
    <p>실행파일이 아니라 스크롤로 비교하는 시각 디자인 후보집입니다. Tkinter에서 무리 없이 적용 가능한 레이아웃, 색상, 버튼, 팝업 중심으로만 구성했습니다.</p>
    <div class="toolbar">
      <button class="active" data-filter="all">전체 100</button>
      <button data-filter="메인">메인</button>
      <button data-filter="신규 DNC 팝업">신규 DNC 팝업</button>
      <button data-filter="설정">설정</button>
      <button data-filter="알림/확인 팝업">알림/확인 팝업</button>
      <button data-filter="low">적용 쉬움</button>
    </div>
  </section>
  <main class="wrap">
    {cards}
  </main>
  <script>
    const buttons = document.querySelectorAll('.toolbar button');
    const cards = document.querySelectorAll('.card');
    buttons.forEach(button => button.addEventListener('click', () => {{
      buttons.forEach(item => item.classList.remove('active'));
      button.classList.add('active');
      const filter = button.dataset.filter;
      cards.forEach(card => {{
        const show = filter === 'all' || card.dataset.kind === filter || card.dataset.risk === filter;
        card.style.display = show ? '' : 'none';
      }});
    }}));
  </script>
</body>
</html>
"""


def main() -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    OUT_FILE.write_text(build_html(), encoding="utf-8")
    print(OUT_FILE)


if __name__ == "__main__":
    main()
