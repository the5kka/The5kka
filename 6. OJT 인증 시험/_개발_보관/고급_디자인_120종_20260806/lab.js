const roles = [
  { key: "general", name: "일반용", code: "GENERAL", icon: "clipboard-list", process: "PRESS 일반용", detail: "객관식 20 · 주관식 4", tag: "정기 인증" },
  { key: "electrical", name: "전장용", code: "ELECTRICAL", icon: "cpu", process: "PRESS 전장용", detail: "VDA 2 · 객관식 20 · 주관식 3", tag: "VDA 포함" },
  { key: "newcomer", name: "신입용", code: "NEWCOMER", icon: "graduation-cap", process: "액분석 신입용", detail: "객관식 20 · 주관식 4", tag: "교육 연계" },
  { key: "foreigner", name: "외국인용", code: "GLOBAL", icon: "languages", process: "PRESS 외국인용", detail: "언어 선택 · 객관식 · 주관식", tag: "추후 적용" }
];

const builders = [];

builders.push(
  () => `${topbar()}<div class="body"><main class="page-gallery">${miniPages(6,0)}</main><aside class="panel overview-side">${headingBlock("페이지 개요",`${role.process} · 6페이지`)}<div style="display:grid;align-content:start;gap:12px">${metrics()}${panel("페이지 구성","files",listRows(6))}${conditions()}</div><div style="display:grid;grid-template-columns:1fr 1fr;gap:8px">${B("출력 미리보기","scan-search","secondary")}${B("문제 + 답안 프린트","printer")}</div></aside></div>`,

  () => `${topbar()}<div class="body"><section>${headingBlock("교육 진행",`${role.name} 인증 경로`)}<div class="path-flow" style="margin-top:20px">${[["book-open","기초 교육"],["clipboard-list","문제 학습"],["badge-check","중간 확인"],["file-check-2","인증 시험"],["award","완료"]].map((x,i)=>`<div class="path-step ${i<3?"done":""}"><div class="dot">${I(x[0])}</div><b>${x[1]}</b><span>${i<3?"완료":"대기"}</span></div>`).join("")}</div></section><section style="display:grid;grid-template-columns:1fr 360px;gap:14px">${panel("추천 시험","sparkles",`<div style="padding:18px;display:grid;gap:14px">${headingBlock("NEXT",role.process)}${metrics()}${listRows(4)}</div>`)}${panel("작업자","user-round",`<div style="padding:16px;display:grid;gap:14px">${identity()}${conditions()}${B("인증 시험 시작","play")}</div>`)}</section><section class="panel" style="padding:14px;display:flex;align-items:center;gap:18px">${I("shield-check")}<b>출력 전 문제은행, 문항 수, 배점, 작업자 정보를 자동으로 검사합니다.</b><span style="flex:1"></span>${S("검증 정상","info")}</section></div>`,

  () => `<aside class="side-nav"><div class="side-brand">OJT <b>LIBRARY</b></div>${roles.map((r,i)=>`<div class="nav-item ${i===roleIndex?"active":""}">${I(r.icon)}<span>${r.name}</span></div>`).join("")}<div class="nav-item">${I("archive")}<span>보관함</span></div><div class="nav-bottom">${B("설정","settings","secondary")}</div></aside><main class="library-main"><header class="topbar compact"><div class="brand">시험 템플릿</div><div class="search-box" style="flex:1">${I("search")}<input value="${role.process}" aria-label="템플릿 검색"></div>${B("새 템플릿","plus")}</header><section class="library-grid">${["표준 인증 시험","정기 재평가","신규 작업자","그림 문항 중심","VDA 집중형","현장 단축형","다국어 기본형","관리자 검토형","출력 보관본"].map((x,i)=>`<article class="template-card"><div class="template-preview"><div class="mini-page">${Array.from({length:6},()=>"<div></div>").join("")}</div><div><b>${x}</b><span>${role.name}<br>${i%2?"최근 사용 3일 전":"오늘 사용"}</span></div></div><div class="action-row">${S(i===0?"기본":"보관","info")}<span class="grow"></span>${I("more-horizontal")}</div></article>`).join("")}</section></main>`,

  () => `${topbar()}<div class="body"><main class="queue-main"><section class="printer-row">${[["printer","현장 프린터 1","준비"],["printer-check","품질 사무실","연결됨"],["hard-drive","PDF 보관함","사용 가능"]].map((x,i)=>`<article class="printer">${I(x[0])}<div><b>${x[1]}</b><span>${x[2]} · A4 세로</span></div></article>`).join("")}</section>${panel("출력 대기열","list-ordered",`<div class="list">${["PRESS 일반용 · 오국진","TRIM 전장용 · 박지영","X-Ray 일반용 · 이왕건","수직흑화 신입용 · 김수현","CZ 전처리 전장용 · 정유진","출하검사 일반용 · 최민석","Lay-up 외국인용 · 준비"].map((x,i)=>`<div class="list-row"><span class="index">${i+1}</span><div><b>${x}</b><small>${i===0?"문제 4페이지 + 답안지 1페이지":"대기 중"}</small></div>${S(i===0?"출력 중":"대기",i===0?"warn":"")}</div>`).join("")}</div>`)}</main><aside class="panel" style="padding:14px;display:grid;grid-template-rows:auto 1fr auto;gap:14px">${headingBlock("선택 작업",role.process)}<div class="doc-stage">${docPage()}</div><div style="display:grid;gap:9px">${identity()}${B("대기열에 추가","printer")}</div></aside></div>`,

  () => `${topbar()}<div class="body"><aside class="panel" style="padding:18px">${headingBlock("시험 이력",role.process)}<div class="timeline" style="margin-top:20px">${[["08:42","문제은행 새로고침"],["08:43","공정 선택"],["08:44","랜덤 문항 생성"],["08:45","문항·배점 검증"],["08:46","출력 미리보기"],["08:48","문제 + 답안 출력"]].map((x,i)=>`<div class="timeline-item"><b>${x[1]}</b><span>${x[0]} · ${i<5?"정상 완료":"현재 단계"}</span></div>`).join("")}</div></aside><main style="min-width:0;display:grid;grid-template-rows:170px 1fr;gap:12px">${panel("감사 요약","history",`<div style="padding:14px">${metrics()}</div>`)}${panel("세부 기록","rows-3",listRows(8))}</main><aside style="display:grid;grid-template-rows:1fr 260px;gap:12px">${panel("증빙 문서","file-check-2",`<div class="doc-stage" style="height:100%">${docPage()}</div>`)}${panel("현재 사용자","user-round",`<div style="padding:14px;display:grid;gap:14px">${identity()}${B("이력 내보내기","download","secondary")}</div>`)}</aside></div>`,

  () => `${topbar()}<div class="body"><section>${headingBlock("언어 선택",`${role.name} 시험 스테이션`)}<div class="language-grid" style="margin-top:18px">${[["KR","한국어"],["EN","English"],["VI","Tiếng Việt"],["CN","中文"]].map((x,i)=>`<div class="language ${i===(roleIndex===3?1:0)?"active":""}"><b>${x[0]}</b><span>${x[1]}</span></div>`).join("")}</div></section><section style="display:grid;grid-template-columns:1fr 1fr;gap:14px">${panel("한국어 시험 정보","file-text",`<div style="padding:18px;display:grid;gap:14px">${headingBlock("KOREAN",role.process)}${conditions()}${identity()}</div>`)}${panel("Selected language","globe-2",`<div style="padding:18px;display:grid;gap:14px">${headingBlock("ENGLISH", roleIndex===3?"Global Operator Exam":"Korean / English Support")}<div class="checklist">${["Question bank loaded","Score total verified","Worker name entered","Print pages ready"].map((x,i)=>`<div class="check"><span class="box">✓</span><b>${x}</b><span>OK</span></div>`).join("")}</div></div>`)}</section><div class="bilingual-actions"><article class="bi-action"><h3>화면 시험</h3><p>한 문제씩 큰 글씨로 표시합니다.<br>Show one question at a time.</p>${B("시작 / START","play")}</article><article class="bi-action"><h3>전체 페이지</h3><p>모든 시험 페이지를 확인합니다.<br>Review all exam pages.</p>${B("보기 / VIEW","files","secondary")}</article><article class="bi-action"><h3>인쇄</h3><p>문제와 답안지를 함께 출력합니다.<br>Print exam and answer sheet.</p>${B("출력 / PRINT","printer","secondary")}</article></div></div>`,

  () => `${topbar()}<div class="body"><div class="action-row"><div>${headingBlock("준비 현황",`${role.name} 시험 보드`)}</div><span class="grow"></span>${B("새 시험 카드","plus")}</div><main class="kanban">${[["요청 접수",["PRESS 일반용","TRIM 전장용"]],["문항 검토",[role.process,"X-Ray 일반용","수직흑화 신입용"]],["출력 준비",["CZ 전처리 전장용","출하검사 일반용"]],["완료",["Lay-up 일반용","ITS 각인 전장용"]]].map((col,ci)=>`<section class="kanban-col"><h3>${col[0]}<span>${col[1].length}</span></h3>${col[1].map((x,i)=>`<article class="kanban-card"><b>${x}</b><small>${ci===1&&i===0?role.detail:"문제은행 검증 완료"}</small><div class="action-row" style="margin-top:8px">${S(ci===3?"완료":ci===2?"출력":"진행",ci===2?"warn":"info")}</div></article>`).join("")}</section>`).join("")}</main></div>`,

  () => `${topbar()}<div class="body"><main class="panel" style="padding:14px;display:grid;grid-template-rows:auto 1fr;gap:12px"><div class="action-row">${headingBlock("2026년 8월","인증 시험 일정")}<span class="grow"></span>${B("오늘","calendar-check-2","secondary")}${B("일정 추가","plus")}</div><div class="calendar">${["월","화","수","목","금","토","일"].map(x=>`<div class="weekday">${x}</div>`).join("")}${Array.from({length:35},(_,i)=>`<div><b>${i<4?i+29:i-3}</b>${[3,6,9,13,17,22,26,30].includes(i)?`<div class="event">${i===9?role.process:"정기 인증"}</div>`:""}</div>`).join("")}</div></main><aside class="panel" style="padding:14px;display:grid;grid-template-rows:auto 1fr auto;gap:14px">${headingBlock("선택 일정","8월 6일 목요일")}${listRows(7)}${B("시험 준비 열기","arrow-right")}</aside></div>`,

  () => `${topbar()}<div class="body"><section class="metric-wide">${[["오늘 생성","18건","+4"],["출력 완료","16건","정상"],["검증 오류","0건","안정"],["문제은행","1,284문항","08:42"]].map((x,i)=>`<article class="big-metric"><span>${x[0]}</span><b>${x[1]}</b><small>${x[2]}</small></article>`).join("")}</section>${panel("대상별 사용량","bar-chart-3",`<div style="padding:18px"><div class="bars">${[62,88,49,34,76,54,91,66,43,82].map((h,i)=>`<div class="bar" style="height:${h}%"><span>${i+1}일</span></div>`).join("")}</div></div>`)}${panel("검증 현황","shield-check",`<div style="padding:14px;display:grid;gap:12px">${metrics()}${listRows(7)}</div>`)}</div>`,

  () => `<div class="ribbon"><div class="ribbon-tabs">${["파일","시험 생성","문항","보기","출력","관리"].map((x,i)=>`<div class="ribbon-tab ${i===1?"active":""}">${x}</div>`).join("")}</div><div class="ribbon-tools">${[["folder-open","파일 선택"],["refresh-cw","새로고침"],["shuffle","랜덤 생성"],["scan-search","미리보기"],["zoom-in","확대"],["printer","인쇄"],["settings","설정"]].map((x,i)=>`<div class="tool-group"><div class="tool">${I(x[0])}<span>${x[1]}</span></div></div>`).join("")}</div></div><div class="body">${panel("공정 목록","factory",processTable(8))}<main class="doc-stage">${docPage(true)}</main><aside style="display:grid;grid-template-rows:auto 1fr auto;gap:12px">${panel("문서 속성","sliders-horizontal",`<div style="padding:12px">${conditions()}</div>`)}${panel("출력 검사","list-checks",listRows(6))}${B("문제 + 답안 프린트","printer")}</aside></div>`
);

builders.push(
  () => `<header class="topbar compact"><div class="brand">문항 검토 <b>INSPECTOR</b></div><span style="flex:1"></span>${roleTabs()}${B("변경 승인","check","secondary")}${B("시험 생성","play")}</header><div class="canvas-body"><main class="question-sheet">${headingBlock("문항 원본",role.process)}<div style="margin-top:18px">${["작업표준서와 실제 작업 방법이 다를 때 가장 적절한 조치는?","설비 점검 주기가 지났으나 생산 일정이 촉박하다. 어떻게 해야 하는가?","제품 표면의 조도 상태를 올바르게 판정한 것은?","현장 작업 중 이상 냄새가 느껴질 때 가장 먼저 해야 하는 행동은?","최종 수세단 필터 교체 주기로 맞는 것은?","검사 결과가 기준을 벗어났을 때 기록해야 할 내용은?","작업 전 확인해야 하는 안전 조건은?","불량 제품 격리 후 관리자가 확인할 항목은?","교대 시 반드시 전달해야 하는 정보는?","주관식 답안에 포함할 핵심 단어를 쓰시오."].map((x,i)=>`<div class="question-line"><b>${i+1}.</b><span>${x}${i<7?`<br><small>① 즉시 중지　② 관리자 보고　③ 상태 기록　④ 기준 확인</small>`:""}</span><em>(${i<2?"2.5":i<8?"4":"5"}점)</em></div>`).join("")}</div></main><aside class="floating-inspector">${headingBlock("선택 문항","문항 03 속성")}<div class="field-grid"><div class="field"><label>유형</label><div class="control">객관식</div></div><div class="field"><label>배점</label><div class="control">4.0</div></div></div>${panel("자동 점검","shield-check",listRows(7))}<div style="display:grid;gap:8px">${B("다른 문항으로 교체","rotate-cw","secondary")}${B("현재 문항 확정","check-circle-2")}</div></aside></div>`,

  () => `<header class="minimal-head"><div class="brand">OJT <b>EXAM</b></div>${roleTabs()}<div class="action-row">${B("파일","folder-open","secondary")}${B("설정","lock-keyhole","secondary")}</div></header><section class="minimal-command"><article class="minimal-main">${headingBlock("READY",`${role.process} 시험을 준비합니다`)}<div style="margin-top:22px">${B("시험 생성","play")}</div></article>${[["문제은행","1,284 문항","database"],["작업자","오국진","user-round"],["평가 일시","2026.08.06","calendar-days"]].map(x=>`<article class="minimal-action"><div>${I(x[2])}<h3>${x[0]}</h3></div><div><b>${x[1]}</b><span>확인 완료</span></div></article>`).join("")}</section><main style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:18px">${panel("공정","factory",processTable(6))}${panel("검증","shield-check",listRows(7))}${panel("출력","printer",`<div style="height:100%;padding:18px;display:grid;grid-template-rows:1fr auto;gap:14px"><div class="doc-stage">${docPage()}</div>${B("출력 미리보기","scan-search","secondary")}</div>`)}</main>`,

  () => `<header class="contrast-head"><div><b style="font-size:25px">OJT EXAM MAKER</b><span style="margin-left:18px">고대비 현장 모드</span></div><div class="action-row">${S("EXCEL 연결됨","info")}${B("관리자","lock-keyhole","secondary")}</div></header><main class="contrast-grid">${roles.map((r,i)=>`<article class="contrast-tile ${i===roleIndex?"active":""}"><div>${I(r.icon)}</div><div><h2>${r.name}</h2><p>${r.process}<br>${r.detail}</p></div></article>`).join("")}</main><footer style="display:flex;align-items:center;gap:18px;border:3px solid #fff;padding:0 20px;background:#111"><b>선택: ${role.name}</b><span style="flex:1"></span>${B("출력 미리보기","scan-search","secondary")}${B("시험 시작","play")}</footer>`,

  () => `${topbar()}<div class="body"><section class="dual-side"><div class="action-row">${headingBlock("STEP 1","공정과 문항 선택")}<span class="grow"></span>${S("자동 저장","info")}</div>${panel("공정 목록","factory",processTable(8))}<div class="action-row">${B("무작위 선택","shuffle","secondary")}<span class="grow"></span>${B("선택 확정","check")}</div></section><section class="dual-side"><div class="action-row">${headingBlock("STEP 2","출력 문서 확인")}<span class="grow"></span>${B("전체 페이지","files","secondary")}</div><div class="doc-stage">${docPage(true)}</div>${commonActions()}</section></div>`,

  () => `<header class="panel" style="padding:0 18px;display:flex;align-items:center;gap:16px"><div class="brand">OJT <b>PROCESS MAP</b></div><span style="flex:1"></span>${roleTabs()}${B("관리자 설정","lock-keyhole","secondary")}</header><main class="flow-map blueprint-bg">${[["database","문제은행","Excel 갱신"],["factory","공정 선택",role.process],["shuffle","문항 생성",role.detail],["printer-check","출력","문제 + 답안"]].map((x,i)=>`<article class="flow-node ${i===1?"active":""}">${I(x[0])}<b>${x[1]}</b><span>${x[2]}</span></article>`).join("")}</main><footer class="panel" style="padding:12px 16px;display:flex;align-items:center;gap:12px">${S("2단계","info")}<b>${role.process} 선택 완료</b><span style="flex:1"></span>${B("이전","arrow-left","secondary")}${B("다음 단계","arrow-right")}</footer>`,

  () => `<aside class="tool-rail">${[["folder-open","파일"],["refresh-cw","갱신"],["zoom-in","확대"],["zoom-out","축소"],["files","페이지"],["printer","출력"],["settings","설정"]].map(x=>`<button class="btn secondary" title="${x[1]}">${I(x[0])}</button>`).join("")}</aside><main class="paperfirst-stage">${docPage(true)}<span style="position:absolute;left:18px;bottom:16px">1 / 5 페이지</span></main><aside class="paperfirst-side">${headingBlock("종이 우선",role.process)}<div style="display:grid;align-content:start;gap:12px">${identity()}${panel("출력 검사","shield-check",listRows(6))}${conditions()}</div>${B("문제 + 답안 프린트","printer")}</aside>`,

  () => `<header style="display:flex;align-items:center;gap:16px"><div class="brand">OJT <b>ROLE BOARD</b></div><span style="flex:1"></span>${B("파일 선택","folder-open","secondary")}${B("설정","lock-keyhole")}</header><main class="role-bands">${roles.map((r,i)=>`<article class="role-band ${i===roleIndex?"active":""}"><div>${I(r.icon)}</div><b>${r.name}</b><p>${r.process}<br>${r.detail}</p>${i===roleIndex?B("선택됨","check"):I("chevron-right")}</article>`).join("")}</main><section style="display:grid;grid-template-columns:1fr 1fr 360px;gap:14px">${panel("최근 사용 공정","history",listRows(4))}${panel("문제은행 상태","database",`<div style="padding:16px;display:grid;gap:14px">${metrics()}${S("자동 갱신 완료","info")}</div>`)}<div class="panel" style="padding:16px;display:grid;align-content:center;gap:12px">${identity()}${B("선택 대상 시작","play")}</div></section>`,

  () => `<header style="display:flex;align-items:center;gap:16px"><div class="brand">OJT <b>EXAM FLOW</b></div><span style="flex:1"></span>${roleTabs()}${B("설정","lock-keyhole","secondary")}</header><main class="flow-map">${[["folder-open","파일 연결","문제은행 자동 갱신"],["users-round","대상 선택",role.name],["factory","공정 선택",role.process],["shuffle","문항 생성",role.detail]].map((x,i)=>`<article class="flow-node ${i===2?"active":""}">${I(x[0])}<b>${x[1]}</b><span>${x[2]}</span></article>`).join("")}</main><section style="display:grid;grid-template-columns:1fr 1fr;gap:12px">${panel("다음 단계","arrow-right",`<div style="padding:14px">${identity()}</div>`)}<div class="panel" style="padding:16px;display:flex;align-items:center;gap:12px">${S("3 / 6","info")}<b>공정 선택이 완료되었습니다.</b><span style="flex:1"></span>${B("다음","arrow-right")}</div></section>`,

  () => `<main class="brief-main"><div class="brief-title">${headingBlock("작업자 브리핑",`${role.name} 인증 시험`)}</div><section class="start-block"><div class="big-icon">${I(role.icon)}</div><h2>${role.process}</h2><p>성명과 평가 일시를 확인한 뒤 시험을 시작하세요.</p>${B("시험 준비 시작","play")}</section><div class="action-row">${B("다른 대상 선택","users-round","secondary")}${B("전체 페이지 보기","files","secondary")}<span class="grow"></span>${S("문제은행 최신","info")}</div></main><aside class="brief-side">${headingBlock("시작 전 확인","CHECKLIST")}<div class="checklist">${["시험 대상 확인","공정명 확인","작업자 성명 확인","평가 일시 확인","문항 수·점수 확인","프린터 연결 확인","답안지 포함 확인"].map((x,i)=>`<div class="check"><span class="box">✓</span><b>${x}</b><span>${i<6?"완료":"자동"}</span></div>`).join("")}</div><div style="display:grid;gap:10px">${identity()}${B("문제 + 답안 프린트","printer")}</div></aside>`,

  () => `<header class="topbar compact"><div class="brand">OJT <b>PACKAGE EXPLORER</b></div><div class="path">${I("hard-drive","small")}<span>시험 패키지 / ${role.name} / ${role.process}</span></div>${B("새 패키지","package-plus")}</header><div class="explorer-body"><aside class="explorer-tree"><div class="tree"><div class="tree-item active">${I("home")}빠른 시작</div>${roles.map((r,i)=>`<div class="tree-item ${i===roleIndex?"active":""}">${I(r.icon)}${r.name}</div>`).join("")}<div class="tree-item">${I("archive")}출력 보관함</div><div class="tree-item">${I("history")}최근 작업</div><div class="tree-item">${I("settings")}관리자 설정</div></div></aside><main class="explorer-main"><div class="action-row"><div class="search-box" style="flex:1">${I("search")}<input value="${role.process}" aria-label="패키지 검색"></div>${B("보기","layout-grid","secondary")}${B("정렬","arrow-up-down","secondary")}</div>${panel("시험 패키지","package-open",processTable(8))}</main><aside class="explorer-detail">${headingBlock("패키지 정보",role.process)}${metrics()}${panel("포함 파일","files",listRows(6))}<div style="display:grid;gap:9px">${B("출력 미리보기","scan-search","secondary")}${B("패키지 실행","play")}</div></aside></div>`
);

window.addEventListener("DOMContentLoaded", () => {
  const laterBuilders = builders.slice();
  builders.splice(0, builders.length, ...firstTenBuilders, ...laterBuilders);
  const fixedWizard = builders[4];
  builders[4] = () => fixedWizard().replace("${role.name}", role.name);
  const vars = {
    "--bg": theme.bg, "--surface": theme.surface, "--surface2": theme.surface2,
    "--stage": theme.stage, "--ink": theme.ink, "--muted": theme.muted,
    "--line": theme.line, "--accent": theme.a[roleIndex], "--accent2": theme.a2,
    "--accent3": theme.a3, "--button-ink": theme.dark ? "#081014" : "#ffffff"
  };
  Object.entries(vars).forEach(([key, value]) => app.style.setProperty(key, value));
  app.className = `app view-${concept[1]} btn-style-${(conceptIndex % 8) + 1}${theme.dark ? " dark" : ""}`;
  app.innerHTML = builders[conceptIndex]();
  app.insertAdjacentHTML("beforeend", `<div class="screen-note">DESIGN ${String(designNo).padStart(3,"0")} · ${concept[0]} · ${role.name}</div>`);
  if (window.lucide) window.lucide.createIcons({ attrs: { "aria-hidden": "true" } });
});

const concepts = [
  ["통합 관제 데스크", "command"], ["대상별 시작 화면", "launchpad"], ["시험지 문서 스튜디오", "document"],
  ["전체 페이지 월", "printwall"], ["단계별 생성 마법사", "wizard"], ["검색 중심 콘솔", "search"],
  ["공정·대상 매트릭스", "matrix"], ["현장 터치 키오스크", "kiosk"], ["컴팩트 사이드바", "sidebar"],
  ["한 문제 집중 화면", "focus"], ["페이지 개요 보드", "overview"], ["신입 교육 경로", "training"],
  ["시험 템플릿 라이브러리", "library"], ["출력 대기열 센터", "queue"], ["감사 이력 타임라인", "timeline"],
  ["다국어 시험 스테이션", "bilingual"], ["시험 준비 칸반", "kanban"], ["시험 일정 캘린더", "calendar"],
  ["품질 지표 콕핏", "metrics"], ["리본형 업무 공간", "ribbon"], ["문항 검토 인스펙터", "inspector"],
  ["미니멀 실행 화면", "minimal"], ["고대비 접근성 화면", "contrast"], ["선정·출력 듀얼 화면", "dual"],
  ["공정 흐름 블루프린트", "blueprint"], ["종이 우선 미리보기", "paperfirst"], ["대상별 밴드 선택", "rolebands"],
  ["시험 생성 플로우 맵", "flow"], ["작업자 브리핑 화면", "brief"], ["시험 패키지 탐색기", "explorer"]
];

const themes = [
  { bg:"#e9eff0",surface:"#ffffff",surface2:"#dbe5e6",stage:"#cbd6d8",ink:"#162328",muted:"#607178",line:"#a8b7bb",a:["#087f74","#2563eb","#24834b","#c44b3f"],a2:"#2c6e91",a3:"#f0b429" },
  { bg:"#161b1f",surface:"#22292e",surface2:"#303940",stage:"#101417",ink:"#f2f6f7",muted:"#9cabb2",line:"#4b5960",a:["#3ad6c7","#f28c28","#8fd14f","#ff6f91"],a2:"#54a9ff",a3:"#f2c84b",dark:true },
  { bg:"#f1f1ef",surface:"#ffffff",surface2:"#e8e6e2",stage:"#d8d6d1",ink:"#211b1c",muted:"#726a6b",line:"#bbb4b5",a:["#8d2638","#005eb8","#2f7d4a","#c2415d"],a2:"#126e82",a3:"#d49b1f" },
  { bg:"#e8eef7",surface:"#ffffff",surface2:"#d7e1ef",stage:"#c7d3e2",ink:"#17243a",muted:"#617087",line:"#9dadc0",a:["#08766d","#164db5","#3f7f35","#ba3f74"],a2:"#b8472d",a3:"#e0a51d" },
  { bg:"#edf2ed",surface:"#ffffff",surface2:"#dce6dd",stage:"#cfdacf",ink:"#1a2a20",muted:"#68766c",line:"#a9b8ac",a:["#176c42","#2368b0","#5b7f22","#bd4d2e"],a2:"#2e7f82",a3:"#d7a51c" },
  { bg:"#0d1518",surface:"#172126",surface2:"#24333a",stage:"#091012",ink:"#edfafa",muted:"#8fa8ad",line:"#3b535a",a:["#29c6b8","#39a7ff","#9ed23a","#ff776b"],a2:"#ff9d3c",a3:"#f7da4a",dark:true },
  { bg:"#f3f4f5",surface:"#ffffff",surface2:"#e1e4e7",stage:"#d2d6da",ink:"#1e2328",muted:"#687078",line:"#b3bac0",a:["#c62828","#1565c0","#2e7d32","#ad3b74"],a2:"#006d77",a3:"#f0a202" },
  { bg:"#eceff4",surface:"#fbfcfe",surface2:"#dde2ea",stage:"#cdd4df",ink:"#202637",muted:"#697185",line:"#abb4c3",a:["#15756f","#3153a4","#4c7c30","#aa466e"],a2:"#bd5c2c",a3:"#d9a51f" },
  { bg:"#11251f",surface:"#1c332b",surface2:"#29483d",stage:"#0b1c17",ink:"#f2f8f5",muted:"#9bb3a8",line:"#426759",a:["#35c390","#63a9ff","#b5d33d","#ff8f55"],a2:"#f2c14e",a3:"#ffdf6e",dark:true },
  { bg:"#efefec",surface:"#ffffff",surface2:"#e2e0db",stage:"#d4d1ca",ink:"#28241f",muted:"#746e65",line:"#bbb5ab",a:["#0d7a70","#2764aa","#477e32","#d05b31"],a2:"#b33f62",a3:"#e3ad22" },
  { bg:"#eaf0f3",surface:"#ffffff",surface2:"#d9e3e8",stage:"#cad7dc",ink:"#17262c",muted:"#62757d",line:"#a4b6bd",a:["#007c83","#2163b0","#3e7f48","#bf4961"],a2:"#d15c2b",a3:"#e6b227" },
  { bg:"#151719",surface:"#24272a",surface2:"#34393d",stage:"#0d0f10",ink:"#f5f6f6",muted:"#a3aaae",line:"#50575c",a:["#42c9b8","#5aa5f5","#99ce4a","#ff736c"],a2:"#ffad42",a3:"#ffe04f",dark:true },
  { bg:"#f0edf3",surface:"#ffffff",surface2:"#e2dfe7",stage:"#d3ced9",ink:"#26212a",muted:"#716a77",line:"#b5adbd",a:["#14786e","#3b59a3","#4f7c35","#b44475"],a2:"#c35f32",a3:"#dda923" },
  { bg:"#f2f3f3",surface:"#ffffff",surface2:"#e3e6e7",stage:"#d3d8da",ink:"#171b1d",muted:"#646d71",line:"#aeb7bb",a:["#00837b","#006cc1","#378044","#ca3e55"],a2:"#ef7b22",a3:"#e8b21e" },
  { bg:"#e8f1ee",surface:"#ffffff",surface2:"#d6e6e1",stage:"#c6d8d2",ink:"#172723",muted:"#60746e",line:"#9fb7af",a:["#16756a","#2866ae","#527c2f","#c14f42"],a2:"#9a4778",a3:"#dda91f" }
];

const qs = new URLSearchParams(location.search);
const designNo = Math.max(1, Math.min(120, Number(qs.get("d") || 1)));
const conceptIndex = Math.floor((designNo - 1) / 4);
const roleIndex = (designNo - 1) % 4;
const role = roles[roleIndex];
const concept = concepts[conceptIndex];
const theme = themes[conceptIndex % themes.length];
const app = document.getElementById("app");

const I = (name, cls = "") => `<i data-lucide="${name}" class="icon ${cls}"></i>`;
const B = (label, icon = "arrow-right", cls = "") => `<button class="btn ${cls}">${I(icon)}<span>${label}</span></button>`;
const S = (text, cls = "") => `<span class="status ${cls}">${text}</span>`;
const panel = (title, icon, content, extra = "") => `<section class="panel ${extra}"><div class="panel-head"><h2>${I(icon)}${title}</h2><span>OJT EXAM</span></div>${content}</section>`;
const topbar = (compact = false, actions = true) => `<header class="topbar ${compact ? "compact" : ""}">
  <div class="brand">OJT <b>EXAM MAKER</b><small>QUALITY TRAINING SYSTEM</small></div>
  <div class="path">${I("file-spreadsheet", "small")}<span>OJT 시험 문제.xlsm · 자동 새로고침 완료</span></div>
  ${roleTabs()}
  ${actions ? `<div class="action-row">${B("파일 선택", "folder-open", "secondary")}${B("설정", "settings")}</div>` : ""}
</header>`;

function roleTabs() {
  return `<div class="role-tabs">${roles.map((r, i) => `<button class="role-tab ${i === roleIndex ? "active" : ""}">${I(r.icon, "small")}<span>${r.name}</span></button>`).join("")}</div>`;
}

function roleCards() {
  return `<div class="role-card-grid">${roles.map((r, i) => `<article class="role-card ${i === roleIndex ? "active" : ""}">
    <div class="role-icon">${I(r.icon)}</div><div><h3>${r.name}</h3><p>${r.process}<br>${r.detail}</p></div><small>${i === roleIndex ? "현재 선택" : r.tag}</small>
  </article>`).join("")}</div>`;
}

function processTable(limit = 8) {
  const names = roleIndex === 1
    ? ["PRESS 전장용","TRIM 전장용","X-Ray 전장용","CZ 전처리 전장용","노바본드 전장용","출하검사 전장용","ITS 각인 전장용","Dualdetatch 전장용"]
    : roleIndex === 2
      ? ["액분석 신입용","수직흑화 신입용","PRESS 신입용","TRIM 신입용","출하검사 신입용","CZ 전처리 신입용","Lay-up 신입용","수입검사 신입용"]
      : roleIndex === 3
        ? ["PRESS 외국인용","TRIM 외국인용","X-Ray 외국인용","수직흑화 외국인용","출하검사 외국인용","Lay-up 외국인용","CZ 전처리 외국인용","수입검사 외국인용"]
        : ["PRESS 일반용","TRIM 일반용","X-Ray 일반용","CZ 전처리 일반용","노바본드 일반용","출하검사 일반용","ITS 각인 일반용","Dualdetatch 일반용"];
  return `<table class="process-table"><thead><tr><th>공정명</th><th class="num">문항</th><th class="num">배점</th><th class="num">상태</th></tr></thead><tbody>${names.slice(0, limit).map((n,i)=>`<tr class="${i===0?"active":""}"><td>${n}</td><td class="num">${i===0?(roleIndex===1?25:24):20+i}</td><td class="num">100</td><td class="num">${i===0?"선택":"대기"}</td></tr>`).join("")}</tbody></table>`;
}

function conditions() {
  const rows = roleIndex === 1 ? [["VDA",2,"2.5"],["객관식",20,"4.0"],["주관식",3,"5.0"]] : [["객관식",20,"4.0"],["주관식",4,"5.0"]];
  return `<div class="conditions">${rows.map(r=>`<div class="condition"><b>${r[0]}</b><span>${r[1]}</span><span>${r[2]}</span></div>`).join("")}</div><div class="total-band"><span>목표 점수</span><b>TOTAL 100</b></div>`;
}

function identity() {
  return `<div class="field-grid"><div class="field"><label>성명</label><div class="control">오국진${I("user-round","small")}</div></div><div class="field"><label>평가 일시</label><div class="control">2026.08.06${I("calendar-days","small")}</div></div></div>`;
}

function listRows(n = 7) {
  const rows = ["문제은행 자동 새로고침","공정별 문항 수 검증","그림 포함 문항 확인","배점 합계 100점 확인","성명·평가일 입력 확인","A4 페이지 배치 완료","답안지 마지막 페이지 생성","출력 장치 연결 확인"];
  return `<div class="list">${rows.slice(0,n).map((x,i)=>`<div class="list-row"><span class="index">${String(i+1).padStart(2,"0")}</span><div><b>${x}</b><small>${i<4?role.process:"자동 검사 항목"}</small></div>${S(i<5?"완료":"대기",i===5?"warn":"")}</div>`).join("")}</div>`;
}

function metrics() {
  return `<div class="metric-row"><div class="metric"><span>사용 가능 문항</span><b>${roleIndex===1?184:152}</b></div><div class="metric"><span>그림 문항</span><b>${roleIndex===1?31:18}</b></div><div class="metric"><span>최근 갱신</span><b>08:42</b></div></div>`;
}

function docPage(large = false, page = 1) {
  const q = page === 1 ? [
    "[VDA 6.3] 작업표준서와 실제 작업 방법이 다를 때 가장 적절한 조치는?",
    "설비 점검 주기가 지났으나 생산 일정이 촉박하다. 어떻게 해야 하는가?",
    "제품 표면에 형성되는 조도 상태를 올바르게 판정한 것은?",
    "작업 중 이상 냄새가 발생했을 때 가장 먼저 해야 하는 행동은?",
    "최종 수세단 필터 교체 주기로 맞는 것은?"
  ] : ["Carrier Detach 공정의 핵심 관리 항목은?","작업 전 확인해야 하는 설비 조건은?","불량 발견 시 기록 순서를 쓰시오.","제품 격리 후 관리자에게 보고할 내용은?"];
  return `<div class="doc-page ${large?"large":""}"><div class="paper-head"><div class="paper-table"><span>부서명</span><span>가공팀</span><span>평가자</span><span>오국진</span><span>평가 일시</span><span>2026.08.06</span></div><div class="paper-title">교육 평가서</div><div class="paper-approval"><span>기안</span><span>심의</span><span>결정</span><span></span><span></span><span></span><span>/</span><span>/</span><span>/</span></div></div><div class="paper-info"><span>직무명</span><span>${role.process}</span><span>성명</span><span>오국진</span><span>평가 방법</span><span style="grid-column:2/5">시험 평가, 결과 보고서, 직무 평가</span><span>개정 차수</span><span>1</span><span>제정일</span><span>2026.08.06</span></div><div class="paper-questions">${q.map((x,i)=>`<div class="paper-q"><b>${i+1+(page-1)*5}.</b><span>${x}</span><b>(${i<2?"2.5":"4"}점)</b><span class="opts">① 작업 중지　② 관리자 보고　③ 상태 기록　④ 기준 확인</span></div>`).join("")}</div><div class="paper-footer"><span>JIQP-0202-02</span><span>( 주 ) 지 인</span><span>A4(210*297)mm</span></div></div>`;
}

function miniPages(count = 4, active = 0) {
  return Array.from({length:count},(_,i)=>`<div class="page-thumb ${i===active?"active":""}"><div class="mini-page">${Array.from({length:8},()=>"<div></div>").join("")}</div></div>`).join("");
}

function questionCard() {
  return `<article class="question-card"><div class="question-number"><span>QUESTION 08 / 24</span>${S(role.tag,"info")}</div><h2>현장 작업 중 이상 상태를 발견했을 때 가장 먼저 해야 하는 행동은 무엇입니까?</h2><div class="answers">${["작업을 계속하며 경과를 본다","제품을 격리하고 관리자에게 알린다","다음 교대자에게만 전달한다","검사 결과가 나올 때까지 보류한다"].map((x,i)=>`<div class="answer ${i===1?"selected":""}"><span class="circle">${i+1}</span><span>${x}</span></div>`).join("")}</div><div class="action-row"><span class="grow"></span>${B("이전", "arrow-left", "secondary")}${B("다음 문제", "arrow-right")}</div></article>`;
}

function commonActions() {
  return `<div class="action-row"><span class="grow"></span>${B("랜덤 미리보기","shuffle","secondary")}${B("출력 미리보기","scan-search","secondary")}${B("문제 + 답안 프린트","printer")}</div>`;
}

function headingBlock(kicker = "시험 생성", title = `${role.process} 인증 시험`) {
  return `<div><span class="eyebrow">${kicker.toUpperCase()}</span><h1 class="heading">${title}</h1><p class="subcopy">대상: ${role.name} · ${role.detail} · 문제은행 갱신 08:42</p></div>`;
}

const firstTenBuilders = [
  () => `${topbar()}<div class="body"><aside class="panel profile-rail">${headingBlock("대상 선택",role.name)}<div class="role-stack">${roles.map((r,i)=>`<button class="role-tab ${i===roleIndex?"active":""}">${I(r.icon)}<span>${r.name}</span></button>`).join("")}</div>${B("관리자 설정","lock-keyhole","secondary")}</aside><main class="center">${panel("시험 상태","activity",`<div style="padding:12px">${metrics()}</div>`)}${panel("공정 선택","factory",processTable(8))}${panel("출력 준비","printer-check",`<div style="padding:12px;display:grid;gap:10px">${identity()}${commonActions()}</div>`)}</main><aside class="right">${panel("시험지 페이지","file-text",`<div class="doc-stage" style="height:100%">${docPage()}</div>`)}${panel("자동 검증","shield-check",listRows(3))}</aside></div>`,

  () => `${topbar()}<div class="body"><section class="welcome">${headingBlock("오늘의 시험",`${role.name} 시험 시작`)}<div class="action-row">${B("최근 작업 열기","history","secondary")}${B("새 시험 만들기","plus")}</div></section>${roleCards()}<section class="recent-band">${panel("최근 사용 공정","clock-3",listRows(6))}${panel("바로 실행","play",`<div style="padding:18px;display:grid;gap:14px">${identity()}${B(role.process + " 시작","play")}${B("출력 미리보기","scan-search","secondary")}</div>`)}</section></div>`,

  () => `${topbar()}<div class="body"><aside class="thumb-rail"><div class="panel-head" style="margin:-10px -10px 0"><h3>${I("files")}페이지</h3></div>${miniPages(4,0)}</aside><main class="doc-stage">${docPage(true)}</main><aside class="inspector">${headingBlock("문서 속성","출력 미리보기")}<div class="action-row">${B("축소","zoom-out","secondary")}${B("확대","zoom-in","secondary")}${B("맞춤","maximize-2","secondary")}</div>${panel("시험 정보","sliders-horizontal",`<div style="padding:12px;display:grid;gap:12px">${identity()}${conditions()}</div>`)}${B("문제 + 답안 프린트","printer")}</aside></div>`,

  () => `${topbar(true)}<div class="body"><div class="wall-toolbar">${headingBlock("전체 페이지",`${role.process} · 6페이지`)}<span style="flex:1"></span>${B("페이지 맞춤","grid-2x2","secondary")}${B("답안지 포함","clipboard-check","secondary")}${B("인쇄","printer")}</div><main class="wall">${miniPages(8,7)}</main></div>`,

  () => `${topbar()}<div class="body"><aside class="panel stepper">${[["대상 선택",role.name],["공정 선택",role.process],["문항 구성",role.detail],["작업자 정보","성명·일자"],["검토 및 출력","최종 확인"]].map((x,i)=>`<div class="step ${i===2?"active":""}"><span class="step-no">${i+1}</span><div><b>${x[0]}</b><span>${x[1]}</span></div></div>`).join("")}</aside><main class="panel form-pane">${headingBlock("3 / 5", "문항 구성을 확인하세요")}<div style="display:grid;align-content:start;gap:16px">${conditions()}${identity()}${panel("출제 규칙","list-checks",listRows(4))}</div><div class="action-row">${B("이전","arrow-left","secondary")}<span class="grow"></span>${B("다음 단계","arrow-right")}</div></main><aside class="panel summary-pane">${headingBlock("실시간 요약",role.process)}${metrics()}<div class="doc-stage">${docPage()}</div>${B("현재 상태 저장","save","secondary")}</aside></div>`,

  () => `${topbar()}<div class="body"><section class="search-hero"><div>${headingBlock("빠른 검색","공정명이나 직무명을 입력하세요")}<div class="search-box" style="margin-top:12px">${I("search")}<input value="${role.process}" aria-label="공정 검색">${S("32개 결과","info")}</div></div><div class="action-row wrap">${B("최근 공정","history","secondary")}${B("즐겨찾기","star","secondary")}${B("상세 필터","filter")}</div></section><section class="result-grid">${panel("검색 결과","rows-3",processTable(8))}${panel("선택한 시험","check-circle-2",`<div style="padding:14px;display:grid;gap:14px">${headingBlock("SELECTED",role.process)}${conditions()}${identity()}</div>`)}</section>${commonActions()}</div>`,

  () => `${topbar()}<div class="body"><section class="panel" style="padding:14px"><div class="matrix"><div class="m-head">공정 / 대상</div>${roles.map(r=>`<div class="m-head">${I(r.icon,"small")}${r.name}</div>`).join("")} ${["PRESS","TRIM","X-Ray","CZ 전처리","Lay-up","출하검사","수직흑화","수입검사","ITS 각인","노바본드"].map((p,pi)=>`<div class="m-process">${p}</div>${roles.map((r,ri)=>`<div class="m-cell ${pi===0&&ri===roleIndex?"active":""}">${pi===0&&ri===roleIndex?"선택됨":ri===3&&pi>4?"준비":"사용"}</div>`).join("")}`).join("")}</div></section><section style="display:grid;grid-template-columns:1fr 320px;gap:12px">${panel("선택 정보","mouse-pointer-2",`<div style="padding:14px">${metrics()}</div>`)}<div style="display:grid;gap:10px">${B("출력 미리보기","scan-search","secondary")}${B("시험 생성","sparkles")}</div></section></div>`,

  () => `<header class="kiosk-head"><div class="brand">OJT <b>EXAM STATION</b></div><div class="action-row">${S("문제은행 최신","info")}${B("관리자","lock-keyhole","secondary")}</div></header><main class="kiosk-grid">${roles.map((r,i)=>`<article class="kiosk-tile ${i===roleIndex?"active":""}"><div class="kiosk-icon">${I(r.icon)}</div><div><h2>${r.name}</h2><p>${r.process}<br>${r.detail}</p></div>${I("chevron-right")}</article>`).join("")}</main><footer class="panel" style="padding:12px 18px;display:flex;align-items:center;gap:18px">${I("circle-help")}<b>화면에서 시험 대상을 선택하세요.</b><span style="flex:1"></span>${I("wifi")}<span>EXCEL 연결됨</span><span>2026.08.06 08:42</span></footer>`,

  () => `<aside class="side-nav"><div class="side-brand">OJT <b>EXAM</b></div>${[["home","홈"],["factory","공정 선택"],["file-text","시험 생성"],["scan-search","출력 미리보기"],["printer","출력 이력"],["settings","설정"]].map((x,i)=>`<div class="nav-item ${i===2?"active":""}">${I(x[0])}<span>${x[1]}</span></div>`).join("")}<div class="nav-bottom">${B("종료","log-out","secondary")}</div></aside><main class="workspace"><div class="topbar compact"><div class="brand">시험 생성</div><span style="flex:1"></span>${roleTabs()}${B("설정","lock-keyhole","secondary")}</div><div class="workspace-body">${panel("공정","factory",processTable(8))}<section style="min-width:0;display:grid;grid-template-rows:150px 1fr 130px;gap:12px">${panel("시험 정보","clipboard-check",`<div style="padding:12px">${metrics()}</div>`)}${panel("문항 구성","list-checks",listRows(8))}${panel("작업자 정보","user-round",`<div style="padding:12px">${identity()}</div>`)}</section>${panel("출력","printer",`<div style="height:calc(100% - 42px);padding:14px;display:grid;grid-template-rows:auto 1fr auto;gap:12px"><div>${conditions()}</div><div class="doc-stage">${docPage()}</div>${B("문제 + 답안 프린트","printer")}</div>`)}</div></main>`,

  () => `<aside class="focus-side">${headingBlock("문항 이동",role.process)}<div class="number-grid">${Array.from({length:25},(_,i)=>`<div class="number ${i===7?"active":""}">${i+1}</div>`).join("")}</div><div style="display:grid;gap:9px">${B("문항 다시 뽑기","shuffle","secondary")}${B("전체 페이지 보기","files","secondary")}</div></aside><main class="focus-main">${questionCard()}${commonActions()}</main><aside class="focus-meta">${headingBlock("시험 정보",role.name)}${identity()}${panel("현재 구성","sliders-horizontal",`<div style="padding:12px">${conditions()}</div>`)}</aside>`
];
