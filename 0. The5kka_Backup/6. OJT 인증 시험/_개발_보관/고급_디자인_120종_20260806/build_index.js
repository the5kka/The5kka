const fs = require('fs');
const path = require('path');
const concepts = ['통합 관제 데스크','대상별 시작 화면','시험지 문서 스튜디오','전체 페이지 월','단계별 생성 마법사','검색 중심 콘솔','공정·대상 매트릭스','현장 터치 키오스크','컴팩트 사이드바','한 문제 집중 화면','페이지 개요 보드','신입 교육 경로','시험 템플릿 라이브러리','출력 대기열 센터','감사 이력 타임라인','다국어 시험 스테이션','시험 준비 칸반','시험 일정 캘린더','품질 지표 콕핏','리본형 업무 공간','문항 검토 인스펙터','미니멀 실행 화면','고대비 접근성 화면','선정·출력 듀얼 화면','공정 흐름 블루프린트','종이 우선 미리보기','대상별 밴드 선택','시험 생성 플로우 맵','작업자 브리핑 화면','시험 패키지 탐색기'];
const roles = ['일반용','전장용','신입용','외국인용'];
const recommended = new Set([2,3,8,10,16,26]);
const lines = ['디자인번호,구조번호,전체화면 구조,시험대상,추천24종'];
for (let i = 1; i <= 120; i += 1) {
  const c = Math.floor((i - 1) / 4) + 1;
  const r = (i - 1) % 4;
  lines.push(`${String(i).padStart(3,'0')},${String(c).padStart(2,'0')},${concepts[c-1]},${roles[r]},${recommended.has(c)?'추천':''}`);
}
fs.writeFileSync(path.join(__dirname, '디자인_번호표.csv'), '\ufeff' + lines.join('\r\n'), 'utf8');
fs.writeFileSync(path.join(__dirname, '추천_24종_번호.txt'), '추천 구조: 02, 03, 08, 10, 16, 26\r\n추천 디자인 번호:\r\n' + [...recommended].flatMap(c => [1,2,3,4].map(r => String((c-1)*4+r).padStart(3,'0'))).join(', '), 'utf8');
