// deck.js를 그대로 쓰는 최소 예제.
//   node example.js  →  out/example.pptx
// 확인:
//   soffice --headless --convert-to pdf --outdir out out/example.pptx
const fs = require("fs");
const path = require("path");
const api = new Function(
  fs.readFileSync(path.join(__dirname, "deck.js"), "utf8") +
  "; return { buildPptx, PT };"
)();
const { buildPptx, PT } = api;

const slides = [
  // 표지
  { visual: { type: "cover", eyebrow: "MONTHLY PRODUCTION REVIEW",
      title: "월간 생산실적 보고", subtitle: "2026-09   ·   생산본부",
      meta: ["투입 1,250톤", "달성율 105%", "종합가동율 92.4%", "작성 2026-10-01"] } },

  // 1. 요약 — 액션 타이틀 + 큰 숫자 3~4개 + 시사점 3가지
  { kicker: "요약", accent: PT.green,
    title: "9월 생산량 125.4만 개로 목표 5% 초과 달성,\n수율 99%대 안정세 진입",
    subtitle: "Executive Summary : Key Performance Indicators",
    visual: { type: "kpi", items: [
      { label: "총 생산량 (만 개)", value: "125.4", color: PT.text,
        delta: "목표 대비 +5.4% (초과)", deltaColor: PT.green, deltaDir: "up" },
      { label: "종합 수율", value: "99.2", unit: "%", color: PT.green,
        delta: "전월 대비 +0.3%p", deltaColor: PT.green, deltaDir: "up" },
      { label: "설비 가동률", value: "95.0", unit: "%", color: PT.amber,
        delta: "C라인 정비로 소폭 하락", deltaColor: PT.amber, deltaDir: "down" },
    ] },
    points: [
      { label: "생산", color: PT.green, text: "목표 119만 개 대비 125.4만 개 달성 · 달성율 105.4%" },
      { label: "품질", color: PT.green, text: "종합 수율 99.2% · 3개월 연속 상승" },
      { label: "제약", color: PT.red, text: "C라인 조립 공정 불량률 증가 · 전체 불량의 45% 점유" },
    ] },

  // 2. 문제 — 미달 하나만 빨강
  { kicker: "문제", accent: PT.red,
    title: "A/B 라인은 목표를 초과 달성했으나,\nC라인은 초기 수율 안정화 지연으로 미달",
    subtitle: "Target vs. Actual by Line (단위: %)",
    visual: { type: "bars", mark: 100, markLabel: "목표 100%", rows: [
      { label: "Line A (주력)", value: 108, text: "108%", color: PT.green },
      { label: "Line B (신규)", value: 102, text: "102%", color: PT.green },
      { label: "Line C (특수)", value: 85, text: "85%", color: PT.red, strong: true },
    ] },
    banner: { kind: "action", label: "Action Required",
      text: "C라인 조립 공정 불량률 증가가 주요 원인 · 금주 내 엔지니어 TF 투입 예정" } },

  // 3. 추이 — 값은 점 위에, 축 눈금 없음
  { kicker: "분석", accent: PT.green,
    title: "상반기 설비 투자 효과로 6월 기점 생산량 우상향 추세 전환",
    subtitle: "Monthly Production Trend (최근 6개월)",
    visual: { type: "trend", color: PT.green, rows: [
      { label: "4월", value: 98, text: "98" }, { label: "5월", value: 92, text: "92" },
      { label: "6월", value: 105, text: "105" }, { label: "7월", value: 112, text: "112" },
      { label: "8월", value: 119, text: "119" }, { label: "9월", value: 125, text: "125" },
    ] },
    banner: { kind: "insight", label: "Insight",
      text: "5월 라인 병목 해소(신규 장비 도입) 이후 매월 평균 7% 증가폭 기록" } },

  // 4. 현안 과제 — 아이콘 + 우선순위 + 현상/원인/대응
  { kicker: "대안", accent: PT.red,
    title: "상위 3개 현안에 대한 대응 방안 수립,\nC라인 불량 우선 조치로 수율 1.2%p 회복 목표",
    subtitle: "Key Issues & Action Plans",
    visual: { type: "issues", cards: [
      { title: "C라인 외관 스크래치 불량", icon: "gear", color: PT.red, badge: "긴급 (High Priority)",
        rows: [
          { k: "현상", v: "이송 컨베이어 벨트 마모로 인한 스크래치 다수 발생 (불량 비중 45%)" },
          { k: "원인", v: "벨트 교체 주기(6개월) 초과" },
          { k: "대응", v: "9/10(금) 야간 교체 작업 배정 완료" } ] },
      { title: "핵심 부품(센서) 입고 지연", icon: "truck", color: PT.blue, badge: "모니터링 (Medium)",
        rows: [
          { k: "현상", v: "A사 물류 파업으로 센서 입고 3일 지연" },
          { k: "영향", v: "안전 재고 1주 분량으로 방어 중이나 장기화 시 리스크" },
          { k: "대응", v: "B사 대체품 긴급 수배 (샘플 테스트 중)" } ] },
      { title: "추석 연휴 특근 편성", icon: "cycle", color: PT.green, badge: "완료 (Resolved)",
        rows: [
          { k: "현상", v: "연휴 기간 납기 준수를 위한 라인 가동 필요" },
          { k: "대응", v: "노사 협의 완료 · 연휴 기간 2교대 체제 가동 확정" },
          { k: "효과", v: "예상 차질 물량 100% 커버 가능" } ] },
    ] } },

  // 5. 실행계획 — 항목 / 담당 / 기한
  { kicker: "실행계획", accent: PT.blue,
    title: "4건 실행 과제 확정, 즉시 착수 2건 포함",
    subtitle: "Action Plan (2026-10)",
    visual: { type: "plan", rows: [
      { title: "C라인 컨베이어 벨트 교체", detail: "불량률 45% → 10% 목표", owner: "설비팀", due: "즉시", color: PT.red },
      { title: "센서 대체품 승인 시험", detail: "B사 샘플 3종 평가", owner: "품질팀", due: "즉시", color: PT.amber },
      { title: "벨트 교체 주기 표준 재설정", detail: "6개월 → 4개월", owner: "공무팀", due: "차월 중", color: PT.blue },
      { title: "10월 계획 1,300톤 확정 및 원료 확보", detail: "조업 29일", owner: "생산팀", due: "착수 전", color: PT.blue },
    ] } },
];

const out = path.join(__dirname, "out");
fs.mkdirSync(out, { recursive: true });
const bytes = buildPptx(slides.map((s) => ({ ...s, footer: "생산본부 · 2026-09" })));
fs.writeFileSync(path.join(out, "example.pptx"), Buffer.from(bytes));
console.log("out/example.pptx", bytes.length, "bytes");
