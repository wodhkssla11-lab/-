// ---------- 경영진 보고서용 PPT 조립기 ----------
// 라이브러리 없이 .pptx를 만듭니다. .pptx는 XML 몇 개를 담은 zip이거든요.
// 브라우저에서 바로 돌아가고, 노드에서 쓰려면 downloadFile만 파일 쓰기로 바꾸면 됩니다.
//
// 쓰는 법
//   const slides = [
//     { visual: { type: "cover", eyebrow: "MONTHLY REVIEW", title: "...", subtitle: "...", meta: ["..."] } },
//     { kicker: "요약", accent: PT.green, title: "액션 타이틀 1줄\n또는 2줄",
//       subtitle: "Executive Summary", visual: { type: "kpi", items: [...] }, points: [...] },
//   ];
//   downloadPptx("보고서.pptx", slides);
//
// 원칙(액션 타이틀 · 원포인트 컬러 · 데이터 다이어트 · 픽토그램)은 ../SKILL.md 참고.

const xmlEsc = (v) => String(v)
  .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

// 한글 조사 — 받침이 있으면 "이/을/은", 없으면 "가/를/는".
function josa(word, pair) {
  const [withBatchim, without] = pair;
  const t = String(word || "").trim();
  if (!t) return without;
  const ch = t.charCodeAt(t.length - 1);
  if (!(ch >= 0xAC00 && ch <= 0xD7A3)) return without;
  return (ch - 0xAC00) % 28 > 0 ? withBatchim : without;
}

// ---------- zip (OPC 패키지) ----------
const CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let i = 0; i < 256; i++) {
    let c = i;
    for (let k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
    t[i] = c >>> 0;
  }
  return t;
})();
function crc32(bytes) {
  let c = 0xFFFFFFFF;
  for (let i = 0; i < bytes.length; i++) c = CRC_TABLE[(c ^ bytes[i]) & 0xFF] ^ (c >>> 8);
  return (c ^ 0xFFFFFFFF) >>> 0;
}
const utf8Bytes = (str) => new TextEncoder().encode(str);

// 압축 없이(stored) 담는 zip — 장부 파일은 작아서 압축이 필요 없고, 구현이 단순해 깨질 여지가 없어요.
function makeZip(files) {
  const u16 = (n) => [n & 0xFF, (n >>> 8) & 0xFF];
  const u32 = (n) => [n & 0xFF, (n >>> 8) & 0xFF, (n >>> 16) & 0xFF, (n >>> 24) & 0xFF];
  const DOS_DATE = ((2024 - 1980) << 9) | (1 << 5) | 1; // 고정 날짜 (파일마다 결과가 같도록)
  const entries = files.map((f) => {
    const name = utf8Bytes(f.name);
    const data = utf8Bytes(f.content);
    return { name, data, crc: crc32(data) };
  });
  const parts = [];
  const central = [];
  let offset = 0;
  entries.forEach((f) => {
    const header = [].concat(
      u32(0x04034b50), u16(20), u16(0x0800), u16(0), u16(0), u16(DOS_DATE),
      u32(f.crc), u32(f.data.length), u32(f.data.length), u16(f.name.length), u16(0)
    );
    parts.push(new Uint8Array(header), f.name, f.data);
    central.push(new Uint8Array([].concat(
      u32(0x02014b50), u16(20), u16(20), u16(0x0800), u16(0), u16(0), u16(DOS_DATE),
      u32(f.crc), u32(f.data.length), u32(f.data.length),
      u16(f.name.length), u16(0), u16(0), u16(0), u16(0), u32(0), u32(offset)
    )));
    offset += header.length + f.name.length + f.data.length;
  });
  const cdOffset = offset;
  let cdSize = 0;
  central.forEach((c, i) => {
    parts.push(c, entries[i].name);
    cdSize += c.length + entries[i].name.length;
  });
  parts.push(new Uint8Array([].concat(
    u32(0x06054b50), u16(0), u16(0), u16(entries.length), u16(entries.length),
    u32(cdSize), u32(cdOffset), u16(0)
  )));
  const out = new Uint8Array(parts.reduce((sum, x) => sum + x.length, 0));
  let pos = 0;
  parts.forEach((x) => { out.set(x, pos); pos += x.length; });
  return out;
}


// rows: 2차원 배열. 숫자는 숫자 셀로 넣어서 엑셀에서 바로 합계/정렬이 돼요.
// 첫 줄은 머리글로 굵게 칠하고 틀고정(스크롤해도 안 사라짐)까지 해둬요.

function downloadFile(filename, content, mime) {
  const blob = new Blob([content], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

// ---------- 경영진 보고용 슬라이드 드로잉 ----------
// .pptx도 .xlsx처럼 XML 몇 개를 담은 zip이라, 라이브러리 없이 여기서 직접 조립해요.
// 좌표는 1280×720 기준으로 쓰고(16:9), 마지막에 EMU로 환산합니다.
//
// 디자인 원칙 네 가지 — 보고서를 받는 사람이 제목만 넘겨봐도 상황을 알 수 있게:
//   1) 액션 타이틀 : 제목이 곧 결론. "9월 생산 현황"이 아니라 "목표 105% 초과 달성"
//   2) 원 포인트 컬러 : 바탕은 무채색, 잘한 것 하나·못한 것 하나에만 색
//   3) 데이터 다이어트 : 큰 숫자 3~4개만 앞에, 상세 표는 별첨으로
//   4) 픽토그램 : 글이 많은 장표는 아이콘으로 먼저 읽히게
const EMU = 914400 / 96;
const SLIDE_W = 12192000, SLIDE_H = 6858000;
const px = (v) => Math.round(v * EMU);

// 무채색 바탕 + 포인트 컬러. 색이 많아지면 어디를 봐야 할지 알 수 없게 돼요.
const PT = {
  bg: "0F172A",        // 딥 네이비
  card: "1B2437",
  cardAlt: "232F45",
  line: "2E3B52",
  text: "FFFFFF",
  dim: "94A3B8",
  faint: "64748B",
  green: "10B981",     // 달성 · 개선
  red: "EF4444",       // 미달 · 문제  ← 시선이 여기로 꽂히게
  blue: "3B82F6",      // 중립 강조
  amber: "F59E0B",     // 주의
  redSoft: "2A1A22",
  greenSoft: "14262A",
  blueSoft: "16233C",
};
const hex = (c) => String(c == null ? "" : c).replace("#", "") || PT.dim;

// 투명도가 있으면 srgbClr를 열고 alpha를 넣은 뒤 닫아야 해요.
// 한 줄로 붙이려다 </a:srgbClr/> 같은 게 나오면 파워포인트가 파일을 아예 못 엽니다.
const solidFill = (color, alpha) =>
  alpha == null
    ? `<a:solidFill><a:srgbClr val="${hex(color)}"/></a:solidFill>`
    : `<a:solidFill><a:srgbClr val="${hex(color)}"><a:alpha val="${Math.round(alpha * 100000)}"/></a:srgbClr></a:solidFill>`;

let ppShapeId = 1;
const nextId = () => ++ppShapeId;

function ppShape(x, y, w, h, fill, opts = {}) {
  const o = opts || {};
  const ln = o.line
    ? `<a:ln w="${o.lineW || 12700}"><a:solidFill><a:srgbClr val="${hex(o.line)}"/></a:solidFill></a:ln>`
    : `<a:ln><a:noFill/></a:ln>`;
  const rot = o.rot ? ` rot="${Math.round(o.rot * 60000)}"` : "";
  return `<p:sp><p:nvSpPr><p:cNvPr id="${nextId()}" name="s${ppShapeId}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr><a:xfrm${rot}><a:off x="${px(x)}" y="${px(y)}"/><a:ext cx="${px(w)}" cy="${px(h)}"/></a:xfrm>` +
    `<a:prstGeom prst="${o.geom || "rect"}"><a:avLst/></a:prstGeom>` +
    (fill ? solidFill(fill, o.alpha) : `<a:noFill/>`) +
    ln + `</p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>`;
}
const ppRect = ppShape;

// 글상자. lines = [{ t, size, color, bold, bullet, align, spaceBefore }]
// t 에 배열을 넣으면 한 줄 안에서 색·크기를 섞을 수 있어요 (숫자만 크게 같은 것).
function ppText(x, y, w, h, lines, opts = {}) {
  const o = opts || {};
  const body = lines.filter(Boolean).map((l) => {
    const runs = (Array.isArray(l.t) ? l.t : [{ t: l.t }]).map((r) => {
      const c = hex(r.color || l.color || PT.text);
      const b = (r.bold != null ? r.bold : l.bold) ? 1 : 0;
      const sz = Math.round((r.size || l.size || 14) * 100);
      const sp = r.spc != null ? ` spc="${Math.round(r.spc * 100)}"` : "";
      return `<a:r><a:rPr lang="ko-KR" sz="${sz}" b="${b}"${sp} dirty="0">` +
        `<a:solidFill><a:srgbClr val="${c}"/></a:solidFill>` +
        `<a:latin typeface="Malgun Gothic"/><a:ea typeface="Malgun Gothic"/><a:cs typeface="Malgun Gothic"/></a:rPr>` +
        `<a:t>${xmlEsc(r.t)}</a:t></a:r>`;
    }).join("");
    const bu = l.bullet ? `<a:buFont typeface="Arial"/><a:buChar char="•"/>` : `<a:buNone/>`;
    return `<a:p><a:pPr algn="${l.align || o.align || "l"}" marL="${l.bullet ? 152400 : 0}" indent="${l.bullet ? -152400 : 0}">` +
      `<a:lnSpc><a:spcPct val="${l.lnSpc || o.lnSpc || 110000}"/></a:lnSpc>` +
      `<a:spcBef><a:spcPts val="${(l.spaceBefore || 0) * 100}"/></a:spcBef>${bu}</a:pPr>${runs}</a:p>`;
  }).join("");
  return `<p:sp><p:nvSpPr><p:cNvPr id="${nextId()}" name="t${ppShapeId}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="${px(x)}" y="${px(y)}"/><a:ext cx="${px(w)}" cy="${px(h)}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>` +
    `<p:txBody><a:bodyPr wrap="square" anchor="${o.anchor || "t"}" lIns="0" tIns="0" rIns="0" bIns="0">` +
    `<a:normAutofit/></a:bodyPr><a:lstStyle/>${body}</p:txBody></p:sp>`;
}

// 꺾은선 · 다각형 (추이 차트용)
function ppLine(pts, color, width = 2.5, opts = {}) {
  if (!pts || pts.length < 2) return "";
  const xs = pts.map((p) => p[0]), ys = pts.map((p) => p[1]);
  const x0 = Math.min(...xs), y0 = Math.min(...ys);
  const w = Math.max(1, Math.max(...xs) - x0), h = Math.max(1, Math.max(...ys) - y0);
  const path = pts.map((p, i) =>
    `<a:${i === 0 ? "moveTo" : "lnTo"}><a:pt x="${px(p[0] - x0)}" y="${px(p[1] - y0)}"/></a:${i === 0 ? "moveTo" : "lnTo"}>`).join("");
  const close = opts.close ? "<a:close/>" : "";
  return `<p:sp><p:nvSpPr><p:cNvPr id="${nextId()}" name="l${ppShapeId}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="${px(x0)}" y="${px(y0)}"/><a:ext cx="${px(w)}" cy="${px(h)}"/></a:xfrm>` +
    `<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/>` +
    `<a:pathLst><a:path w="${px(w)}" h="${px(h)}">${path}${close}</a:path></a:pathLst></a:custGeom>` +
    (opts.fill ? `<a:solidFill><a:srgbClr val="${hex(opts.fill)}"><a:alpha val="${Math.round((opts.fillAlpha != null ? opts.fillAlpha : 0.18) * 100000)}"/></a:srgbClr></a:solidFill>` : `<a:noFill/>`) +
    `<a:ln w="${Math.round(width * 12700)}" cap="rnd"><a:solidFill><a:srgbClr val="${hex(color)}"/></a:solidFill><a:round/></a:ln>` +
    `</p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>`;
}

// 픽토그램 — 아이콘 파일을 끌어오지 않고 파워포인트 기본 도형으로 그려요.
// (외부 아이콘은 라이선스도 걸리고, 파일에 이미지를 심으면 용량이 커져요)
const PICTO = {
  gear: "gear6", cycle: "donut", box: "cube", bolt: "lightningBolt",
  truck: "flowChartInputOutput", warn: "triangle", dot: "ellipse",
  up: "upArrow", down: "downArrow", right: "rightArrow", flag: "plaque",
  target: "donut", people: "flowChartMultidocument", chart: "flowChartMagneticDrum",
};
function ppIcon(x, y, size, kind, color) {
  const geom = PICTO[kind] || "ellipse";
  return ppShape(x, y, size, size, PT.card, { geom: "roundRect", line: color, lineW: 9525 }) +
    ppShape(x + size * 0.26, y + size * 0.26, size * 0.48, size * 0.48, color, { geom });
}

// 표 — 별첨에서만 씁니다. 본문에 로우 데이터를 붙이지 않아요.
function ppTable(x, y, w, rows, opts = {}) {
  const o = opts || {};
  const colW = o.colW || rows[0].map(() => w / rows[0].length);
  const rh = o.rowH || 25;
  let out = "";
  rows.forEach((row, ri) => {
    const ry = y + rh * ri;
    if (ri === 0) out += ppRect(x, ry, w, rh, PT.cardAlt);
    else if (ri % 2 === 0) out += ppRect(x, ry, w, rh, PT.card, { alpha: 0.6 });
    let cx = x;
    row.forEach((cell, ci) => {
      out += ppText(cx + 9, ry, colW[ci] - 18, rh, [{
        t: String(cell), size: 9.5, bold: ri === 0,
        color: ri === 0 ? PT.dim : PT.text, align: ci > 0 ? "r" : "l",
      }], { anchor: "ctr" });
      cx += colW[ci];
    });
  });
  return out;
}

// ── 슬라이드 골격 ─────────────────────────────────────────
// 왼쪽 세로 포인트 바 + 액션 타이틀 + 영문 부제. 첨부해주신 샘플 보고서와 같은 구성이에요.
function slideFrame(s, idx, total) {
  let out = ppRect(0, 0, 1280, 720, PT.bg);
  const titleLines = String(s.title || "").split("\n").length;
  // 26pt 한 줄이 대략 42px을 먹어요. 줄 수에 맞춰 부제 위치를 내려야 글자가 겹치지 않습니다.
  const lineH = 42;
  const titleH = lineH * titleLines;
  out += ppRect(64, 50, 5, titleH + (s.subtitle ? 26 : 0), s.accent || PT.green);
  out += ppText(88, 48, 1120, titleH + 8, [{ t: s.title, size: 26, bold: true, color: PT.text, lnSpc: 118000 }]);
  if (s.subtitle) out += ppText(88, 48 + titleH + 6, 1120, 24, [{ t: s.subtitle, size: 13, color: PT.dim }]);
  // 바닥선 + 쪽번호
  out += ppRect(64, 668, 1152, 1, PT.line);
  out += ppText(64, 682, 700, 16, [{ t: s.footer || "", size: 8.5, color: PT.faint }]);
  out += ppText(900, 682, 316, 16, [{ t: `${idx} / ${total}`, size: 8.5, color: PT.faint, align: "r" }]);
  return out;
}

// 하단 강조 배너 — "그래서 무엇을 할 것인가"를 한 줄로 못박아요.
function ppBanner(y, kind, label, text) {
  const color = kind === "action" ? PT.red : kind === "insight" ? PT.green : PT.blue;
  const soft = kind === "action" ? PT.redSoft : kind === "insight" ? PT.greenSoft : PT.blueSoft;
  let out = ppRect(64, y, 1152, 54, soft, { geom: "roundRect" });
  out += ppRect(64, y, 4, 54, color);
  out += ppShape(86, y + 19, 16, 16, color, { geom: kind === "action" ? "triangle" : "ellipse" });
  out += ppText(112, y, 1090, 54, [{ t: [
    { t: `${label}  `, size: 11.5, bold: true, color },
    { t: text, size: 11.5, color: PT.text },
  ] }], { anchor: "ctr" });
  return out;
}

// ── 장표 유형별 본문 ───────────────────────────────────────
// 1) KPI — 큰 숫자 3~4개만. 소수점 둘째 자리보다 추세와 비율이 궁금한 자리니까요.
function bodyKpi(v, s) {
  const items = v.items || [];
  const n = Math.max(1, items.length);
  const gap = 20, x0 = 64, w = (1152 - gap * (n - 1)) / n;
  let out = "";
  items.forEach((k, i) => {
    const x = x0 + (w + gap) * i;
    out += ppRect(x, 196, w, 168, PT.card, { geom: "roundRect", line: PT.line, lineW: 9525 });
    out += ppText(x + 26, 220, w - 52, 20, [{ t: k.label, size: 11, color: PT.dim }]);
    out += ppText(x + 26, 250, w - 52, 66, [{ t: [
      { t: k.value, size: 40, bold: true, color: k.color || PT.text },
      { t: k.unit ? ` ${k.unit}` : "", size: 14, color: PT.dim },
    ] }]);
    if (k.delta) {
      out += ppShape(x + 26, 328, 13, 13, k.deltaColor || PT.dim,
        { geom: k.deltaDir === "up" ? "upArrow" : k.deltaDir === "down" ? "downArrow" : "ellipse" });
      out += ppText(x + 46, 325, w - 72, 20, [{ t: k.delta, size: 10.5, bold: true, color: k.deltaColor || PT.dim }]);
    }
  });
  // 시사점 세 가지 — 결론 밑에 근거를 개조식으로.
  if (s.points && s.points.length) {
    out += ppText(64, 400, 1152, 24, [{ t: "핵심 시사점", size: 11, bold: true, color: PT.dim, spc: 1 }]);
    s.points.forEach((p, i) => {
      const y = 432 + i * 62;
      out += ppRect(64, y, 1152, 50, PT.card, { geom: "roundRect", alpha: 0.55 });
      out += ppShape(84, y + 17, 16, 16, p.color || PT.blue, { geom: "ellipse" });
      out += ppText(92, y + 17, 16, 16, [{ t: String(i + 1), size: 9, bold: true, color: PT.bg, align: "ctr" }]);
      out += ppText(116, y, 1080, 50, [{ t: [
        { t: `${p.label}  `, size: 11.5, bold: true, color: p.color || PT.text },
        { t: p.text, size: 11.5, color: PT.dim },
      ] }], { anchor: "ctr" });
    });
  }
  return out;
}

// 2) 가로 막대 — 목표선까지의 트랙을 깔고, 문제 있는 항목 하나만 빨강.
function bodyBars(v, s) {
  const rows = v.rows || [];
  if (rows.length === 0) return ppText(64, 300, 1152, 40, [{ t: "표시할 데이터가 없습니다.", size: 13, color: PT.faint, align: "ctr" }]);
  const hasBanner = !!s.banner;
  const top = 206, boxH = hasBanner ? 356 : 420;
  let out = ppRect(64, top, 1152, boxH, PT.card, { geom: "roundRect", alpha: 0.55 });
  const max = Math.max(1, ...rows.map((r) => Math.abs(Number(r.value) || 0)), v.axisMax || 0);
  const padY = 30, rowH = (boxH - padY * 2) / rows.length;
  const barH = Math.min(34, rowH * 0.52);
  const labelW = 206, valW = 132;
  const trackX = 64 + 30 + labelW, trackW = 1152 - 60 - labelW - valW;
  rows.forEach((r, i) => {
    const cy = top + padY + rowH * i + (rowH - barH) / 2;
    const val = Math.abs(Number(r.value) || 0);
    const c = hex(r.color || PT.blue);
    out += ppText(64 + 30, cy - 4, labelW - 14, barH + 8,
      [{ t: r.label, size: 11.5, bold: !!r.strong, color: r.strong ? c : PT.dim, align: "r" }], { anchor: "ctr" });
    out += ppRect(trackX, cy, trackW, barH, PT.cardAlt, { geom: "roundRect" });
    if (val > 0) out += ppRect(trackX, cy, Math.max(6, (val / max) * trackW), barH, c, { geom: "roundRect" });
    // 목표선 (100% 등) — 어디까지 갔어야 하는지 한눈에
    if (v.mark != null && v.mark > 0) {
      const mx = trackX + (v.mark / max) * trackW;
      if (i === 0) out += ppRect(mx, top + 18, 1.6, boxH - 36, PT.dim, { alpha: 0.55 });
    }
    // "5시간 50분 · 2건" 처럼 두 정보가 붙어 있으면 줄을 나눠요 —
    // 한 줄로 밀어넣으면 글꼴에 따라 어정쩡하게 접혀서 "건"만 다음 줄로 떨어집니다.
    const txt = r.text != null ? String(r.text) : String(val);
    const cut = txt.indexOf(" · ");
    const vLines = cut > 0
      ? [{ t: txt.slice(0, cut), size: 11.5, bold: true, color: r.strong ? c : PT.text },
         { t: txt.slice(cut + 3), size: 9.5, color: PT.faint }]
      : [{ t: txt, size: 11.5, bold: true, color: r.strong ? c : PT.text }];
    out += ppText(trackX + trackW + 12, cy - 8, valW, barH + 16, vLines, { anchor: "ctr" });
  });
  if (v.mark != null && v.mark > 0) {
    const mx = trackX + (v.mark / max) * trackW;
    out += ppText(mx - 60, top + 4, 120, 16, [{ t: v.markLabel || `목표 ${v.mark}`, size: 9, color: PT.dim, align: "ctr" }]);
  }
  if (hasBanner) out += ppBanner(586, s.banner.kind, s.banner.label, s.banner.text);
  return out;
}

// 3) 추이 — 꺾은선 + 면. 값은 점 위에 직접 적어서 축 눈금을 지웁니다.
function bodyTrend(v, s) {
  const rows = (v.rows || []).filter((r) => r.value != null);
  if (rows.length < 2) return ppText(64, 300, 1152, 40, [{ t: "추이를 그릴 만큼 데이터가 쌓이지 않았습니다.", size: 13, color: PT.faint, align: "ctr" }]);
  const hasBanner = !!s.banner;
  const top = 206, boxH = hasBanner ? 356 : 420;
  let out = ppRect(64, top, 1152, boxH, PT.card, { geom: "roundRect", alpha: 0.55 });
  const padL = 56, padR = 56, padT = 62, padB = 52;
  const x0 = 64 + padL, y0 = top + padT;
  const w = 1152 - padL - padR, h = boxH - padT - padB;
  const vals = rows.map((r) => Number(r.value) || 0);
  const lo = Math.min(...vals), hi = Math.max(...vals);
  const span = Math.max(1e-6, hi - lo);
  const pad = span * 0.28;
  const yOf = (val) => y0 + h - ((val - (lo - pad)) / (span + pad * 2)) * h;
  const xOf = (i) => x0 + (rows.length === 1 ? w / 2 : (w * i) / (rows.length - 1));
  // 은은한 가로 눈금
  [0.25, 0.5, 0.75].forEach((g) => { out += ppRect(x0, y0 + h * g, w, 0.8, PT.line, { alpha: 0.5 }); });
  const pts = rows.map((r, i) => [xOf(i), yOf(Number(r.value) || 0)]);
  const color = v.color || PT.green;
  out += ppLine(pts.concat([[xOf(rows.length - 1), y0 + h], [x0, y0 + h]]), color, 0.5, { fill: color, fillAlpha: 0.16, close: true });
  out += ppLine(pts, color, 3);
  rows.forEach((r, i) => {
    const [cx, cy] = pts[i];
    out += ppShape(cx - 7, cy - 7, 14, 14, PT.bg, { geom: "ellipse", line: color, lineW: 28000 });
    out += ppText(cx - 60, cy - 34, 120, 20, [{ t: r.text != null ? r.text : String(r.value), size: 11.5, bold: true, color: PT.text, align: "ctr" }]);
    out += ppText(cx - 60, y0 + h + 14, 120, 20, [{ t: r.label, size: 10, color: PT.dim, align: "ctr" }]);
  });
  if (hasBanner) out += ppBanner(586, s.banner.kind, s.banner.label, s.banner.text);
  return out;
}

// 4) 현안 과제 — 카드 3장. 아이콘·우선순위 뱃지로 읽지 않아도 감이 오게.
function bodyIssues(v) {
  const cards = (v.cards || []).slice(0, 3);
  if (cards.length === 0) return ppText(64, 300, 1152, 40, [{ t: "현안 과제가 없습니다.", size: 13, color: PT.faint, align: "ctr" }]);
  const gap = 22, w = (1152 - gap * (cards.length - 1)) / cards.length;
  let out = "";
  cards.forEach((c, i) => {
    const x = 64 + (w + gap) * i;
    const color = hex(c.color || PT.blue);
    out += ppRect(x, 200, w, 400, PT.card, { geom: "roundRect" });
    out += ppRect(x, 200, w, 4, color);                     // 색은 위쪽 테두리에만
    out += ppIcon(x + 24, 228, 38, c.icon || "dot", color);
    out += ppText(x + 74, 232, w - 96, 34, [{ t: c.title, size: 14.5, bold: true, color: PT.text }]);
    out += ppRect(x + 24, 284, Math.min(w - 48, 8 + c.badge.length * 8.6), 24, color, { geom: "roundRect", alpha: 0.16 });
    out += ppText(x + 24, 284, Math.min(w - 48, 8 + c.badge.length * 8.6), 24,
      [{ t: c.badge, size: 9.5, bold: true, color, align: "ctr" }], { anchor: "ctr" });
    (c.rows || []).slice(0, 4).forEach((r, j) => {
      const y = 326 + j * 64;
      out += ppText(x + 24, y, w - 48, 58, [{ t: [
        { t: `${r.k}  `, size: 11, bold: true, color: PT.text },
        { t: r.v, size: 11, color: PT.dim },
      ], lnSpc: 126000 }]);
    });
  });
  return out;
}

// 5) 실행계획 — 타임라인. 언제까지 누가 무엇을.
function bodyPlan(v) {
  const rows = (v.rows || []).slice(0, 5);
  if (rows.length === 0) return ppText(64, 300, 1152, 40, [{ t: "실행 항목이 없습니다.", size: 13, color: PT.faint, align: "ctr" }]);
  let out = ppRect(120, 232, 2, rows.length * 84 - 40, PT.line); // 세로 축
  rows.forEach((r, i) => {
    const y = 214 + i * 84;
    const color = hex(r.color || PT.blue);
    out += ppShape(112, y + 16, 18, 18, color, { geom: "ellipse" });
    out += ppRect(160, y, 1056, 66, PT.card, { geom: "roundRect", alpha: 0.55 });
    out += ppText(182, y + 12, 700, 24, [{ t: r.title, size: 12.5, bold: true, color: PT.text }]);
    out += ppText(182, y + 36, 700, 20, [{ t: r.detail || "", size: 10.5, color: PT.dim }]);
    out += ppRect(950, y + 20, 96, 26, color, { geom: "roundRect", alpha: 0.16 });
    out += ppText(950, y + 20, 96, 26, [{ t: r.owner || "-", size: 10, bold: true, color, align: "ctr" }], { anchor: "ctr" });
    out += ppText(1060, y + 20, 136, 26, [{ t: r.due || "-", size: 10.5, color: PT.dim, align: "r" }], { anchor: "ctr" });
  });
  return out;
}

// 6) 별첨 — 상세 표는 여기로 몰아둡니다.
function bodyTable(v) {
  const head = v.head || [];
  const rows = (v.rows || []).slice(0, 14);
  const ratio = v.colRatio || head.map((_, i) => (i === 0 ? 0.28 : 0.72 / Math.max(1, head.length - 1)));
  const t = ppTable(64, 208, 1152, [head, ...rows], { colW: ratio.map((r) => r * 1152), rowH: 25 });
  return t + (v.note ? ppText(64, 208 + 25 * (rows.length + 1) + 14, 1152, 20,
    [{ t: v.note, size: 9.5, color: PT.faint }]) : "");
}

// 7) 표지
function bodyCover(v) {
  let out = ppRect(0, 0, 1280, 720, PT.bg);
  out += ppRect(0, 0, 1280, 4, PT.green);
  out += ppText(96, 236, 1088, 40, [{ t: v.eyebrow || "", size: 13, bold: true, color: PT.green, spc: 2 }]);
  out += ppText(96, 282, 1088, 120, [{ t: v.title, size: 40, bold: true, color: PT.text, lnSpc: 116000 }]);
  out += ppRect(96, 424, 64, 3, PT.green);
  out += ppText(96, 452, 1088, 30, [{ t: v.subtitle || "", size: 15, color: PT.dim }]);
  const meta = (v.meta || []).map((m) => `${m}`).join("     ·     ");
  out += ppText(96, 616, 1088, 24, [{ t: meta, size: 11, color: PT.faint }]);
  return out;
}

function slideBody(s) {
  const v = s.visual || {};
  if (v.type === "cover") return bodyCover(v);
  if (v.type === "kpi") return bodyKpi(v, s);
  if (v.type === "bars") return bodyBars(v, s);
  if (v.type === "trend") return bodyTrend(v, s);
  if (v.type === "issues") return bodyIssues(v);
  if (v.type === "plan") return bodyPlan(v);
  if (v.type === "table") return bodyTable(v);
  return s.bullets && s.bullets.length
    ? ppText(64, 210, 1152, 380, s.bullets.map((b, i) => ({ t: b, size: 13, color: PT.dim, bullet: true, spaceBefore: i ? 12 : 0 })))
    : "";
}

function slideXml(s, idx, total) {
  ppShapeId = 1;
  const isCover = s.visual && s.visual.type === "cover";
  const body = isCover ? slideBody(s) : slideFrame(s, idx, total) + slideBody(s);
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>${body}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`;
}


const PPT_THEME = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="DRT"><a:themeElements><a:clrScheme name="DRT"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="0E1420"/></a:dk2><a:lt2><a:srgbClr val="E8EDF5"/></a:lt2><a:accent1><a:srgbClr val="E9692F"/></a:accent1><a:accent2><a:srgbClr val="4C8DFF"/></a:accent2><a:accent3><a:srgbClr val="2FBF71"/></a:accent3><a:accent4><a:srgbClr val="F59E0B"/></a:accent4><a:accent5><a:srgbClr val="A855F7"/></a:accent5><a:accent6><a:srgbClr val="EF4444"/></a:accent6><a:hlink><a:srgbClr val="4C8DFF"/></a:hlink><a:folHlink><a:srgbClr val="A855F7"/></a:folHlink></a:clrScheme><a:fontScheme name="DRT"><a:majorFont><a:latin typeface="Malgun Gothic"/><a:ea typeface="Malgun Gothic"/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Malgun Gothic"/><a:ea typeface="Malgun Gothic"/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="DRT"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>`;

const PPT_LAYOUT = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank" preserve="1"><p:cSld name="빈 화면"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`;

const PPT_MASTER = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="0E1420"/></a:solidFill><a:effectLst/></p:bgPr></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst><p:txStyles><p:titleStyle><a:lvl1pPr><a:defRPr sz="2600"/></a:lvl1pPr></p:titleStyle><p:bodyStyle><a:lvl1pPr><a:defRPr sz="1400"/></a:lvl1pPr></p:bodyStyle><p:otherStyle><a:lvl1pPr><a:defRPr sz="1400"/></a:lvl1pPr></p:otherStyle></p:txStyles></p:sldMaster>`;

function buildPptx(slides) {
  const n = slides.length;
  const files = [];
  const ct = [
    `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>`,
    `<Default Extension="xml" ContentType="application/xml"/>`,
    `<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>`,
    `<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>`,
    `<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>`,
    `<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>`,
    ...slides.map((_, i) => `<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`),
  ].join("");
  files.push({ name: "[Content_Types].xml", content: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">${ct}</Types>` });
  files.push({ name: "_rels/.rels", content: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>` });

  const sldIds = slides.map((_, i) => `<p:sldId id="${256 + i}" r:id="rId${i + 2}"/>`).join("");
  files.push({ name: "ppt/presentation.xml", content: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" saveSubsetFonts="1"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst>${sldIds}</p:sldIdLst><p:sldSz cx="${SLIDE_W}" cy="${SLIDE_H}"/><p:notesSz cx="${SLIDE_H}" cy="${SLIDE_W}"/></p:presentation>` });

  const presRels = [`<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>`]
    .concat(slides.map((_, i) => `<Relationship Id="rId${i + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${i + 1}.xml"/>`))
    .concat([`<Relationship Id="rId${n + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>`]).join("");
  files.push({ name: "ppt/_rels/presentation.xml.rels", content: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${presRels}</Relationships>` });

  files.push({ name: "ppt/slideMasters/slideMaster1.xml", content: PPT_MASTER });
  files.push({ name: "ppt/slideMasters/_rels/slideMaster1.xml.rels", content: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/></Relationships>` });
  files.push({ name: "ppt/slideLayouts/slideLayout1.xml", content: PPT_LAYOUT });
  files.push({ name: "ppt/slideLayouts/_rels/slideLayout1.xml.rels", content: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>` });
  files.push({ name: "ppt/theme/theme1.xml", content: PPT_THEME });

  slides.forEach((s, i) => {
    files.push({ name: `ppt/slides/slide${i + 1}.xml`, content: slideXml(s, i + 1, n) });
    files.push({ name: `ppt/slides/_rels/slide${i + 1}.xml.rels`, content: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>` });
  });
  return makeZip(files);
}

function downloadPptx(filename, slides) {
  downloadFile(filename, buildPptx(slides), "application/vnd.openxmlformats-officedocument.presentationml.presentation");
}

// 노드에서 쓸 때: module.exports = { buildPptx, slideXml, PT, ppText, ppShape, ppBanner };
