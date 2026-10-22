/* 实践支队预算表生成器 —— 纯前端逻辑
 * 口径与 实践部/自动化/*.py 完全一致：
 *   交通 = 行程链逐段二等座最高票价之和（北京→A→B→北京）
 *   住宿 = 标间300 × ceil(人数/2) × (天数-1) × 比例（含"承担住宿"字样则 0）
 *   保险 = 30 × 人数 × ceil(天数/7) × 比例（30元/人/周）
 *   比例 = 承办单位含"就协"/"就业协会" → 1，否则 2/3
 */

const KNOWN_CITIES = ["上海", "西安", "武汉", "合肥", "广州", "深圳", "珠海", "成都",
  "福州", "宁德", "南京", "苏州", "无锡", "杭州", "重庆", "江门",
  "青岛", "沈阳", "大连", "长沙", "郑州", "天津", "济南", "兰州", "南宁"];

/* 云函数实时查价（票价库未收录的分段在线查询 12306，失败回落"待查/手填"） */
const LIVE_API = "https://1386267309-6stucp8l6g.ap-shanghai.tencentscf.com/";

let FARES = null;

/* ---------- 纯函数（可在 node 中测试） ---------- */

function fmtNum(x) {
  if (x === null || x === undefined || x === "") return "";
  const v = Math.round((Number(x) + 1e-9) * 100) / 100;
  return Number.isInteger(v) ? String(v) : String(v);
}

function knownCities(fares) {
  const s = new Set(KNOWN_CITIES);
  for (const key of Object.keys((fares && fares.segments) || {})) {
    for (const ep of key.split("-")) if (ep !== "北京") s.add(ep);
  }
  return [...s].sort((a, b) => b.length - a.length); // 长名优先
}

function parseCities(text, cities) {
  const t = String(text || "").replace(/\s+/g, "");
  const found = [];
  for (const c of cities) {
    if (c === "北京") continue;
    const i = t.indexOf(c);
    if (i >= 0) found.push([i, c]);
  }
  return found.sort((a, b) => a[0] - b[0]).map(x => x[1]);
}

function parseDays(text, today = new Date()) {
  const s = String(text || "");
  let m = s.match(/(\d+)\.(\d+)\s*-\s*(?:(\d+)\.)?(\d+)/) ||
          s.match(/(\d+)月(\d+)日\s*-\s*(?:(\d+)月)?(\d+)日/);
  if (!m) return null;
  const m1 = +m[1], d1 = +m[2], m2 = +(m[3] || m[1]), d2 = +m[4];
  const y = today.getFullYear();
  let start = new Date(y, m1 - 1, d1), end = new Date(y, m2 - 1, d2);
  if (end < start) end = new Date(y + 1, m2 - 1, d2);
  if (end < today) { start = new Date(y + 1, m1 - 1, d1); end = new Date(y + 1, m2 - 1, d2); }
  return Math.round((end - start) / 86400000) + 1;
}

function segBasis(key, price, meta) {
  const m = meta || {};
  let src;
  if (m.manual) src = "人工录入·需复核" + (m.note ? "·" + m.note : "");
  else if (m.live && m.train && m.window) {
    const [w0, w1] = m.window.split("~");
    src = `网页实时查询·${m.train}·窗口${(w0 || "").slice(5)}~${(w1 || "").slice(5)}·最高日${String(m.query_date || "").slice(5)}`;
  }
  else if (m.train && m.window && m.window.includes("~")) {
    const [w0, w1] = m.window.split("~");
    src = `12306实时·${m.train}·窗口${w0.slice(5)}~${w1.slice(5)}·最高日${String(m.query_date || "").slice(5)}`;
  } else if (m.train) src = `12306实时·${m.train}·${m.query_date}查`;
  else src = "缓存·源自2026暑假预算表";
  return `${key} ${fmtNum(price)}元[${src}]`;
}

function fareFor(cities, fares) {
  const segs = (fares && fares.segments) || {};
  const meta = (fares && fares.meta) || {};
  if (!cities.length) return { status: "待查", total: 0, expr: null, basis: "未识别实践地，待人工确认" };
  const chain = ["北京", ...cities, "北京"];
  const parts = [], prices = [];
  for (let i = 0; i < chain.length - 1; i++) {
    const key = `${chain[i]}-${chain[i + 1]}`, rev = `${chain[i + 1]}-${chain[i]}`;
    const k = segs[key] !== undefined ? key : (segs[rev] !== undefined ? rev : null);
    if (k === null) {
      if (cities.length === 1 && fares.roundtrip && fares.roundtrip[cities[0]] !== undefined)
        return { status: "往返总价", total: fares.roundtrip[cities[0]], expr: null,
                 basis: `${cities[0]}往返总价 ${fmtNum(fares.roundtrip[cities[0]])}元[缓存·源自2026暑假预算表]` };
      const us = cities.join("、");
      if (cities.length > 1 && fares.unsplit && fares.unsplit[us] !== undefined)
        return { status: "多城总价", total: fares.unsplit[us], expr: null,
                 basis: `${us}总价 ${fmtNum(fares.unsplit[us])}元[缓存·源自2026暑假预算表，分段待补]` };
      parts.push(`待查：${key}`);
      return { status: "待查", total: prices.reduce((a, b) => a + b, 0),
               expr: null, basis: parts.join("；") };
    }
    prices.push(segs[k]);
    parts.push(segBasis(key, segs[k], meta[k]));
  }
  return { status: "分段", total: prices.reduce((a, b) => a + b, 0),
           expr: prices.map(fmtNum).join("+"), basis: parts.join("；") };
}

function ratioOf(d) {
  const t = String(d || "").replace(/\s+/g, "");
  return (t.includes("就协") || t.includes("就业协会")) ? 1 : 2 / 3;
}

function computeTeam(team, fares, cities) {
  const cs = parseCities(team.B, cities);
  const days = parseDays(team.E);
  const people = Number(team.F) || 12;
  const ac = ratioOf(team.D);
  const acTxt = ac === 1 ? "1" : "2/3";
  const lodgingFree = /承担住宿|住宿自理/.test(
    [team.B, team.C, team.D, team.E, team.H].map(x => String(x || "")).join(""));
  const V = lodgingFree ? 0 : 300, Z = 30;
  const manualT = Number(team.T_manual) > 0 ? Number(team.T_manual) : null;
  const fare = manualT !== null
    ? { status: "人工录入", total: manualT, expr: null,
        basis: `单人交通费 ${fmtNum(manualT)}元[网页人工录入]` }
    : fareFor(cs, fares);
  const T = fare.total;  const nights = days ? days - 1 : 0;
  const weeks = days ? Math.ceil(days / 7) : 1;
  const W = Math.ceil(people / 2);
  const U = people * T * ac, Y = V * W * nights * ac, AA = Z * people * weeks * ac;
  return {
    ...team, cities: cs, days, people, ac, acTxt, V, Z, W, weeks,
    nights, fare, T, U, Y, AA, AB: U + Y + AA,
    L: `${fmtNum(T)}/人x${acTxt}`,
    N: `${fmtNum(V / 2)}/人·天*${acTxt}*${nights}晚`,
    P: `${fmtNum(Z)}/人·周*${weeks}周x${acTxt}`,
    S: ac === 1 ? "就协全部" : "院系1/3，剩余就协",
  };
}

/* ---------- 云函数实时查价（浏览器端） ---------- */

function cityTokens(text, cities) {
  // 先按已知城市表匹配；全未命中时按分隔符切词，作为实时查询的候选城市
  const hits = parseCities(text, cities);
  if (hits.length) return hits;
  return String(text || "").replace(/\s+/g, "")
    .split(/[、，,\/·—]+/)
    .map(s => s.trim())
    .filter(s => /^[一-龥]{2,5}$/.test(s) && !/承担住宿|住宿自理|实践|支队|基地/.test(s));
}

async function liveSegment(from, to) {
  const url = LIVE_API + "?from=" + encodeURIComponent(from) + "&to=" + encodeURIComponent(to) + "&days=14";
  const ctrl = new AbortController();
  const timer = setTimeout(() => ctrl.abort(), 75000);
  try {
    const res = await fetch(url, { signal: ctrl.signal });
    const d = await res.json();
    if (d.error) throw new Error(d.error);
    if (!(d.price > 0)) throw new Error("无直达车次");
    return d;
  } finally { clearTimeout(timer); }
}

async function prefetchLive(valid, onProgress) {
  // 对票价库缺失的分段逐个实时查询，结果写入 FARES（内存，本次会话有效）
  const segs = FARES.segments, meta = FARES.meta || (FARES.meta = {});
  const failed = [];
  for (const t of valid) {
    const toks = cityTokens(t.B, knownCities(FARES));
    if (!toks.length) continue;
    const chain = ["北京", ...toks, "北京"];
    for (let i = 0; i < chain.length - 1; i++) {
      const key = `${chain[i]}-${chain[i + 1]}`, rev = `${chain[i + 1]}-${chain[i]}`;
      if (segs[key] !== undefined || segs[rev] !== undefined) continue;
      if (meta[key] && meta[key].live_failed) continue;
      onProgress && onProgress(`⏳ 正在实时查询 ${key} 票价（约10~30秒）…`);
      try {
        const r = await liveSegment(chain[i], chain[i + 1]);
        segs[key] = r.price;
        meta[key] = { train: r.train, query_date: r.query_date, window: r.window,
                      days_queried: r.days_queried, live: true };
      } catch (e) {
        meta[key] = { live_failed: true, note: String((e && e.message) || e) };
        failed.push(key);
      }
    }
  }
  return failed;
}

/* ---------- xlsx 生成（复刻模板布局，公式+缓存值双写） ---------- */

const HEADERS6 = ["序号", "实践地点", "拟主要参访点位", "承办院系", "参访日期", "支队人数", "", "院系已有支持金额",
  "实践地点", "参访天数", "支队人数", "往返交通费", "合计", "住宿", "合计", "保险", "合计", "总计", "备注",
  "单人交通费", "交通费总和", "住宿费标间", "房间数", "住宿天数", "住宿费用总和",
  "保险费每人", "保险费总额", "队伍总额", "报销比例", "交通费取价依据"];

const R2 = { A: "序号", B: "实践地点", C: "拟主要参访点位", D: "承办单位", E: "拟参访日期",
  F: "支队人数", H: "院系已有支持金额", I: "实践地点", J: "拟参访天数", K: "支队人数",
  L: "往返交通费", N: "住宿", P: "保险", R: "总计",
  T: "单人交通费", U: "交通费总和", V: "住宿费标间", W: "房间数", X: "住宿天数",
  Y: "住宿费用总和", Z: "保险费每人", AA: "保险费总额", AB: "队伍总额", AD: "交通费取价依据" };

const NOTES3 = { A: "填写说明", D: "承办学院（就协组织填“就协”）", E: "（格式：7.19-7.23）",
  F: "默认12，可按实际调整", T: "北京-实践地各段高铁二等座最高价之和（多城市逐段相加）",
  AD: "逐段：票价[12306实时·车次·查询日期]；缓存价标注来源",
  V: "每人每天住宿报销上限150（标间300/2人）", Z: "保险费每人每周不超过30元（≤7天即30元/人）" };

const COL_W = { A: 6, B: 12, C: 40, D: 16, E: 12, F: 9, G: 3, H: 13, I: 12, J: 9, K: 9,
  L: 14, M: 10, N: 18, O: 10, P: 12, Q: 10, R: 11, S: 26,
  T: 11, U: 11, V: 11, W: 11, X: 11, Y: 11, Z: 11, AA: 11, AB: 11, AC: 11, AD: 70 };

function colLetter(i) { // 1-based
  let s = "";
  while (i > 0) { const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = (i - 1 - m) / 26; }
  return s;
}

function daysFormula(r) {
  const e = `E${r}`;
  return `IFERROR(DATE(YEAR(TODAY()),MID(${e},FIND("-",${e})+1,FIND(".",${e},FIND("-",${e}))-FIND("-",${e})-1),` +
    `RIGHT(${e},LEN(${e})-FIND(".",${e},FIND("-",${e}))))` +
    `-DATE(YEAR(TODAY()),LEFT(${e},FIND(".",${e})-1),MID(${e},FIND(".",${e})+1,FIND("-",${e})-FIND(".",${e})-1))+1,"待填")`;
}

function acFormula(r) {
  return `IF(OR(ISNUMBER(SEARCH("就协",D${r})),ISNUMBER(SEARCH("就业协会",D${r}))),1,2/3)`;
}

function setF(ws, addr, f, v, isStr) {
  ws[addr] = isStr ? { t: "str", f, v } : { t: "n", f, v };
}

function buildWorkbook(computed) {
  const ws = {};
  const startRow = 7;
  // 行1 区域标题
  ws["A1"] = { t: "s", v: "院系提供" };
  ws["I1"] = { t: "s", v: "主要呈现列（公式，勿手改）" };
  ws["T1"] = { t: "s", v: "计算辅助列" };
  ws["AD1"] = { t: "s", v: "取价依据" };
  // 行2/3 小表头与说明
  for (const [c, v] of Object.entries(R2)) ws[`${c}2`] = { t: "s", v };
  for (const [c, v] of Object.entries(NOTES3)) ws[`${c}3`] = { t: "s", v };
  // 行6 正式表头
  HEADERS6.forEach((h, i) => { if (h) ws[`${colLetter(i + 1)}6`] = { t: "s", v: h }; });

  computed.forEach((t, idx) => {
    const r = startRow + idx;
    ws[`A${r}`] = { t: "n", v: idx + 1 };
    for (const c of ["B", "C", "D", "E"]) ws[`${c}${r}`] = { t: "s", v: String(t[c] || "") };
    ws[`F${r}`] = { t: "n", v: t.people };
    if (t.H !== undefined && t.H !== null && t.H !== "") ws[`H${r}`] = { t: "n", v: Number(t.H) || 0 };
    setF(ws, `I${r}`, `B${r}`, String(t.B || ""), true);
    setF(ws, `J${r}`, daysFormula(r), t.days || "待填", !t.days);
    setF(ws, `K${r}`, `F${r}`, t.people);
    setF(ws, `L${r}`, `T${r}&"/人"&"x"&TEXT(AC${r},"# ?/?")`, t.L, true);
    setF(ws, `M${r}`, `U${r}`, t.U);
    setF(ws, `N${r}`, `V${r}/2&"/人·天*"&TRIM(TEXT(AC${r},"# ?/?"))&"*"&X${r}&"晚"`, t.N, true);
    setF(ws, `O${r}`, `Y${r}`, t.Y);
    setF(ws, `P${r}`, `Z${r}&"/人·周*"&CEILING(J${r}/7,1)&"周"&"x"&TEXT(AC${r},"# ?/?")`, t.P, true);
    setF(ws, `Q${r}`, `AA${r}`, t.AA);
    setF(ws, `R${r}`, `AB${r}`, t.AB);
    setF(ws, `S${r}`, `IF(OR(ISNUMBER(SEARCH("就协",D${r})),ISNUMBER(SEARCH("就业协会",D${r}))),"就协全部","院系"&TEXT(1-AC${r},"# ?/?")&"，剩余就协")`, t.S, true);
    if (t.fare.status === "待查") ws[`T${r}`] = { t: "s", v: "待查" };
    else if (t.fare.expr) setF(ws, `T${r}`, t.fare.expr, t.T);
    else ws[`T${r}`] = { t: "n", v: t.T };
    setF(ws, `U${r}`, `K${r}*T${r}*AC${r}`, t.U);
    ws[`V${r}`] = { t: "n", v: t.V };
    setF(ws, `W${r}`, `CEILING(K${r}/2,1)`, t.W);
    setF(ws, `X${r}`, `IF(ISNUMBER(J${r}),J${r}-1,"")`, t.nights);
    setF(ws, `Y${r}`, `V${r}*W${r}*X${r}*AC${r}`, t.Y);
    ws[`Z${r}`] = { t: "n", v: t.Z };
    setF(ws, `AA${r}`, `Z${r}*K${r}*CEILING(J${r}/7,1)*AC${r}`, t.AA);
    setF(ws, `AB${r}`, `U${r}+Y${r}+AA${r}`, t.AB);
    setF(ws, `AC${r}`, acFormula(r), t.ac);
    ws[`AD${r}`] = { t: "s", v: t.fare.basis };
  });

  // 对账区
  const last = startRow + computed.length - 1, chk = last + 2;
  const total = computed.reduce((a, t) => a + t.AB, 0);
  ws[`Q${chk}`] = { t: "s", v: "过会金额" };
  ws[`Q${chk + 1}`] = { t: "s", v: "预算合计" };
  setF(ws, `R${chk + 1}`, `SUM(R${startRow}:R${last})`, total);
  ws[`Q${chk + 2}`] = { t: "s", v: "差额" };
  setF(ws, `R${chk + 2}`, `R${chk}-R${chk + 1}`, -total);

  ws["!ref"] = `A1:AD${chk + 2}`;
  ws["!merges"] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: 7 } },
    { s: { r: 0, c: 8 }, e: { r: 0, c: 18 } },
    { s: { r: 0, c: 19 }, e: { r: 0, c: 28 } },
  ];
  const cols = [];
  for (let i = 1; i <= 30; i++) cols.push({ wch: COL_W[colLetter(i)] || 11 });
  ws["!cols"] = cols;
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "各支队预算按实际人数校准");
  return wb;
}

/* ---------- 公文版 docx 生成（复刻 gen_budget_docx.py 排版） ---------- */

function docxDescs(t) {
  // 与 Python 端 gen_budget_docx 的描述列口径一致
  const rs = t.ac === 1 ? "" : " x 2/3";
  return {
    t_desc: t.fare.status === "待查" ? "待查" : `${fmtNum(t.T)}/人${rs}`,
    l_desc: t.V === 0 ? "实践地承担"
      : (t.ac === 1 ? `${fmtNum(t.V / 2)}/人·天*${t.nights}晚`
                    : `${fmtNum(t.V / 2)}/人·天*2/3*${t.nights}晚`),
    i_desc: t.weeks > 1 ? `${fmtNum(t.Z)}/人·周*${t.weeks}周${rs}` : `${fmtNum(t.Z)}/人${rs}`,
    remark: (t.ac === 1 ? "就协全部" : "院系 1/3，剩余就协")
      + (t.V === 0 ? "；部分经费由地方支持" : ""),
  };
}

function buildDocx(computed, title, season) {
  const D = globalThis.docx;
  if (!D) throw new Error("docx 库未加载（需联网加载 CDN 脚本）");
  const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
          AlignmentType, WidthType, BorderStyle, VerticalAlign, TableLayoutType } = D;

  const HP = pt => pt * 2;                       // half-points
  const CM = cm => Math.round(cm * 567);         // twips
  const F_BODY = { ascii: "Times New Roman", eastAsia: "仿宋_GB2312" };
  const F_CELL = { ascii: "Times New Roman", eastAsia: "宋体" };
  const F_TITLE = { ascii: "Times New Roman", eastAsia: "方正大标宋简体" };

  const para = (text, font, pt, opts = {}) => new Paragraph({
    alignment: opts.center ? AlignmentType.CENTER : AlignmentType.LEFT,
    spacing: { after: 120 },
    children: [new TextRun({ text, font, size: HP(pt), bold: !!opts.bold })],
  });

  const cell = (text, opts = {}) => new TableCell({
    rowSpan: opts.rowSpan, columnSpan: opts.colSpan,
    verticalAlign: VerticalAlign.CENTER,
    width: opts.w ? { size: CM(opts.w), type: WidthType.DXA } : undefined,
    children: String(text).split("\n").map(line => new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: line, font: F_CELL, size: HP(opts.pt || 9), bold: !!opts.bold })],
    })),
  });

  const border = { style: BorderStyle.SINGLE, size: 2, color: "000000" };
  const borders = { top: border, bottom: border, left: border, right: border,
                    insideHorizontal: border, insideVertical: border };
  const widths = [0.9, 3.0, 0.9, 2.2, 1.3, 2.5, 1.3, 1.8, 1.2, 1.3, 1.6]; // cm，合计18cm适配竖版版心

  const teams = computed.map(t => ({ ...t, ...docxDescs(t) }));
  const total = teams.reduce((a, t) => a + t.AB, 0);

  // 表头三行
  const r0 = new TableRow({ children: [
    cell("企业实习实践支队", { colSpan: 11, pt: 11, bold: true })] });
  const r1 = new TableRow({ children: [
    cell("序号", { rowSpan: 2, bold: true, w: widths[0] }),
    cell("支队目的地", { rowSpan: 2, bold: true, w: widths[1] }),
    cell("人数", { rowSpan: 2, bold: true, w: widths[2] }),
    cell("交通", { colSpan: 2, bold: true }),
    cell("住宿", { colSpan: 2, bold: true }),
    cell("保险", { colSpan: 2, bold: true }),
    cell("总计", { rowSpan: 2, bold: true, w: widths[9] }),
    cell("备注", { rowSpan: 2, bold: true, w: widths[10] })] });
  const r2 = new TableRow({ children: [
    cell("单人往返费用", { bold: true }), cell("合计", { bold: true }),
    cell("单间费用", { bold: true }), cell("合计", { bold: true }),
    cell("单人费用", { bold: true }), cell("合计", { bold: true })] });

  const dataRows = teams.map((t, i) => new TableRow({ children: [
    cell(i + 1), cell(`${t.B}\n${t.D}`), cell(t.people),
    cell(t.t_desc), cell(fmtNum(t.U)), cell(t.l_desc), cell(fmtNum(t.Y)),
    cell(t.i_desc), cell(fmtNum(t.AA)), cell(fmtNum(t.AB)), cell(t.remark)] }));

  const totalRow = new TableRow({ children: [
    cell("合计", { colSpan: 4, bold: true }), cell(fmtNum(total), { colSpan: 7, bold: true })] });

  const table = new Table({
    layout: TableLayoutType.FIXED,
    width: { size: widths.map(CM).reduce((a, b) => a + b, 0), type: WidthType.DXA },
    columnWidths: widths.map(CM),
    borders,
    rows: [r0, r1, r2, ...dataRows, totalRow],
  });

  const intro = `职业发展中心于${season}组织【  】支基层实习实践支队、${teams.length}支企业主题实践支队。` +
    `其中，基层实习实践支队实践预算为【  】元（基研会部分，另行填写）；` +
    `企业主题实践支队${teams.length}支，预算共${fmtNum(total)}元。实践预算总额为【  】元，详情如下：`;

  const note1 = "注：（1）上述预算中交通费用、保险费用为0的支队是表示由实习（实践）当地部门负责承担或同学自行承担交通费用和保险费用。" +
    "交通费用一般按照高铁（二等座）来计算，更适合飞机出行的基地（路程远、机票费用低于高铁二等座）则按照机票费用来计算；" +
    "保险费用按照平安保险-短期综合意外险费用计算。";
  const note2 = "（2）与院系联合组织支队按照支持上限为“北京往返实践地交通费用+当地住宿费用总和2/3”的标准执行，不含保险、队旗及其他费用。";

  return new Document({
    sections: [{
      properties: { page: {
        size: { width: CM(21), height: CM(29.7) },   // A4 竖版
        margin: { left: CM(1.5), right: CM(1.5), top: CM(1.5), bottom: CM(1.5) },
      } },
      children: [
        para(title, F_TITLE, 18, { center: true, bold: true }),
        para(intro, F_BODY, 12),
        table,
        para("", F_BODY, 12),
        para(note1, F_BODY, 10.5),
        para(note2, F_BODY, 10.5),
      ],
    }],
  });
}

function downloadBlob(blob, filename) {
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 5000);
}

/* node 测试导出 */
if (typeof module !== "undefined") {
  module.exports = { fmtNum, knownCities, parseCities, parseDays, fareFor, ratioOf, computeTeam, buildWorkbook, docxDescs, buildDocx };
}

/* ---------- 以下为浏览器 UI ---------- */
if (typeof document !== "undefined") {
  const $ = id => document.getElementById(id);
  let teams = [];
  let lastComputed = null;   // 预览通过后才允许下载；任何编辑都会使其失效

  function invalidate() {
    lastComputed = null;
    $("generate").disabled = true;
    $("generate-docx").disabled = true;
    $("preview").style.display = "none";
    $("warn").style.display = "none";
  }

  function saveLocal() { try { localStorage.setItem("budgetTeams", JSON.stringify(teams)); } catch (e) {} }
  function loadLocal() {
    try { const s = localStorage.getItem("budgetTeams"); if (s) teams = JSON.parse(s); } catch (e) {}
    if (!teams.length) teams = [{ B: "", C: "", D: "就协", E: "", F: 12, H: "" }];
  }

  function renderEditor() {
    const tb = document.querySelector("#editor tbody");
    tb.innerHTML = "";
    teams.forEach((t, i) => {
      const tr = document.createElement("tr");
      tr.innerHTML = `<td>${i + 1}</td>` +
        ["B", "C", "D", "E", "F", "T_manual", "H"].map(c =>
          `<td><input data-i="${i}" data-c="${c}" value="${String(t[c] ?? "").replace(/"/g, "&quot;")}" ` +
          (c === "B" ? `list="city-list" ` : "") +
          `placeholder="${{ B: "可自由输入，如 合肥、上海", C: "企业全称…", D: "就协", E: "7.19-7.23", F: "12", T_manual: "自动", H: "0" }[c]}"></td>`).join("") +
        `<td><button data-del="${i}" style="border:none;background:none;cursor:pointer;color:#c00">✕</button></td>`;
      tb.appendChild(tr);
    });
    tb.querySelectorAll("input").forEach(inp => inp.addEventListener("input", e => {
      const i = +e.target.dataset.i, c = e.target.dataset.c;
      teams[i][c] = (c === "F" || c === "T_manual") ? (Number(e.target.value) || "") : e.target.value;
      if (c === "E") {  // 日期即时校验：可解析变绿、不可解析变红、空白还原
        const v = e.target.value.trim();
        e.target.className = !v ? "" : (parseDays(v) ? "good" : "bad");
        e.target.title = e.target.className === "bad" ? "无法识别，格式应为 7.19-7.23" : "";
      }
      if (c === "B" && FARES) {  // 地点：命中票价库变绿；库外城市不标红（可手填交通费），仅气泡提示
        const v = e.target.value.trim();
        const hit = !v || parseCities(v, knownCities(FARES)).length > 0;
        e.target.className = (v && hit) ? "good" : "";
        e.target.title = (v && !hit) ? "票价库未收录该地点：可照常在右侧「单人交通费」列手填票价" : "";
      }
      saveLocal();
      invalidate();
    }));
    tb.querySelectorAll("[data-del]").forEach(b => b.addEventListener("click", e => {
      teams.splice(+e.target.dataset.del, 1);
      if (!teams.length) teams = [{ B: "", C: "", D: "就协", E: "", F: 12, H: "" }];
      renderEditor(); saveLocal(); invalidate();
    }));
  }

  function collectValid() {
    return teams.filter(t => String(t.B || "").trim() && String(t.E || "").trim());
  }

  function showWarn(lines) {
    const w = $("warn");
    if (!lines.length) { w.style.display = "none"; return; }
    w.style.display = "block";
    w.textContent = "⚠ 注意：\n" + lines.join("\n");
  }

  function showPreview(computed) {
    const p = $("preview");
    p.style.display = "table";
    p.innerHTML = "<tr><th>实践地点</th><th>承办单位</th><th>人数</th><th>天数</th><th>比例</th>" +
      "<th>单人交通费</th><th>预算总额</th><th>票价状态</th></tr>" +
      computed.map(t => `<tr><td>${t.B}</td><td>${t.D}</td><td>${t.people}</td>` +
        `<td>${t.days ?? "?"}</td><td>${t.acTxt}</td><td>${fmtNum(t.T)}</td>` +
        `<td>${fmtNum(t.AB)}</td><td>${t.fare.status === "待查" ? "⚠ 待查" : "✓ " + t.fare.status}</td></tr>`).join("") +
      `<tr><th colspan="6">合计</th><th>${fmtNum(computed.reduce((a, t) => a + t.AB, 0))}</th><th></th></tr>`;
  }

  async function computeAll() {
    const valid = collectValid();
    if (!valid.length) { alert("请至少填写一支支队（实践地点 + 拟参访日期必填）"); return null; }
    const btn = $("preview-btn");
    btn.disabled = true;
    const setStatus = msg => { $("fare-info").textContent = msg; };
    let liveFailed = [];
    try {
      liveFailed = await prefetchLive(valid, setStatus);
    } finally {
      btn.disabled = false;
      const n = Object.keys(FARES.segments || {}).length;
      const updated = (FARES.meta && FARES.meta.updated_at) || "未知";
      $("fare-info").textContent = `票价库：${n} 个分段 · 更新于 ${updated} · 取 12306 预售期 14 天窗口二等座最高价`;
    }
    const cities = knownCities(FARES);
    const computed = valid.map(t => computeTeam(t, FARES, cities));
    const warns = [];
    computed.forEach(t => {
      if (t.fare.status === "待查") warns.push(`「${t.B}」${t.fare.basis}——可在网页"单人交通费"列直接手填后重新生成`);
      if (!t.days) warns.push(`「${t.B}」日期「${t.E}」无法识别，天数/住宿/保险未算`);
      if (t.fare.status === "往返总价" || t.fare.status === "多城总价")
        warns.push(`「${t.B}」使用${t.fare.status}（非实时分段价）`);
    });
    if (liveFailed.length)
      warns.push(`以下分段实时查询失败或无直达车：${liveFailed.join("、")}——请手填"单人交通费"，或联系实践部用 --set 人工录入`);
    showWarn(warns);
    showPreview(computed);
    lastComputed = computed;
    $("generate").disabled = false;
    $("generate-docx").disabled = false;
    return computed;
  }

  function generate(fmt) {
    const computed = lastComputed;
    if (!computed) { alert("请先点击「生成预览」核对数据"); return; }
    const d = new Date(), pad = n => String(n).padStart(2, "0");
    const stamp = `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}`;
    if (fmt === "xlsx") {
      const wb = buildWorkbook(computed);
      XLSX.writeFile(wb, `实践预算_${stamp}.xlsx`);
    } else {
      const title = $("doc-title").value.trim() || "实践预算";
      const season = $("doc-season").value.trim() || "本期";
      docx.Packer.toBlob(buildDocx(computed, title, season))
        .then(blob => downloadBlob(blob, `${title}.docx`));
    }
  }

  /* 上传解析：找 B 列非空且 F 列为数字的行 */
  async function handleFile(file) {
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf);
    const name = wb.SheetNames.includes("各支队预算按实际人数校准") ? "各支队预算按实际人数校准" : wb.SheetNames[0];
    const ws = wb.Sheets[name];
    const range = XLSX.utils.decode_range(ws["!ref"]);
    const found = [];
    for (let r = range.s.r; r <= range.e.r; r++) {
      const g = c => { const cell = ws[`${c}${r + 1}`]; return cell ? cell.v : undefined; };
      const B = g("B"), F = g("F");
      if (B && typeof B === "string" && B.trim() && typeof F === "number") {
        found.push({ B: String(B).trim(), C: g("C") || "", D: g("D") || "", E: g("E") || "", F, H: g("H") ?? "" });
      }
    }
    if (!found.length) { $("upload-info").textContent = "未识别到支队行（需 B 列实践地点 + F 列人数为数字）"; return; }
    teams = found;
    saveLocal(); renderEditor(); invalidate();
    $("upload-info").textContent = `已导入 ${found.length} 支支队（表：${name}），可切换到「手动填写」查看修改`;
    $("tab-manual").click();
  }

  document.addEventListener("DOMContentLoaded", async () => {
    loadLocal(); renderEditor();
    $("tab-manual").addEventListener("click", () => {
      $("tab-manual").classList.add("active"); $("tab-upload").classList.remove("active");
      $("panel-manual").style.display = ""; $("panel-upload").style.display = "none";
    });
    $("tab-upload").addEventListener("click", () => {
      $("tab-upload").classList.add("active"); $("tab-manual").classList.remove("active");
      $("panel-upload").style.display = ""; $("panel-manual").style.display = "none";
    });
    $("add-row").addEventListener("click", () => {
      teams.push({ B: "", C: "", D: "就协", E: "", F: 12, H: "" });
      renderEditor(); saveLocal(); invalidate();
    });
    $("preview-btn").addEventListener("click", computeAll);
    $("generate").addEventListener("click", () => generate("xlsx"));
    $("generate-docx").addEventListener("click", () => generate("docx"));
    $("drop").addEventListener("click", () => $("file").click());
    $("drop").addEventListener("dragover", e => { e.preventDefault(); $("drop").classList.add("hover"); });
    $("drop").addEventListener("dragleave", () => $("drop").classList.remove("hover"));
    $("drop").addEventListener("drop", e => {
      e.preventDefault(); $("drop").classList.remove("hover");
      if (e.dataTransfer.files[0]) handleFile(e.dataTransfer.files[0]);
    });
    $("file").addEventListener("change", e => { if (e.target.files[0]) handleFile(e.target.files[0]); });

    try {
      const res = await fetch("fares.json");
      FARES = await res.json();
      const n = Object.keys(FARES.segments || {}).length;
      const updated = (FARES.meta && FARES.meta.updated_at) || "未知";
      $("fare-info").textContent = `票价库：${n} 个分段 · 更新于 ${updated} · 取 12306 预售期 14 天窗口二等座最高价`;
      $("preview-btn").disabled = false;
      // 地点自动联想候选（票价库全部城市）
      $("city-list").innerHTML = knownCities(FARES).map(c => `<option value="${c}">`).join("");
    } catch (e) {
      $("fare-info").textContent = "⚠ fares.json 加载失败（请通过 http(s) 访问本页，file:// 直开无法读取票价库）";
      $("preview-btn").disabled = true;
    }
  });
}
