/* 实践支队预算表生成器 —— 纯前端逻辑
 * 口径与 实践部/自动化/*.py 完全一致：
 *   交通 = 行程链逐段二等座最高票价之和（北京→A→B→北京）
 *   住宿 = 150元/人·晚 × 人数 × (天数-1) × 比例（含"承担住宿"字样则 0；不再按房间数向上取整）
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
  const U = people * T * ac, Y = (V / 2) * people * nights * ac, AA = Z * people * weeks * ac;
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

const LIVE_CACHE_KEY = "fareLiveCacheV1";

function todayStr(d) {
  d = d || new Date();
  return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0") + "-" + String(d.getDate()).padStart(2, "0");
}

function readLiveCache() {
  try {
    if (typeof localStorage === "undefined") return { date: todayStr(), items: {} };
    const c = JSON.parse(localStorage.getItem(LIVE_CACHE_KEY) || "{}");
    if (c && c.date === todayStr() && c.items) return c;
  } catch (e) { /* ignore */ }
  return { date: todayStr(), items: {} };
}

function writeLiveCache(cache) {
  try {
    if (typeof localStorage !== "undefined") localStorage.setItem(LIVE_CACHE_KEY, JSON.stringify(cache));
  } catch (e) { /* quota/private mode */ }
}

function cacheEntry(cache, key) {
  const rev = key.split("-").reverse().join("-");
  return (cache.items && (cache.items[key] || cache.items[rev])) || null;
}

async function prefetchLive(valid, onProgress) {
  // 实时优先：链上所有分段都调云函数查最新价（3路并发）；失败的分段回落票价库快照（无快照才标失败）
  const segs = FARES.segments, meta = FARES.meta || (FARES.meta = {});
  const liveWanted = (typeof document === "undefined") ? true
    : (document.getElementById("live-first") && document.getElementById("live-first").checked);
  const today = todayStr();
  const cache = readLiveCache();
  const tasks = [], seen = new Set();
  for (const t of valid) {
    const toks = cityTokens(t.B, knownCities(FARES));
    if (!toks.length) continue;
    const chain = ["北京", ...toks, "北京"];
    for (let i = 0; i < chain.length - 1; i++) {
      const key = `${chain[i]}-${chain[i + 1]}`, rev = `${chain[i + 1]}-${chain[i]}`;
      if (seen.has(key) || seen.has(rev)) continue;
      seen.add(key);
      const cached = cacheEntry(cache, key);
      if (cached && cached.price > 0) {
        segs[key] = cached.price;
        meta[key] = { ...cached, live: true, cached: true };
        continue;
      }
      if (meta[key] && meta[key].live && meta[key].query_date === today) continue;   // 本次会话当天已查
      if (meta[key] && meta[key].query_date === today && segs[key] !== undefined) {  // 快照当天已查
        meta[key] = { ...meta[key], cached: true };
        continue;
      }
      if (meta[key] && meta[key].live_failed && meta[key].failed_date === today) continue;
      const hasSnap = segs[key] !== undefined || segs[rev] !== undefined;
      if (!liveWanted && hasSnap) continue;               // 快照模式：只补库外分段
      tasks.push([chain[i], chain[i + 1], key, hasSnap]);
    }
  }
  const failed = [], fellBack = [];
  let done = 0;
  const runOne = async ([a, b, key, hasSnap]) => {
    onProgress && onProgress(`⏳ 实时查询票价（${++done}/${tasks.length}）${key}…`);
    try {
      const r = await liveSegment(a, b);
      const entry = { price: r.price, train: r.train, query_date: r.query_date || today,
                      window: r.window, days_queried: r.days_queried, at: Date.now() };
      segs[key] = r.price;
      meta[key] = { ...entry, live: true };
      cache.items[key] = entry;
      writeLiveCache(cache);
    } catch (e) {
      if (hasSnap) fellBack.push(key);                    // 静默回落快照
      else {
        meta[key] = { live_failed: true, failed_date: today, note: String((e && e.message) || e) };
        failed.push(key);
      }
    }
  };
  const queue = [...tasks];
  const workers = Array.from({ length: Math.min(3, queue.length) }, async () => {
    while (queue.length) await runOne(queue.shift());
  });
  await Promise.all(workers);
  return { failed, fellBack };
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
    setF(ws, `Y${r}`, `V${r}/2*K${r}*X${r}*AC${r}`, t.Y);
    ws[`Z${r}`] = { t: "n", v: t.Z };
    setF(ws, `AA${r}`, `Z${r}*K${r}*CEILING(J${r}/7,1)*AC${r}`, t.AA);
    setF(ws, `AB${r}`, `U${r}+Y${r}+AA${r}`, t.AB);
    setF(ws, `AC${r}`, acFormula(r), t.ac);
    ws[`AD${r}`] = { t: "s", v: t.fare.basis };
  });

  // 注：对账区（过会金额/差额）仅内部 rebuild_budget.py 生成，网页版面向院系不附带
  const last = startRow + computed.length - 1;
  const totalRow = last + 1;

  // 合计行：30 列模板原来没有，补上后金山/网页都能直接看到期次合计
  ws[`A${totalRow}`] = { t: "s", v: "合计" };
  const sumDefs = [
    ["F", "people"], ["H", "H"], ["K", "people"], ["M", "U"], ["O", "Y"],
    ["Q", "AA"], ["R", "AB"], ["U", "U"], ["W", "W"], ["Y", "Y"],
    ["AA", "AA"], ["AB", "AB"]
  ];
  for (const [col, field] of sumDefs) {
    const sum = computed.reduce((a, t) => a + (Number(t[field]) || 0), 0);
    setF(ws, `${col}${totalRow}`, `SUM(${col}${startRow}:${col}${last})`, sum);
  }

  // 模板配色（与 预算表模板.xlsx 一致）：院系提供区橙、呈现列蓝、辅助列绿、T列输入高亮黄
  // 数据区全部自动换行+垂直居中，避免列宽不够文字挤在一起
  const FILL = { orange: "FFFFF2CC", blue: "FFD9E1F2", green: "FFE2EFDA", yellow: "FFFFE699" };
  const fillOf = (ci, r) => {
    if (r === 1 || r === 6 || r === totalRow) return ci <= 8 ? FILL.orange : ci <= 19 ? FILL.blue : FILL.green;
    if (r >= startRow && r <= last) {
      if (ci <= 8) return FILL.orange;
      if (ci === 20) return FILL.yellow;
    }
    return null;
  };
  for (let r = 1; r <= totalRow; r++) {
    for (let ci = 1; ci <= 30; ci++) {
      const rgb = fillOf(ci, r);
      const isHeader = r === 1 || r === 6, isData = r >= startRow && r <= last;
      const isTotal = r === totalRow;
      if (!rgb && !isHeader && !isData && !isTotal) continue;
      const addr = `${colLetter(ci)}${r}`;
      if (!ws[addr]) ws[addr] = { t: "s", v: "" };
      const s = ws[addr].s || (ws[addr].s = {});
      if (rgb) s.fill = { patternType: "solid", fgColor: { rgb } };
      if (isHeader || isTotal) {
        s.font = { bold: true };
        s.alignment = { horizontal: "center", vertical: "center", wrapText: true };
      } else if (isData) {
        s.alignment = { vertical: "center", wrapText: true };
      }
    }
  }

  ws["!ref"] = `A1:AD${totalRow}`;
  ws["!merges"] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: 7 } },
    { s: { r: 0, c: 8 }, e: { r: 0, c: 18 } },
    { s: { r: 0, c: 19 }, e: { r: 0, c: 28 } },
    { s: { r: totalRow - 1, c: 0 }, e: { r: totalRow - 1, c: 4 } },
  ];
  const cols = [];
  for (let i = 1; i <= 30; i++) cols.push({ wch: COL_W[colLetter(i)] || 11 });
  ws["!cols"] = cols;
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "各支队预算按实际人数校准");
  return wb;
}

/* ---------- 公文版 docx 生成（复刻 gen_budget_docx.py 排版） ---------- */

function splitCompanyNames(C) {
  return String(C || "").split(/[、,，;；\n]+/)
    .map(function (s) {
      return String(s || "")
        .replace(/[（(]\s*(承担住宿|住宿自理|住宿)\s*[）)]/g, "")
        .replace(/承担住宿|住宿自理/g, "")
        .trim();
    })
    .filter(Boolean);
}

function wordCompanyName(c, fallback) {
  let n = c ? (c.canonical_name || c.short_name || c.raw || fallback) : fallback;
  n = String(n || "").trim();
  if (c && n.indexOf("·") < 0) {
    n = n.replace(/(股份有限公司|有限责任公司|有限公司|集团有限公司)$/, "");
  }
  return n || fallback;
}

function companyRemark(t) {
  const api = (typeof window !== "undefined") ? window.EnterpriseData : null;
  const items = splitCompanyNames(t.C).map(function (p, idx) {
    const c = api ? api.find(p) : null;
    return { p: p, idx: idx, c: c };
  });
  items.sort(function (a, b) {
    const da = (a.c && (a.c.defense || a.c.troops)) ? 0 : 1;
    const db = (b.c && (b.c.defense || b.c.troops)) ? 0 : 1;
    return da - db || a.idx - b.idx;
  });
  const names = [];
  const seen = {};
  items.forEach(function (x) {
    let n = wordCompanyName(x.c, x.p);
    n = String(n || "").trim();
    if (!n || seen[n]) return;
    seen[n] = true;
    names.push(n);
  });
  if (names.length > 4) return names.slice(0, 4).join("、") + "等";
  return names.join("、");
}

function docxDescs(t) {
  // 与 Python 端 gen_budget_docx 的描述列口径一致
  const rs = t.ac === 1 ? "" : " x 2/3";
  return {
    t_desc: t.fare.status === "待查" ? "待查" : `${fmtNum(t.T)}/人${rs}`,
    l_desc: t.V === 0 ? "实践地承担"
      : (t.ac === 1 ? `${fmtNum(t.V / 2)}/人·天*${t.nights}晚`
                    : `${fmtNum(t.V / 2)}/人·天*2/3*${t.nights}晚`),
    i_desc: t.weeks > 1 ? `${fmtNum(t.Z)}/人·周*${t.weeks}周${rs}` : `${fmtNum(t.Z)}/人${rs}`,
    remark: companyRemark(t),
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
  const widths = [0.9, 2.8, 0.9, 2.0, 1.2, 2.3, 1.2, 1.7, 1.1, 1.2, 2.7]; // cm，合计18cm适配竖版版心
  let termFlag = String(season || "");
  if (typeof document !== "undefined") {
    const sel = document.getElementById("doc-term");
    if (sel) termFlag += " " + sel.value;
  }
  const inTerm = /学期中/.test(termFlag);

  const teams = computed.map(t => ({ ...t, ...docxDescs(t), remark: companyRemark(t) }));
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

  // 支队目的地列：地点 12 号字、承办院系 10 号字（与公文样例一致）
  const destCell = t => {
    const kids = [
      new Paragraph({ alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: String(t.B || ""), font: F_CELL, size: HP(12) })] }),
    ];
    if (!inTerm && String(t.D || "").trim()) {
      kids.push(new Paragraph({ alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: String(t.D || ""), font: F_CELL, size: HP(10) })] }));
    }
    return new TableCell({ verticalAlign: VerticalAlign.CENTER, children: kids });
  };

  const dataRows = teams.map((t, i) => new TableRow({ children: [
    cell(i + 1), destCell(t), cell(t.people),
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
        ...(inTerm ? [] : [para(note2, F_BODY, 10.5)]),
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

/* ---------- 调整后预算表 → 公文 Word 取值 ----------
 * 优先读表内公式缓存值（用户在 Excel 里的手改随之生效），无缓存的行回退按 A-H 输入列重算 */
function teamsFromSheet(ws, fares) {
  const cities = knownCities(fares);
  const range = XLSX.utils.decode_range(ws["!ref"]);
  const out = [];
  for (let r = range.s.r; r <= range.e.r; r++) {
    const g = c => { const cell = ws[`${c}${r + 1}`]; return cell ? cell.v : undefined; };
    const B = g("B");
    if (!B || typeof B !== "string" || !B.trim()) continue;
    const R = g("R"), K = g("K"), L = g("L"), M = g("M");
    if (typeof R === "number" && typeof K === "number" && L !== undefined && M !== undefined) {
      out.push({ B: String(B).trim(), C: String(g("C") || ""), D: String(g("D") || ""),
        E: String(g("E") || ""), people: K,
        t_desc: String(L), U: Number(M) || 0,
        l_desc: String(g("N") ?? ""), Y: Number(g("O")) || 0,
        i_desc: String(g("P") ?? ""), AA: Number(g("Q")) || 0,
        AB: R, remark: String(g("S") || "") });
    } else if (typeof g("F") === "number") {
      out.push(computeTeam({ B: String(B).trim(), C: g("C") || "", D: g("D") || "",
        E: g("E") || "", F: g("F"), H: g("H") ?? "" }, fares, cities));
    }
  }
  return out;
}

/* node 测试导出 */
if (typeof module !== "undefined") {
  module.exports = { fmtNum, knownCities, parseCities, parseDays, fareFor, ratioOf, computeTeam, buildWorkbook, docxDescs, buildDocx, teamsFromSheet };
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
    if ($("generate-push")) $("generate-push").disabled = true;
    if ($("budget-sync")) { $("budget-sync").disabled = true; $("budget-sync-info").textContent = ""; }
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
        `<td><div class="ent-cell"><button data-ent="${i}" class="ent-btn" title="企业参考">企</button><button data-del="${i}" style="border:none;background:none;cursor:pointer;color:#c00">✕</button></div></td>`;
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
      window.dispatchEvent(new CustomEvent("budget:row-change", { detail: { i, c, value: e.target.value } }));
    }));
    tb.querySelectorAll("[data-ent]").forEach(b => b.addEventListener("click", e => {
      const i = +e.currentTarget.dataset.ent;
      if (window.EnterprisePanel) window.EnterprisePanel.open(i);
    }));
    tb.querySelectorAll("[data-del]").forEach(b => b.addEventListener("click", e => {
      teams.splice(+e.target.dataset.del, 1);
      if (!teams.length) teams = [{ B: "", C: "", D: "就协", E: "", F: 12, H: "" }];
      renderEditor(); saveLocal(); invalidate();
    }));
    window.dispatchEvent(new CustomEvent("budget:render"));
  }

  // 供企业参考侧栏读写当前行（只改 C 列，不影响预算计算）
  window.BudgetBridge = {
    getRow(i) { return teams[i] || null; },
    rowCount() { return teams.length; },
    setPoint(i, value) {
      if (!teams[i]) return;
      teams[i].C = value;
      const inp = document.querySelector(`#editor input[data-i="${i}"][data-c="C"]`);
      if (inp) inp.value = value;
      saveLocal();
      invalidate();
    }
  };

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
    let liveRes = { failed: [], fellBack: [] };
    try {
      liveRes = await prefetchLive(valid, setStatus);
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
    if (liveRes.failed.length)
      warns.push(`以下分段实时查询失败或无直达车：${liveRes.failed.join("、")}——请手填"单人交通费"，或联系实践部用 --set 人工录入`);
    if (liveRes.fellBack.length)
      warns.push(`以下分段实时查询未成功，已使用票价库快照价（每周一自动刷新，影响极小）：${liveRes.fellBack.join("、")}`);
    showWarn(warns);
    showPreview(computed);
    lastComputed = computed;
    $("generate").disabled = false;
    $("generate-docx").disabled = false;
    if ($("generate-push")) $("generate-push").disabled = false;
    if ($("budget-sync")) $("budget-sync").disabled = false;
    return computed;
  }

  /* ---------- 期次下拉：自动拼「公文标题」与「季度表述」 ---------- */
  const TERM_TITLE = { '寒假': '寒假', '暑期': '暑假', '春季学期中': '春季学期中', '秋季学期中': '秋季学期中' };

  function fillYearOptions() {
    const sel = $("doc-year");
    if (!sel || sel.options.length) return;
    const now = new Date().getFullYear();
    for (let y = now - 2; y <= now + 4; y++) {
      const op = document.createElement("option");
      op.value = String(y); op.textContent = y + "年";
      sel.appendChild(op);
    }
    sel.value = String(now);
    const m = ($("doc-season").value || "").match(/(20\d{2})/);
    if (m) sel.value = m[1];
  }

  function applyTerm() {
    const y = $("doc-year").value || String(new Date().getFullYear());
    const term = $("doc-term").value;
    if (!term) {
      $("doc-season").value = "";
      $("doc-title").value = "清华学子地域行实践预算";
      return;
    }
    $("doc-season").value = y + "年" + term;                       // 2027年暑期
    $("doc-title").value = y + "年" + (TERM_TITLE[term] || term) + "“清华学子地域行”实践预算";
  }

  function resetTerm() {
    const term = $("doc-term");
    if (term) term.value = "";
    applyTerm();
  }

  /* ---------- 同步预算到金山「预算」文件 ---------- */
  const SYNC_HOOK = 'https://1386267309-3bh6z757na.ap-shanghai.tencentscf.com';

  function budgetSheetName() {
    const season = ($("doc-season").value || "").trim();   // 例：2027年寒假
    if (season) return season.replace(/年/, "") + "预算";    // 2027寒假预算
    const t = ($("doc-title").value || "").trim();
    return (t.replace(/预算$/, "") || "本期") + "预算";
  }

  async function callBudget(argv) {
    const resp = await fetch(SYNC_HOOK, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ target: "budget", Context: { argv: argv } })
    });
    const d = await resp.json();
    if (d.error) throw new Error(d.error);
    return d.data && d.data.result;
  }

  /** 把「上传的调整后表格」转成金山预算表的行 */
  function uploadedRows(list) {
    return list.map((t, i) => {
      const people = Number(t.people) || 0;
      const ac = /就协|就业协会/.test(String(t.D || "")) ? 1 : 2 / 3;
      const U = Number(t.U) || 0;
      const T = people ? U / people / ac : 0;
      return [i + 1, t.B || "", t.C || "", t.D || "", t.E || "", people,
              fmtNum(T), fmtNum(U), fmtNum(Number(t.Y) || 0),
              fmtNum(Number(t.AA) || 0), fmtNum(Number(t.AB) || 0), t.remark || t.S || ""];
    });
  }

  async function syncBudget() {
    if (!(($("doc-term") || {}).value || "").trim()) {
      alert("请先选择期次（寒假 / 暑期 / 春季学期中 / 秋季学期中），再同步到金山");
      return;
    }
    // 数据来源：优先用完整网格（上传的 / 记住的 / 网页刚生成的 30 列模板）
    const computed = lastComputed;
    if (!computed || !computed.length) {
      alert("请先点「生成预览」核对数据，再同步到金山");
      return;
    }

    let grid = (lastUploadedGrid && lastUploadedGrid.cells && lastUploadedGrid.cells.length)
      ? lastUploadedGrid : null;
    let gridSource = grid ? "已记住的模板" : "";
    if (!grid) {
      try {
        const wb = buildWorkbook(computed);
        const sh = wb.Sheets[wb.SheetNames[0]];
        const g = extractGrid(sh);
        if (g && g.cells.length) { grid = g; gridSource = "网页生成的完整模板"; }
      } catch (e) { /* 生成失败则退回精简列 */ }
    }
    // 数据来自金山读回 → 把编辑器里的最新值再套回网格
    if (grid && gridFromReadback) patchGridWithTeams(grid, teams);
    const useGrid = !!grid;
    const btn = $("budget-sync"), info = $("budget-sync-info");

    // ① 确认是否为最终版
    const totalAB = computed.reduce((a, t) => a + (Number(t.AB) || 0), 0);
    const okFinal = confirm(
      "即将写入金山预算表：\n\n" +
      "工作表：" + budgetSheetName() + "\n" +
      "写入方式：" + (useGrid ? "完整 30 列（" + gridSource + "，含颜色/列宽/合并）" : "精简 12 列") + "\n" +
      "支队数：" + computed.length + " 支\n" +
      "预算总额：" + fmtNum(totalAB) + " 元\n\n" +
      "确认这是最终版本？写入后可在金山继续查看。" +
      "若还要修改，建议先点「取消」，改完再同步。");
    if (!okFinal) return;

    // ② 同名表已存在 → 询问覆盖还是另存新版本
    let sheetName = budgetSheetName();
    try {
      const names = await callBudget({ action: "list" });
      if (Array.isArray(names) && names.indexOf(sheetName) >= 0) {
        const overwrite = confirm(
          "金山里已经有工作表「" + sheetName + "」。\n\n" +
          "点「确定」= 覆盖它（旧内容会被清空，不可恢复）\n" +
          "点「取消」= 另存为新版本（保留旧版）");
        if (!overwrite) {
          const d = new Date(), pad = n => String(n).padStart(2, "0");
          const suggested = sheetName + " V" + `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}-${pad(d.getHours())}${pad(d.getMinutes())}`;
          const input = prompt("新工作表名称：", suggested);
          if (!input) return;
          sheetName = input.trim();
        }
      }
    } catch (e) {
      if (!confirm("⚠ 无法确认是否已有同名工作表（" + e.message + "）。继续写入可能覆盖旧表，是否继续？")) return;
    }

    btn.disabled = true;
    info.textContent = "正在写入金山…";

    // 上传来源且拿到了原始网格 → 原样复刻（列宽/行高/合并/加粗都保留）
    if (useGrid) {
      try {
        const g = grid;
        const r = await callBudget({
          action: "syncBudgetGrid",
          sheetName: sheetName,
          cols: g.cols, rows: g.rows, merges: g.merges,
          cells: g.cells, styles: g.styles, ncols: g.ncols,
          total: totalAB, teams: computed.length
        });
        info.textContent = `✅ 已原样写入金山工作表「${r && r.sheet ? r.sheet : sheetName}」（${r.rows} 行 × ${r.cols} 列，含列宽/行高/合并）`;
        resetTerm();
        info.textContent += "；期次已清空，请重新选择";
        if (window.SJ_OVERVIEW) window.SJ_OVERVIEW.load();
        return;
      } catch (e) {
        info.textContent = "原样写入失败，改用精简格式…（" + e.message + "）";
      } finally {
        btn.disabled = false;
      }
    }
    const rows = computed.map((t, i) => [
      i + 1, t.B, t.C || "", t.D || "", t.E || "", t.people,
      fmtNum(t.T), fmtNum(t.U), fmtNum(t.Y), fmtNum(t.AA), fmtNum(t.AB), t.S || ""
    ]);
    const sumBy = k => computed.reduce((a, t) => a + (Number(t[k]) || 0), 0);
    try {
      const r = await callBudget({
        action: "syncBudget",
        sheetName: sheetName,
        title: ($("doc-title").value || "实践预算") + "表",
        rows: rows,
        total_G: fmtNum(sumBy("T")), total_H: fmtNum(sumBy("U")),
        total_I: fmtNum(sumBy("Y")), total_J: fmtNum(sumBy("AA")),
        total: fmtNum(sumBy("AB"))
      });
      info.textContent = `✅ 已写入金山工作表「${r && r.sheet ? r.sheet : sheetName}」（${computed.length} 支队，合计 ${fmtNum(totalAB)} 元）`;
      resetTerm();
      info.textContent += "；期次已清空，请重新选择";
      if (window.SJ_OVERVIEW) window.SJ_OVERVIEW.load();
    } catch (e) {
      info.textContent = "❌ 同步失败：" + e.message + "（请确认云函数已配 KDOCS_WEBHOOK_BUDGET）";
    } finally {
      btn.disabled = false;
    }
  }

  /* ---------- 从金山读回预算（金山改完 → 网页继续改） ---------- */
  async function refreshBudgetList() {
    const sel = $("budget-load-sheet"), info = $("budget-load-info");
    if (!sel) return;
    sel.innerHTML = "";
    info.textContent = "读取中…";
    try {
      const names = await callBudget({ action: "list" });
      const usable = (names || []).filter(n => !/汇总/.test(n) && !/决算表v1/i.test(n));
      usable.forEach(n => {
        const op = document.createElement("option");
        op.value = n; op.textContent = n; sel.appendChild(op);
      });
      // 默认选名字里含「预算」的最后一个（汇总表已过滤）
      const hits = usable.filter(n => /预算/.test(n));
      if (hits.length) sel.value = hits[hits.length - 1];
      info.textContent = usable.length
        ? `共 ${usable.length} 个工作表（已隐藏汇总表）`
        : "没有可读取的预算工作表（汇总表已隐藏）";
    } catch (e) {
      info.textContent = "读取失败：" + e.message + "（请确认云函数已配 KDOCS_WEBHOOK_BUDGET）";
    }
  }

  async function loadBudgetFromSheet() {
    const sel = $("budget-load-sheet"), info = $("budget-load-info");
    const name = sel && sel.value;
    if (!name) { alert("请先选一个金山工作表（可点「刷新列表」）"); return; }
    if (/汇总/.test(name)) { alert("汇总表不能读回，请选择某一期的预算表"); return; }
    if (lastComputed && !confirm("读回会覆盖当前编辑表格里的内容，继续吗？")) return;
    info.textContent = "读取中…";
    try {
      const res = await callBudget({ action: "readBudget", sheet: name });
      const rows = (res && res.rows) || [];
      if (!rows.length) throw new Error("这个工作表里没有读到支队数据");
      teams = rows.map(r => ({
        B: r.B || "", C: r.C || "", D: r.D || "就协", E: r.E || "",
        F: Number(r.F) || 12,
        T_manual: Number(r.T_manual) > 0 ? Number(r.T_manual) : "",
        H: ""
      }));
      renderEditor(); saveLocal(); invalidate();
      resetTerm();
      if (lastUploadedGrid) {
        gridFromReadback = true;
        patchGridWithTeams(lastUploadedGrid, teams);
        info.textContent = `✅ 已读回「${name}」的 ${rows.length} 支支队；同步时将按你上传的完整列模板写入；期次已清空，请重新选择`;
      } else {
        info.textContent = `✅ 已读回「${name}」的 ${rows.length} 支支队（当前无完整模板，同步会是精简列；上传一次原表后即可全列写入）；期次已清空，请重新选择`;
      }
    } catch (e) {
      info.textContent = "读取失败：" + e.message;
    }
  }

  function splitPushCompanies(C) {
    let parts = String(C || "").split(/[、,，;；\n]+/)
      .map(function (s) {
        return String(s || "")
          .replace(/[（(]\s*(承担住宿|住宿自理|住宿)\s*[）)]/g, "")
          .replace(/承担住宿|住宿自理/g, "")
          .trim();
      })
      .filter(Boolean);
    if (parts.length && parts[parts.length - 1] === "等") {
      parts.pop();
      if (parts.length) parts[parts.length - 1] += "等";
    }
    const api = (typeof window !== "undefined") ? window.EnterpriseData : null;
    return (parts.length ? parts : [""]).map(function (p) {
      const c = api ? api.find(p) : null;
      return c ? wordCompanyName(c, p) : p;
    });
  }

  function setPushCell(tc, lines, W) {
    const ps = Array.from(tc.getElementsByTagNameNS(W, "p"));
    if (!ps.length) return;
    const proto = ps[0].cloneNode(true);
    ps.forEach(function (p) { p.remove(); });
    lines.forEach(function (line) {
      const p = proto.cloneNode(true);
      const runs = Array.from(p.getElementsByTagNameNS(W, "r"));
      runs.slice(1).forEach(function (r) { r.remove(); });
      const ts = Array.from(p.getElementsByTagNameNS(W, "t"));
      if (ts.length) {
        ts[0].textContent = line;
        ts[0].setAttribute("xml:space", "preserve");
        ts.slice(1).forEach(function (t) { t.remove(); });
      }
      tc.appendChild(p);
    });
  }

  async function generatePush(list) {
    const term = (($("doc-term") || {}).value || "").trim();
    if (!term) { alert("请先选择期次（寒假 / 暑期 / 春季学期中 / 秋季学期中），再生成招募推送"); return; }
    const computed = list || lastComputed;
    if (!computed || !computed.length) { alert("请先点击「生成预览」核对数据"); return; }
    const PizZipCtor = window.PizZip;
    if (!PizZipCtor) { alert("推送模板库未加载，请刷新页面后重试"); return; }
    try {
      const res = await fetch("templates/recruitment_template.docx", { cache: "no-cache" });
      if (!res.ok) throw new Error("模板读取失败：HTTP " + res.status);
      const zip = new PizZipCtor(await res.arrayBuffer());
      const docFile = zip.file("word/document.xml");
      if (!docFile) throw new Error("模板缺少 word/document.xml");
      const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
      const xml = new DOMParser().parseFromString(docFile.asText(), "application/xml");
      if (xml.getElementsByTagName("parsererror").length) throw new Error("模板 XML 解析失败");

      Array.from(xml.getElementsByTagNameNS(W, "t")).forEach(function (t) {
        if (t.textContent.indexOf("{{term}}") >= 0) {
          t.textContent = t.textContent.replace(/\{\{term\}\}/g, term);
        }
      });

      try { if (window.EnterpriseData) await window.EnterpriseData.ensure(); } catch (e) { /* 用原文 */ }
      const tbl = xml.getElementsByTagNameNS(W, "tbl")[0];
      if (!tbl) throw new Error("模板缺少表格");
      const trs = Array.from(tbl.getElementsByTagNameNS(W, "tr"));
      if (trs.length < 2) throw new Error("模板缺少数据行");
      const proto = trs[1].cloneNode(true);
      trs.slice(1).forEach(function (tr) { tr.remove(); });

      computed.forEach(function (t, i) {
        const tr = proto.cloneNode(true);
        const tcs = Array.from(tr.getElementsByTagNameNS(W, "tc"));
        setPushCell(tcs[0], [String(i + 1)], W);
        setPushCell(tcs[1], [String(t.B || "")], W);
        setPushCell(tcs[2], [String(t.E || "")], W);
        setPushCell(tcs[3], splitPushCompanies(t.C), W);
        setPushCell(tcs[4], [String(t.D || "就协")], W);
        tbl.appendChild(tr);
      });

      zip.file("word/document.xml", new XMLSerializer().serializeToString(xml));
      const blob = zip.generate({
        type: "blob",
        mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
      });
      const year = (($("doc-year") || {}).value || String(new Date().getFullYear()));
      downloadBlob(blob, `【就协】${year}${term}社会实践招募推送${computed.length}支.docx`);
    } catch (e) {
      alert("生成招募推送失败：" + (e && e.message ? e.message : e));
    }
  }

  function termFromName(name) {
    const s = String(name || "");
    if (/春季学期中/.test(s)) return "春季学期中";
    if (/秋季学期中/.test(s)) return "秋季学期中";
    if (/寒假/.test(s)) return "寒假";
    if (/暑期|暑假/.test(s)) return "暑期";
    return "";
  }

  async function handlePushFile(file) {
    const info = $("push-info");
    if (info) info.textContent = "读取预算 xlsx…";
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf);
      const name = wb.SheetNames.includes("各支队预算按实际人数校准")
        ? "各支队预算按实际人数校准" : wb.SheetNames[0];
      const list = teamsFromSheet(wb.Sheets[name], FARES);
      if (!list.length) throw new Error("没有找到 B 列非空的支队行");
      const term = termFromName(file.name) || (($("doc-term") || {}).value || "");
      if (!term) throw new Error("文件名里没有期次，请先手动选择期次再上传");
      const ym = String(file.name || "").match(/(20\d{2})/);
      if (ym) {
        const sel = $("doc-year");
        if (sel && Array.from(sel.options).some(o => o.value === ym[1])) sel.value = ym[1];
      }
      $("doc-term").value = term;
      applyTerm();

      const raw = list.map(function (t) {
        return { B: t.B || "", C: t.C || "", D: t.D || "就协", E: t.E || "",
          F: t.F != null ? t.F : (t.people != null ? t.people : 12), H: t.H != null ? t.H : "" };
      });
      teams = raw;
      renderEditor();
      saveLocal();
      invalidate();
      if (info) info.textContent = `已读取 ${list.length} 支支队，正在生成推送…`;
      await generatePush(list);
      if (info) info.textContent = `✅ 已生成 ${list.length} 支支队的招募推送`;
    } catch (e) {
      if (info) info.textContent = "生成失败：" + (e && e.message ? e.message : e);
      alert("从预算 xlsx 生成推送失败：" + (e && e.message ? e.message : e));
    }
  }

  async function loadPushFromSheet() {
    const sel = $("budget-load-sheet"), info = $("push-info");
    if (info) info.textContent = "读取金山预算工作表…";
    try {
      if (!sel || !sel.options.length) await refreshBudgetList();
      const name = sel && sel.value;
      if (!name) throw new Error("没有可读取的预算工作表，请先刷新列表");
      if (/汇总/.test(name)) throw new Error("汇总表不能生成推送，请选择某一期预算表");
      const res = await callBudget({ action: "readBudget", sheet: name });
      const rows = (res && res.rows) || [];
      if (!rows.length) throw new Error("这个工作表里没有读到支队数据");
      const term = termFromName(name) || (($("doc-term") || {}).value || "");
      if (!term) throw new Error("无法从工作表名识别期次，请先手动选择期次");
      const ym = name.match(/(20\d{2})/);
      if (ym) {
        const ys = $("doc-year");
        if (ys && Array.from(ys.options).some(o => o.value === ym[1])) ys.value = ym[1];
      }
      $("doc-term").value = term;
      applyTerm();
      const list = rows.map(function (r) {
        return { B: r.B || "", C: r.C || "", D: r.D || "就协", E: r.E || "",
          F: Number(r.F) || 12, H: r.H || "" };
      });
      teams = list;
      renderEditor();
      saveLocal();
      invalidate();
      if (info) info.textContent = `已读取「${name}」${list.length} 支支队，正在生成推送…`;
      await generatePush(list);
      if (info) info.textContent = `✅ 已从「${name}」生成 ${list.length} 支支队的招募推送`;
    } catch (e) {
      if (info) info.textContent = "读取/生成失败：" + (e && e.message ? e.message : e);
      alert("从金山读取预算生成推送失败：" + (e && e.message ? e.message : e));
    }
  }

  async function generate(fmt) {
    const computed = lastComputed;
    if (!computed) { alert("请先点击「生成预览」核对数据"); return; }
    const d = new Date(), pad = n => String(n).padStart(2, "0");
    const stamp = `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}`;
    if (fmt === "xlsx") {
      const wb = buildWorkbook(computed);
      XLSX.writeFile(wb, `实践预算_${stamp}.xlsx`);
    } else {
      if (!(($("doc-term") || {}).value || "").trim()) {
        alert("请先选择期次（寒假 / 暑期 / 春季学期中 / 秋季学期中），再生成公文 Word");
        return;
      }
      try { if (window.EnterpriseData) await window.EnterpriseData.ensure(); } catch (e) { /* 用原始点位名 */ }
      const title = $("doc-title").value.trim() || "实践预算";
      const season = $("doc-season").value.trim() || "本期";
      const blob = await docx.Packer.toBlob(buildDocx(computed, title, season));
      downloadBlob(blob, `${title}.docx`);
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
    resetTerm();
    $("upload-info").textContent = `已导入 ${found.length} 支支队（表：${name}），可切换到「手动填写」查看修改；期次已清空，请重新选择`;
    $("tab-manual").click();
  }

  let lastUploaded = null;       // 最近一次「上传调整后的 xlsx」解析出的支队数据
  let lastUploadedGrid = null;   // 该表的原始网格（值/列宽/行高/合并/加粗/样式），用于原样复刻
  let gridFromReadback = false;  // true = 编辑器的数据来自「金山读回」，同步前需要再套回网格

  /** 从工作表提取「原样复刻」所需的一切：值 + 列宽 + 行高 + 合并 + 样式桶 */
  function extractGrid(ws) {
    if (!ws || !ws["!ref"]) return null;
    const range = XLSX.utils.decode_range(ws["!ref"]);
    const cells = [], buckets = {};
    const hex = c => (c && (c.rgb || c.theme !== undefined ? c.rgb : null)) || null;

    for (let r = range.s.r; r <= range.e.r; r++) {
      for (let c = range.s.c; c <= range.e.c; c++) {
        const addr = XLSX.utils.encode_cell({ r, c });
        const cell = ws[addr];
        if (!cell) continue;
        const st = cell.s || null;
        const fg = st && st.fill && st.fill.fgColor ? hex(st.fill.fgColor) : null;
        const hasVal = cell.v !== undefined && cell.v !== null && cell.v !== "";
        const hasFormula = !!cell.f;
        if (!hasVal && !hasFormula && !fg) continue;       // 空且无底色 → 忽略

        if (hasVal || hasFormula) {
          cells.push([r, c,
            hasVal ? (typeof cell.v === "number" ? cell.v : String(cell.v)) : "",
            hasFormula ? String(cell.f).replace(/^=/, "") : ""]);
        }

        if (st) {
          const sig = {};
          if (st.font) {
            if (st.font.name) sig.n = st.font.name;
            if (st.font.sz) sig.sz = st.font.sz;
            if (st.font.bold) sig.b = 1;
            if (st.font.italic) sig.i = 1;
            const fc = hex(st.font.color); if (fc) sig.c = fc;
          }
          if (fg) sig.fg = fg;
          if (st.alignment) {
            if (st.alignment.horizontal) sig.h = st.alignment.horizontal;
            if (st.alignment.vertical) sig.v = st.alignment.vertical;
            if (st.alignment.wrapText) sig.w = 1;
          }
          if (st.numFmt) sig.nf = st.numFmt;
          const key = JSON.stringify(sig);
          if (key !== "{}") {
            if (!buckets[key]) buckets[key] = { s: sig, cells: [] };
            buckets[key].cells.push([r, c]);
          }
        }
      }
    }

    const cols = (ws["!cols"] || []).map(x => (x && (x.wch || x.width)) || 0);
    const rows = (ws["!rows"] || []).map(x => (x && x.hpt) || 0);
    const merges = (ws["!merges"] || []).map(m => [m.s.r, m.s.c, m.e.r, m.e.c]);

    // 单元格地址拼成「A1,B2,…」批量应用，减少调用次数
    const styles = Object.keys(buckets).map(k => {
      const b = buckets[k];
      return { s: b.s, addr: b.cells.map(([r, c]) =>
        XLSX.utils.encode_cell({ r, c })).join(",") };
    });

    return { cells, cols, rows, merges, styles, ncols: Math.min(range.e.c + 1, 60) };
  }

  /** 找出网格里的表头行与关键列号 */
  function gridHeader(grid) {
    const rowsWithCells = {};
    grid.cells.forEach(([r, c, v]) => {
      (rowsWithCells[r] = rowsWithCells[r] || {})[c] = String(v);
    });
    for (let r = 0; r <= 12; r++) {
      const row = rowsWithCells[r];
      if (!row) continue;
      let placeCol = -1;
      Object.keys(row).forEach(c => {
        const t = (row[c] || "").replace(/\s/g, "");
        if (t.indexOf("实践地点") >= 0 && placeCol < 0) placeCol = Number(c);
      });
      if (placeCol >= 0) {
        const cols = { place: placeCol };
        Object.keys(row).forEach(c => {
          const t = (row[c] || "").replace(/\s/g, "");
          const i = Number(c);
          if (t.indexOf("点位") >= 0 && cols.point === undefined) cols.point = i;
          if (t.indexOf("承办") >= 0 && cols.dept === undefined) cols.dept = i;
          if (t.indexOf("日期") >= 0 && cols.date === undefined) cols.date = i;
          if (t.indexOf("人数") >= 0 && cols.people === undefined) cols.people = i;
          if ((t.indexOf("单人交通") >= 0 || t.indexOf("往返交通") >= 0) && cols.fare === undefined) cols.fare = i;
        });
        return { headerRow: r, cols };
      }
    }
    return null;
  }

  /** 把编辑表格里的值套回完整网格（保留列布局与样式） */
  function patchGridWithTeams(grid, list) {
    const h = gridHeader(grid);
    if (!h) return;
    const dataRows = {};
    grid.cells.forEach(([r, c, v]) => { (dataRows[r] = dataRows[r] || {})[c] = v; });
    const rowIdx = Object.keys(dataRows).map(Number).filter(r => r > h.headerRow &&
      String(dataRows[r][h.cols.place] || "").trim()).sort((a, b) => a - b);
    if (!rowIdx.length) return;

    const index = {};
    dataRows && Object.keys(dataRows).forEach(r => {
      const key = String(dataRows[r][h.cols.place] || "").trim();
      if (key) index[key] = Number(r);
    });

    list.forEach((t, i) => {
      const r = index[String(t.B || "").trim()] !== undefined ? index[String(t.B || "").trim()] : rowIdx[i];
      if (r === undefined) return;
      const put = (col, val) => {
        if (col === undefined || col < 0) return;
        const at = grid.cells.findIndex(x => x[0] === r && x[1] === col);
        if (at >= 0) { if (val === "" ) grid.cells.splice(at, 1); else grid.cells[at][2] = val; }
        else if (val !== "") grid.cells.push([r, col, typeof val === "number" ? val : String(val)]);
      };
      put(h.cols.place, t.B || "");
      put(h.cols.point, t.C || "");
      put(h.cols.dept, t.D || "");
      put(h.cols.date, t.E || "");
      put(h.cols.people, Number(t.F) || "");
      put(h.cols.fare, t.T_manual === "" || t.T_manual === undefined ? "" : Number(t.T_manual));
    });
  }

  async function handleDocxFile(file) {
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf);
    const name = wb.SheetNames.includes("各支队预算按实际人数校准") ? "各支队预算按实际人数校准" : wb.SheetNames[0];
    const teams = teamsFromSheet(wb.Sheets[name], FARES);
    if (!teams.length) { $("docx-info").textContent = "未识别到支队数据行"; return; }
    if (!(($("doc-term") || {}).value || "").trim()) {
      $("docx-info").textContent = "请先选择期次（寒假 / 暑期 / 春季学期中 / 秋季学期中），再上传调整后的 xlsx";
      return;
    }
    const title = $("doc-title").value.trim() || "实践预算";
    const season = $("doc-season").value.trim() || "本期";
    try { if (window.EnterpriseData) await window.EnterpriseData.ensure(); } catch (e) { /* 用原始点位名 */ }
    const blob = await docx.Packer.toBlob(buildDocx(teams, title, season));
    downloadBlob(blob, `${title}.docx`);
    resetTerm();
    lastUploaded = teams;
    lastUploadedGrid = extractGrid(wb.Sheets[name]);
    gridFromReadback = false;
    try {
      localStorage.setItem("budgetGrid", JSON.stringify({ at: Date.now(), file: file.name, grid: lastUploadedGrid }));
      localStorage.setItem("budgetGridTeams", JSON.stringify(teams));
    } catch (e) { /* 超配额就只在内存里用 */ }
    if ($("budget-sync")) $("budget-sync").disabled = false;   // 上传后也可直接同步
    $("docx-info").textContent = `已按表内数值生成 Word（${teams.length} 支支队，表：${name}）；可点右侧按钮同步到金山；期次已清空，请重新选择`;
  }

  document.addEventListener("DOMContentLoaded", async () => {
    loadLocal(); renderEditor();
    fillYearOptions();
    ["doc-year", "doc-term"].forEach(function (id) {
      const el = $(id);
      if (el) el.addEventListener("change", function () { applyTerm(); invalidate && invalidate(); });
    });
    // 恢复上次上传的完整网格模板（刷新页面后仍能全列写入金山）
    try {
      const saved = localStorage.getItem("budgetGrid");
      if (saved) {
        const o = JSON.parse(saved);
        if (o && o.grid && o.grid.cells && o.grid.cells.length) {
          lastUploadedGrid = o.grid;
          if ($("docx-info")) {
            $("docx-info").textContent = "已载入上次上传的模板（" + (o.file || "预算表") +
              "，同步时按全部列写入；如需更换请重新上传）";
          }
        }
      }
    } catch (e) {}
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
    if ($("generate-push")) $("generate-push").addEventListener("click", generatePush);
    if ($("budget-sync")) $("budget-sync").addEventListener("click", syncBudget);
    if ($("budget-load-refresh")) $("budget-load-refresh").addEventListener("click", refreshBudgetList);
    if ($("budget-load")) $("budget-load").addEventListener("click", loadBudgetFromSheet);
    $("drop").addEventListener("click", () => $("file").click());
    $("drop").addEventListener("dragover", e => { e.preventDefault(); $("drop").classList.add("hover"); });
    $("drop").addEventListener("dragleave", () => $("drop").classList.remove("hover"));
    $("drop").addEventListener("drop", e => {
      e.preventDefault(); $("drop").classList.remove("hover");
      if (e.dataTransfer.files[0]) handleFile(e.dataTransfer.files[0]);
    });
    $("file").addEventListener("change", e => { if (e.target.files[0]) handleFile(e.target.files[0]); });
    $("upload-docx").addEventListener("click", () => $("file-docx").click());
    $("file-docx").addEventListener("change", e => { if (e.target.files[0]) handleDocxFile(e.target.files[0]); });
    if ($("upload-push")) $("upload-push").addEventListener("click", () => $("file-push").click());
    if ($("file-push")) $("file-push").addEventListener("change", e => { if (e.target.files[0]) handlePushFile(e.target.files[0]); });
    if ($("load-push")) $("load-push").addEventListener("click", loadPushFromSheet);

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
