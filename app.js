/* 实践支队预算表生成器 —— 纯前端逻辑
 * 口径与 实践部/自动化/*.py 完全一致：
 *   交通 = 行程链逐段二等座最高票价之和（北京→A→B→北京）
 *   住宿 = 标间300 × ceil(人数/2) × (天数-1) × 比例（含"承担住宿"字样则 0）
 *   保险 = 30 × 人数 × ceil(天数/7) × 比例（30元/人/周）
 *   比例 = 承办单位含"就协"/"就业协会" → 1，否则 2/3
 */

const KNOWN_CITIES = ["上海", "西安", "武汉", "合肥", "广州", "深圳", "珠海", "成都",
  "福州", "宁德", "南京", "苏州", "无锡", "杭州", "重庆", "江门",
  "青岛", "沈阳", "大连", "长沙", "郑州", "天津"];

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
  const fare = fareFor(cs, fares);
  const T = fare.total;
  const nights = days ? days - 1 : 0;
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

/* node 测试导出 */
if (typeof module !== "undefined") {
  module.exports = { fmtNum, knownCities, parseCities, parseDays, fareFor, ratioOf, computeTeam, buildWorkbook };
}

/* ---------- 以下为浏览器 UI ---------- */
if (typeof document !== "undefined") {
  const $ = id => document.getElementById(id);
  let teams = [];

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
        ["B", "C", "D", "E", "F", "H"].map(c =>
          `<td><input data-i="${i}" data-c="${c}" value="${String(t[c] ?? "").replace(/"/g, "&quot;")}" ` +
          (c === "B" ? `list="city-list" ` : "") +
          `placeholder="${{ B: "合肥、上海", C: "企业全称…", D: "就协", E: "7.19-7.23", F: "12", H: "0" }[c]}"></td>`).join("") +
        `<td><button data-del="${i}" style="border:none;background:none;cursor:pointer;color:#c00">✕</button></td>`;
      tb.appendChild(tr);
    });
    tb.querySelectorAll("input").forEach(inp => inp.addEventListener("input", e => {
      const i = +e.target.dataset.i, c = e.target.dataset.c;
      teams[i][c] = c === "F" ? (Number(e.target.value) || "") : e.target.value;
      if (c === "E") {  // 日期即时校验：可解析变绿、不可解析变红、空白还原
        const v = e.target.value.trim();
        e.target.className = !v ? "" : (parseDays(v) ? "good" : "bad");
        e.target.title = e.target.className === "bad" ? "无法识别，格式应为 7.19-7.23" : "";
      }
      if (c === "B" && FARES) {  // 地点即时校验：未识别到票价库城市时标红
        const v = e.target.value.trim();
        const hit = !v || parseCities(v, knownCities(FARES)).length > 0;
        e.target.className = !v ? "" : (hit ? "good" : "bad");
        e.target.title = hit ? "" : "未识别到票价库中的城市，生成时将标'待查'";
      }
      saveLocal();
    }));
    tb.querySelectorAll("[data-del]").forEach(b => b.addEventListener("click", e => {
      teams.splice(+e.target.dataset.del, 1);
      if (!teams.length) teams = [{ B: "", C: "", D: "就协", E: "", F: 12, H: "" }];
      renderEditor(); saveLocal();
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

  function generate() {
    const valid = collectValid();
    if (!valid.length) { alert("请至少填写一支支队（实践地点 + 拟参访日期必填）"); return; }
    const cities = knownCities(FARES);
    const computed = valid.map(t => computeTeam(t, FARES, cities));
    const warns = [];
    computed.forEach(t => {
      if (t.fare.status === "待查") warns.push(`「${t.B}」票价${t.fare.basis}——xlsx 中 T 列留空，请查询后手填`);
      if (!t.days) warns.push(`「${t.B}」日期「${t.E}」无法识别，天数/住宿/保险未算`);
      if (t.fare.status !== "分段" && t.fare.status !== "待查") warns.push(`「${t.B}」使用${t.fare.status}（非实时分段价）`);
    });
    showWarn(warns);
    showPreview(computed);
    const wb = buildWorkbook(computed);
    const d = new Date(), pad = n => String(n).padStart(2, "0");
    XLSX.writeFile(wb, `实践预算_${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}.xlsx`);
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
    saveLocal(); renderEditor();
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
      renderEditor(); saveLocal();
    });
    $("generate").addEventListener("click", generate);
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
      $("generate").disabled = false;
      // 地点自动联想候选（票价库全部城市）
      $("city-list").innerHTML = knownCities(FARES).map(c => `<option value="${c}">`).join("");
    } catch (e) {
      $("fare-info").textContent = "⚠ fares.json 加载失败（请通过 http(s) 访问本页，file:// 直开无法读取票价库）";
      $("generate").disabled = true;
    }
  });
}
