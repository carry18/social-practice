/* 12306 二等座最高票价代理 —— 腾讯云函数 SCF（Node.js 18，零依赖）
 * 部署：SCF 新建函数 → 粘贴本文件全部内容 → 启用函数 URL（匿名访问）→ 超时设 60s
 * 调用：GET <函数URL>?from=北京&to=昆明&days=14[&date=2026-07-01]
 * 返回：{price, train, from_station, to_station, query_date, window, days_queried, days_skipped}
 * 口径与 fare_query.py 完全一致：同城站按官方城市归属过滤、绕路车按历时过滤、
 * 14 天预售窗口逐日取最高、ze_price 单位角。
 */
const https = require("https");

const STATION_JS_URL = "https://kyfw.12306.cn/otn/resources/js/framework/station_name.js";
const INIT_URL = "https://kyfw.12306.cn/otn/leftTicket/init";
const PRICE_URL = "https://kyfw.12306.cn/otn/leftTicketPrice/queryAllPublicPrice";
const HEADERS = {
  "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36",
  "Referer": "https://kyfw.12306.cn/otn/leftTicket/init",
  "X-Requested-With": "XMLHttpRequest",
  "Accept": "application/json, text/javascript, */*; q=0.01",
};
const PRESALE_DAYS = 15, REQUEST_GAP_MS = 1000, TIMEOUT_MS = 12000;
const FARE_TTL_MS = 6 * 3600 * 1000;        // 票价结果缓存 6h（暖实例内）
const STATION_TTL_MS = 24 * 3600 * 1000;    // 站表缓存 24h

/* ---------- 暖实例缓存（SCF 实例复用时跨请求生效） ---------- */
const g = globalThis.__fareCache || (globalThis.__fareCache = {
  stations: null, stationsAt: 0, fares: new Map(), lastReq: 0,
});

/* ---------- HTTP（自管 cookie 罐，12306 证书链不全需 rejectUnauthorized:false） ---------- */
function httpGet(url, jar) {
  return new Promise((resolve, reject) => {
    const u = new URL(url);
    const cookie = Object.entries(jar).map(([k, v]) => `${k}=${v}`).join("; ");
    const req = https.request({
      hostname: u.hostname, path: u.pathname + u.search, method: "GET",
      rejectUnauthorized: false,
      headers: { ...HEADERS, ...(cookie ? { Cookie: cookie } : {}) },
      timeout: TIMEOUT_MS,
    }, res => {
      for (const c of res.headers["set-cookie"] || []) {
        const pair = c.split(";")[0], i = pair.indexOf("=");
        if (i > 0) jar[pair.slice(0, i).trim()] = pair.slice(i + 1).trim();
      }
      if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
        res.resume();
        return resolve(httpGet(new URL(res.headers.location, url).href, jar));
      }
      const chunks = [];
      res.on("data", c => chunks.push(c));
      res.on("end", () => resolve({ status: res.statusCode, body: Buffer.concat(chunks).toString("utf8") }));
    });
    req.on("timeout", () => req.destroy(new Error("请求超时")));
    req.on("error", reject);
    req.end();
  });
}

async function throttle() {
  const wait = REQUEST_GAP_MS - (Date.now() - g.lastReq);
  if (wait > 0) await new Promise(r => setTimeout(r, wait));
  g.lastReq = Date.now();
}

/* ---------- 站表 ---------- */
async function loadStations() {
  if (g.stations && Date.now() - g.stationsAt < STATION_TTL_MS) return g.stations;
  const { body } = await httpGet(STATION_JS_URL, {});
  const codes = {}, cityOf = {};
  for (const rec of body.split("@").slice(1)) {
    const p = rec.split("|");
    if (p.length >= 8) { codes[p[1]] = p[2].toUpperCase(); cityOf[p[1]] = p[7]; }
  }
  const nameOf = {};
  for (const [n, c] of Object.entries(codes)) nameOf[c] = n;
  g.stations = { codes, cityOf, nameOf };
  g.stationsAt = Date.now();
  return g.stations;
}

function toTelecode(name, codes) {
  name = String(name || "").trim();
  if (/^[A-Za-z]{3}$/.test(name)) return name.toUpperCase();
  if (codes[name]) return codes[name];
  const cands = Object.keys(codes).filter(n => n.startsWith(name)).sort((a, b) => a.length - b.length);
  if (cands.length) return codes[cands[0]];
  throw new Error(`未识别车站: ${name}`);
}

function queryCity(name, st) {
  const stn = st.nameOf[toTelecode(name, st.codes)];
  return st.cityOf[stn] || String(name).trim();
}

function parsePrice(raw) {
  if (!raw || raw === "--" || !/^\d+$/.test(raw)) return null;
  return parseFloat(raw.slice(0, -1) + "." + raw.slice(-1));
}

/* ---------- 票价查询 ---------- */
function fmtDay(d) { return d.toISOString().slice(0, 10); } // 仅用于日期运算（UTC安全：纯日期）

async function queryDay(jar, st, fromName, toName, day) {
  const qs = `leftTicketDTO.train_date=${day}&leftTicketDTO.from_station=${toTelecode(fromName, st.codes)}` +
    `&leftTicketDTO.to_station=${toTelecode(toName, st.codes)}&purpose_codes=ADULT`;
  let lastErr;
  for (let attempt = 0; attempt < 3; attempt++) {
    await throttle();
    try {
      const { status, body } = await httpGet(`${PRICE_URL}?${qs}`, jar);
      if (status !== 200) throw new Error(`HTTP ${status}`);
      const data = (JSON.parse(body).data) || [];
      const rows = [];
      for (const item of data) {
        const dto = item.queryLeftNewDTO || {};
        const p = parsePrice(dto.ze_price || "");
        if (p === null) continue;
        const [h, m] = (dto.lishi || "99:59").split(":");
        const mins = (/^\d+$/.test(h) && /^\d+$/.test(m)) ? (+h) * 60 + (+m) : 9999;
        rows.push({ p, train: dto.station_train_code || "?", from: dto.from_station_name || "",
                    to: dto.to_station_name || "", mins });
      }
      const fc = queryCity(fromName, st), tc = queryCity(toName, st);
      let main = rows.filter(r => st.cityOf[r.from] === fc && st.cityOf[r.to] === tc);
      if (!main.length)
        main = rows.filter(r => r.from.startsWith(fromName) && r.to.startsWith(toName));
      const pool = main.length ? main : rows;
      if (pool.length) {
        const shortest = Math.min(...pool.map(r => r.mins));
        const cap = Math.max(shortest * 2, shortest + 60);
        const direct = pool.filter(r => r.mins <= cap);
        return direct.length ? direct : pool;
      }
      return rows;
    } catch (e) {
      lastErr = e;
      await new Promise(r => setTimeout(r, 2000 * (attempt + 1)));
    }
  }
  throw new Error(`查询失败（已重试3次）: ${lastErr && lastErr.message}`);
}

async function maxFareRange(fromName, toName, start, days) {
  const st = await loadStations();
  const jar = {};
  await throttle();
  await httpGet(INIT_URL, jar);   // 会话 cookie
  // 固定按北京时间计算"今天"，与函数运行时区无关（SCF 默认 UTC）
  const bjNow = new Date(Date.now() + 8 * 3600 * 1000);
  const today = new Date(Date.UTC(bjNow.getUTCFullYear(), bjNow.getUTCMonth(), bjNow.getUTCDate()));
  const saleEnd = new Date(today.getTime() + (PRESALE_DAYS - 1) * 86400000);
  const first = start ? new Date(start + "T00:00:00Z") : new Date(today.getTime() + 86400000);
  const last = new Date(first.getTime() + (days - 1) * 86400000);
  let best = null, queried = 0, skipped = 0;
  for (let d = new Date(first); d <= last; d = new Date(d.getTime() + 86400000)) {
    if (d > saleEnd || d < today) { skipped++; continue; }
    const rows = await queryDay(jar, st, fromName, toName, fmtDay(d));
    if (rows.length) {
      queried++;
      const dayBest = rows.reduce((a, b) => (b.p > a.p ? b : a));
      if (!best || dayBest.p > best.p) best = { ...dayBest, count: rows.length, day: fmtDay(d) };
    }
  }
  if (!best) {
    const rows = await queryDay(jar, st, fromName, toName, fmtDay(saleEnd));
    if (!rows.length) throw new Error(`${fromName}→${toName} 未查到任何有二等座的车次（无直达车？）`);
    const dayBest = rows.reduce((a, b) => (b.p > a.p ? b : a));
    best = { ...dayBest, count: rows.length, day: fmtDay(saleEnd) };
    queried = 1; skipped = 0;
  }
  return {
    price: best.p, train: best.train, from_station: best.from, to_station: best.to,
    train_count: best.count, query_date: best.day,
    window: days > 1 ? `${fmtDay(first)}~${fmtDay(last)}` : null,
    days_queried: queried, days_skipped: skipped,
  };
}

/* ---------- SCF 入口（函数 URL） ---------- */
function extractQuery(event) {
  // 腾讯实际格式：queryStringParameters 可能是空对象（真值！），参数在 queryString 对象里。
  // 因此必须逐个判空，不能用 || 短路。
  let q = event.queryStringParameters;
  if (!q || typeof q !== "object" || !Object.keys(q).length) q = event.queryString;
  if (q && typeof q === "object" && Object.keys(q).length) return q;
  const mv = event.multiValueQueryStringParameters;
  if (mv && typeof mv === "object") {
    const out = {};
    for (const [k, v] of Object.entries(mv)) out[k] = Array.isArray(v) ? v[0] : v;
    if (Object.keys(out).length) return out;
  }
  const raw = typeof q === "string" ? q
    : (event.rawQueryString || (String(event.path || "").includes("?") ? String(event.path).split("?")[1] : ""));
  const out = {};
  try { for (const [k, v] of new URLSearchParams(raw)) out[k] = v; } catch (e) {}
  return out;
}

function safeDecode(s) {
  try { return decodeURIComponent(String(s || "")); } catch (e) { return String(s || ""); }
}

exports.main_handler = async (event) => {
  const cors = {
    "Content-Type": "application/json; charset=utf-8",
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "GET, OPTIONS",
  };
  const method = (event.requestContext && event.requestContext.http && event.requestContext.http.method) || event.httpMethod || "GET";
  if (method === "OPTIONS") return { statusCode: 204, headers: cors, body: "" };
  try {
    const q = extractQuery(event);
    const from = safeDecode(q.from), to = safeDecode(q.to);
    if (!from || !to) {
      // 调试回显：把实际收到的事件结构带回去，便于确认平台参数格式
      return { statusCode: 200, headers: cors, body: JSON.stringify({
        error: "缺少参数：?from=起点&to=终点[&days=14][&date=YYYY-MM-DD]",
        debug: {
          httpMethod: event.httpMethod, path: event.path,
          queryString: event.queryString, queryStringParameters: event.queryStringParameters,
          requestContext: JSON.stringify(event.requestContext || {}).slice(0, 300),
        },
      }) };
    }
    const days = Math.min(Math.max(parseInt(q.days, 10) || 14, 1), 15);
    const date = /^\d{4}-\d{2}-\d{2}$/.test(q.date || "") ? q.date : null;
    const key = `${from}-${to}@${date || "auto"}+${days}d`;
    const hit = g.fares.get(key);
    if (hit && Date.now() - hit.at < FARE_TTL_MS)
      return { statusCode: 200, headers: { ...cors, "X-Cache": "HIT" }, body: JSON.stringify({ ...hit.data, cached: true }) };
    const data = await maxFareRange(from, to, date, days);
    g.fares.set(key, { at: Date.now(), data });
    return { statusCode: 200, headers: { ...cors, "X-Cache": "MISS" }, body: JSON.stringify({ ...data, cached: false }) };
  } catch (e) {
    return { statusCode: 502, headers: cors, body: JSON.stringify({ error: String(e.message || e) }) };
  }
};
