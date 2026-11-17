/**
 * 金山文档 AirScript：实践台账「一键建表 + 增量入库 + 淘汰/捞起归位」
 * ------------------------------------------------------------------
 * 用法：金山文档新建一个空白表格 → 效率 → AirScript → 新建「文档共享脚本」
 *       → 贴入本文件 → 保存 → 更多 → 复制 webhook 链接（给云函数用）。
 *
 * 由云函数调用（HTTP POST，body：{Context:{argv:{...}}}）：
 *   action = 'ping'            连通自检，返回 pong
 *   action = 'ensure'          建表 + 入库（argv.teams 支队列表，argv.rows 答卷行）
 *   action = 'appendEntries'   只入库（表已存在时用）
 *   action = 'reflow'          归位：分表「淘汰」→ 淘汰池；淘汰池「捞入支队」→ 回分表
 *
 * 表结构（一期一份文档）：
 *   支队对照       序号 | 支队名（=分表名） | 名额 | 已分配 | 地点 | 院系 | 参访时间 | 人数
 *   支队概况总览   序号 | 支队名 | 地点 | 承办院系 | 参访时间 | 天数 | 人数 | 名额 | 拟参访单位
 *                  | 协调人 | 同意类型 | 是否拉群 | 是否领队旗 | 支队长 | 支队长联系方式 | 联络人（就协）
 *   企业联络需求表 序号 | 支队名 | 企业名称 | 企业联系人 | 职位 | 企业联系方式 | 备注
 *   开函统计       批次 | 序号 | 院系 | 单位名称 | 对方单位是否审核完毕 | 是否上传网盘 | 联系人
 *                  | 联系方式 | 上传时间 | 已盖章 | 院系接受后备注 | 备注
 *   总表           分配状态 | 序号 | 姓名 | 少干/定向 | 学号 | 手机号 | 性别 | 院系 | 专业
 *                  | 拟毕业时间 | 学历 | 本科院校 | 本科专业 | 简历链接 | 愿任支队长 | 家乡
 *                  | 志愿1..N | entry_id
 *   淘汰池         分配状态 | 捞入支队 | 序号 | 姓名 | … | entry_id
 *   各支队分表     录取状态 | 序号 | 姓名 | … | 家乡 | entry_id
 * ------------------------------------------------------------------
 */

/* ====== 每期改这里（生成好直接粘） ====== */
var CONFIG = {
  term: '2026秋季学期中',
  form: 'QcMAb9',                 // 金数据报名表 token
  token: '<JINSHUJU_TOKEN>',      // 金数据 Access Token
  teams: []                       // 由本地脚本生成后填入：[{num,label,place,dept,date,people,quota,days,company}]
};
/* ======================================= */

var argv = (typeof Context !== 'undefined' && Context && Context.argv) || {};
var action = argv.action || 'ping';

var COLORS = [[80, 129, 194], [31, 187, 106], [228, 180, 13], [233, 85, 45], [142, 100, 214],
              [0, 166, 166], [221, 69, 149], [240, 146, 0], [108, 142, 191], [124, 179, 66],
              [176, 133, 91], [61, 71, 87]];

var BASE_HEAD = ['序号', '姓名', '少干/定向', '学号', '手机号', '性别', '院系', '专业',
                 '拟毕业时间', '学历', '本科院校', '本科专业', '简历链接', '愿任支队长', '家乡'];

/* ---------------- 基础工具 ---------------- */

function sheet(name) {
  try { return Application.Sheets(name); } catch (e) { return null; }
}

function lastRow(ws) {
  try {
    var r = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1;
    return r > 0 ? r : 1;
  } catch (e) { return 1; }
}

function headerMap(ws) {
  var map = {}, cols = ws.UsedRange.Columns.Count;
  for (var c = 1; c <= cols; c++) {
    var v = String(fmt1(ws.Cells(1, c).Value2)).trim();
    if (v && map[v] === undefined) map[v] = c;
  }
  return map;
}

function fmt1(x) { return x === null || x === undefined ? '' : x; }

function ensureSheet(name) {
  var ws = sheet(name);
  if (ws) return ws;
  var s = Application.Sheets.Add();
  s.Name = name;
  return s;
}

function writeHeader(ws, head) {
  ws.Range(ws.Cells(1, 1), ws.Cells(1, head.length)).Value2 = [head];
  ws.Rows(1).Font.Bold = true;
  ws.Rows(1).Interior.Color = RGB(74, 111, 165);
  ws.Rows(1).Font.Color = RGB(255, 255, 255);
}

function setWidths(ws, widths) {
  for (var i = 0; i < widths.length; i++) {
    try { ws.Columns(i + 1).ColumnWidth = widths[i]; } catch (e) {}
  }
}

function teamIndex(teams) {
  var out = {};
  (teams || []).forEach(function (t) { out[String(t.num)] = t; });
  return out;
}

function readRows(ws) {
  var n = lastRow(ws);
  if (n < 2) return [];
  var ncol = ws.UsedRange.Columns.Count;
  var vals = ws.Range(ws.Cells(2, 1), ws.Cells(n, ncol)).Value2;
  if (!vals || !vals.length) return [];
  var out = [];
  for (var i = 0; i < vals.length; i++) {
    var row = vals[i];
    if (!row) continue;
    var has = false;
    for (var j = 0; j < row.length; j++) {
      if (String(fmt1(row[j])).trim() !== '') { has = true; break; }
    }
    if (has) out.push({ r: i + 2, v: row });
  }
  return out;
}

function pickByHeader(srcVals, srcMap, dstMap, skipCount) {
  var names = Object.keys(dstMap);
  names.sort(function (a, b) { return dstMap[a] - dstMap[b]; });
  var out = [];
  for (var i = skipCount || 0; i < names.length; i++) {
    var sc = srcMap[names[i]];
    out.push(sc ? fmt1(srcVals[sc - 1]) : '');
  }
  return out;
}

/* 给区域加下拉列表（JSA：Validation.Add(3=列表, 1=错误提示, 1=介于, formula1)） */
function addList(ws, addr, items) {
  try {
    var rng = ws.Range(addr);
    try { rng.Validation.Delete(); } catch (e0) {}
    var f = '"' + items.join(',') + '"';
    rng.Validation.Add(3, 1, 1, f);
    return true;
  } catch (e) { return false; }
}

function colorCell(ws, r, c, num) {
  if (!num) return;
  var k = (Number(num) - 1) % COLORS.length;
  var p = COLORS[k];
  ws.Cells(r, c).Interior.Color = RGB(p[0], p[1], p[2]);
  ws.Cells(r, c).Font.Color = RGB(255, 255, 255);
}

/* ---------------- ensure：建表 ---------------- */

function ensureSheets(teams) {
  var n = teams.length;

  // 1. 支队对照
  var ws = ensureSheet('支队对照');
  ws.Range('A1:H1').ClearContents();
  writeHeader(ws, ['序号', '支队名（=分表名）', '名额', '已分配', '地点', '院系', '参访时间', '人数']);
  setWidths(ws, [6, 46, 8, 8, 14, 20, 14, 6]);
  teams.forEach(function (t, i) {
    var r = i + 2;
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 8)).Value2 = [[
      Number(t.num), t.label, Number(t.quota) || 12, '',
      t.place || '', t.dept || '', t.date || '', Number(t.people) || 12
    ]];
    ws.Cells(r, 4).Formula = '=COUNTIF(总表!$A:$A,' + Number(t.num) + ')';
    colorCell(ws, r, 3, t.num);
  });

  // 2. 支队概况总览
  ws = ensureSheet('支队概况总览');
  writeHeader(ws, ['序号', '支队名', '地点', '承办院系', '参访时间', '天数', '人数', '名额',
                   '拟参访单位', '协调人', '同意类型', '是否拉群', '是否领队旗', '支队长',
                   '支队长联系方式', '联络人（就协）']);
  setWidths(ws, [6, 34, 14, 22, 14, 6, 6, 8, 40, 10, 12, 10, 12, 10, 16, 14]);
  if (lastRow(ws) < teams.length + 1) {
    teams.forEach(function (t, i) {
      var r = i + 2;
      ws.Range(ws.Cells(r, 1), ws.Cells(r, 9)).Value2 = [[
        Number(t.num), t.label, t.place || '', t.dept || '', t.date || '',
        Number(t.days) || '', Number(t.people) || 12, Number(t.quota) || 12, t.company || ''
      ]];
    });
  }

  // 3. 企业联络需求表
  ws = ensureSheet('企业联络需求表');
  writeHeader(ws, ['序号', '支队名', '企业名称', '企业联系人', '职位', '企业联系方式', '备注']);
  setWidths(ws, [6, 34, 40, 14, 14, 16, 20]);
  if (lastRow(ws) < 2) {
    var r3 = 2;
    teams.forEach(function (t) {
      String(t.company || '').split(/[、,，;；\n]+/).forEach(function (comp) {
        comp = String(comp || '').replace(/[（(]\s*(承担住宿|住宿自理)\s*[）)]/g, '').trim();
        if (!comp) return;
        ws.Range(ws.Cells(r3, 1), ws.Cells(r3, 3)).Value2 = [[Number(t.num), t.label, comp]];
        r3++;
      });
    });
  }

  // 4. 开函统计
  ws = ensureSheet('开函统计');
  writeHeader(ws, ['批次', '序号', '院系', '单位名称', '对方单位是否审核完毕', '是否上传网盘',
                   '联系人', '联系方式', '上传时间', '已盖章', '院系接受后备注', '备注']);
  setWidths(ws, [8, 6, 24, 36, 14, 12, 10, 14, 12, 8, 18, 16]);
  if (lastRow(ws) < teams.length + 1) {
    teams.forEach(function (t, i) {
      ws.Range(ws.Cells(i + 2, 1), ws.Cells(i + 2, 3)).Value2 = [['', Number(t.num), t.dept || '']];
    });
  }

  // 5. 总表
  ws = ensureSheet('总表');
  var head = ['分配状态'].concat(BASE_HEAD);
  for (var v = 1; v <= n; v++) head.push('志愿' + v);
  head.push('entry_id');
  writeHeader(ws, head);
  var widths = [10, 6, 10, 10, 14, 14, 6, 14, 18, 12, 8, 14, 18, 30, 10, 12];
  for (var x = 0; x < n; x++) widths.push(26);
  widths.push(12);
  setWidths(ws, widths);
  try { ws.Columns(head.length).Hidden = true; } catch (e) {}   // 隐藏 entry_id
  var opts = [];
  for (var q = 1; q <= n; q++) opts.push(q);
  opts.push('淘汰');
  addList(ws, 'A2:A1000', opts);                                 // 分配状态下拉

  // 6. 淘汰池
  ws = ensureSheet('淘汰池');
  writeHeader(ws, ['分配状态', '捞入支队'].concat(BASE_HEAD).concat(['entry_id']));
  setWidths(ws, [10, 10, 6, 10, 10, 14, 14, 6, 14, 18, 12, 8, 14, 18, 30, 10, 12, 12]);
  var pn = [];
  for (var q2 = 1; q2 <= n; q2++) pn.push(q2);
  addList(ws, 'B2:B1000', pn);                                    // 捞入支队下拉

  // 7. 各支队分表
  teams.forEach(function (t) {
    var name = String(t.label).substring(0, 31);
    var ts = ensureSheet(name);
    if (lastRow(ts) < 1 || String(fmt1(ts.Cells(1, 1).Value2)).trim() === '') {
      writeHeader(ts, ['录取状态'].concat(BASE_HEAD).concat(['entry_id']));
      setWidths(ts, [10, 6, 10, 10, 14, 14, 6, 14, 18, 12, 8, 14, 18, 30, 10, 12, 12]);
    }
    addList(ts, 'A2:A1000', ['录取', '淘汰']);                     // 录取状态下拉（每次都补）
  });

  return 'sheets ready: ' + (5 + teams.length) + ' 张';
}

/* ---------------- appendEntries：入库 ---------------- */

function appendEntries(rows) {
  var ws = sheet('总表');
  if (!ws) throw new Error('找不到「总表」工作表');
  var m = headerMap(ws);
  var cAssign = m['分配状态'], cId = m['entry_id'];
  if (!cId) throw new Error('总表缺少 entry_id 列');
  var volCols = [];
  for (var v = 1; m['志愿' + v]; v++) volCols.push(m['志愿' + v]);

  var idx = {};
  var n = lastRow(ws);
  for (var r = 2; r <= n; r++) {
    var id = String(fmt1(ws.Cells(r, cId).Value2)).trim();
    if (id) idx[id] = r;
  }

  var base = function (row) {
    var keys = ['sn'].concat(BASE_HEAD.slice(1).map(function (h) { return HEAD2KEY[h]; }));
    var out = [];
    keys.forEach(function (k) { out.push(row[k] == null ? '' : row[k]); });
    return out;
  };

  var added = 0, updated = 0;
  rows.forEach(function (row) {
    if (!row || !row.id) return;
    var b = base(row), vols = row.vols || [];
    var rr = idx[row.id];
    if (rr) {
      ws.Range(ws.Cells(rr, m['序号']), ws.Cells(rr, m['序号'] + b.length - 1)).Value2 = [b];
      if (volCols.length) {
        ws.Range(ws.Cells(rr, volCols[0]), ws.Cells(rr, volCols[volCols.length - 1])).Value2 = [vols];
      }
      updated++;
    } else {
      rr = lastRow(ws) + 1;
      ws.Cells(rr, cAssign).Value2 = row.assign || '';
      colorCell(ws, rr, cAssign, row.assign);
      ws.Range(ws.Cells(rr, m['序号']), ws.Cells(rr, m['序号'] + b.length - 1)).Value2 = [b];
      if (volCols.length) {
        ws.Range(ws.Cells(rr, volCols[0]), ws.Cells(rr, volCols[volCols.length - 1])).Value2 = [vols];
      }
      ws.Cells(rr, cId).Value2 = row.id;
      idx[row.id] = rr;
      added++;
    }
  });

  return 'added=' + added + ', updated=' + updated + ', total=' + (lastRow(ws) - 1);
}

/* 字段名 → 答卷行里的键 */
var HEAD2KEY = {
  '序号': 'sn', '姓名': 'name', '少干/定向': 'minority', '学号': 'sid', '手机号': 'phone',
  '性别': 'gender', '院系': 'dept', '专业': 'major', '拟毕业时间': 'grad', '学历': 'degree',
  '本科院校': 'bschool', '本科专业': 'bmajor', '简历链接': 'resume', '愿任支队长': 'wantLeader', '家乡': 'hometown'
};

/* ---------------- reflow：淘汰 / 捞起归位 ---------------- */

function reflow(force) {
  // 防抖：自动化可能对每次编辑都触发，30 秒内重复调用直接跳过（main 里传 true 强制）
  try {
    var lock = ensureSheet('同步日志');
    var last = Number(lock.Cells(1, 4).Value2 || 0);
    var now = new Date().getTime();
    if (!force && last && now - last < 30000) return 'reflow skipped (30s 防抖)';
    lock.Cells(1, 4).Value2 = now;
  } catch (e0) {}

  var ws = sheet('总表');
  if (!ws) throw new Error('找不到「总表」工作表');
  var m = headerMap(ws);
  var cAssign = m['分配状态'], cId = m['entry_id'];

  var assign = {};
  var n = lastRow(ws);
  for (var r = 2; r <= n; r++) {
    var id = String(fmt1(ws.Cells(r, cId).Value2)).trim();
    if (id) assign[id] = String(fmt1(ws.Cells(r, cAssign).Value2)).trim();
  }

  // ① 分表「淘汰」（同时记下非淘汰行的录取状态，重建后写回）
  var moved = 0;
  var statusOf = {};
  var sheetNames = [];
  for (var i = 1; i <= Application.Sheets.Count; i++) sheetNames.push(Application.Sheets.Item(i).Name);
  sheetNames.forEach(function (name) {
    if (name === '总表' || name === '淘汰池' || name === '支队对照' || name === '支队概况总览' ||
        name === '企业联络需求表' || name === '开函统计') return;
    var ts = sheet(name);
    if (!ts) return;
    var tm = headerMap(ts);
    if (!tm['entry_id'] || !tm['录取状态']) return;
    readRows(ts).forEach(function (o) {
      var status = String(fmt1(o.v[0])).trim();
      var tid = String(fmt1(o.v[tm['entry_id'] - 1])).trim();
      if (tid && status === '淘汰') { assign[tid] = '淘汰'; moved++; }
      else if (tid && status) { statusOf[tid] = status; }
    });
  });

  // ② 淘汰池「捞入支队」
  var ps = sheet('淘汰池');
  var picked = 0;
  if (ps) {
    var pm = headerMap(ps);
    if (pm['entry_id'] && pm['捞入支队']) {
      readRows(ps).forEach(function (o) {
        var target = String(fmt1(o.v[pm['捞入支队'] - 1])).trim();
        var tid = String(fmt1(o.v[pm['entry_id'] - 1])).trim();
        if (target && tid) { assign[tid] = target; picked++; }
      });
    }
  }

  // ③ 写回总表
  for (var r2 = 2; r2 <= n; r2++) {
    var id2 = String(fmt1(ws.Cells(r2, cId).Value2)).trim();
    if (id2 && assign[id2] !== undefined) {
      ws.Cells(r2, cAssign).Value2 = assign[id2];
      colorCell(ws, r2, cAssign, assign[id2]);
    }
  }

  // ④ 重建分表与淘汰池
  var all = readRows(ws);
  var byTeam = {}, pool = [];
  all.forEach(function (o) {
    var a = String(fmt1(o.v[cAssign - 1])).trim();
    var id = String(fmt1(o.v[cId - 1])).trim();
    if (!id) return;
    if (a === '淘汰') pool.push(o);
    else if (a) (byTeam[a] = byTeam[a] || []).push(o);
  });

  var rebuilt = 0;
  var teams = {};
  var tz = sheet('支队对照');
  if (tz) {
    var tzm = headerMap(tz);
    for (var t = 2; t <= lastRow(tz); t++) {
      var num = String(fmt1(tz.Cells(t, tzm['序号']).Value2)).trim();
      var nm = String(fmt1(tz.Cells(t, tzm['支队名（=分表名）']).Value2)).trim();
      if (num && nm) teams[num] = nm;
    }
  }

  Object.keys(teams).forEach(function (num) {
    var ts = sheet(teams[num]);
    if (!ts) return;
    var tm = headerMap(ts);
    var tn = lastRow(ts);
    if (tn >= 2) ts.Range(ts.Cells(2, 1), ts.Cells(tn, tm['entry_id'])).ClearContents();
    (byTeam[num] || []).forEach(function (o, i) {
      var vals = pickByHeader(o.v, m, tm, 1);          // 跳过「录取状态」
      ts.Range(ts.Cells(i + 2, 2), ts.Cells(i + 2, 2 + vals.length - 1)).Value2 = [vals];
      var oid = String(fmt1(o.v[m['entry_id'] - 1])).trim();
      ts.Cells(i + 2, 1).Value2 = statusOf[oid] || '';  // 恢复「录取」等标记
    });
    rebuilt++;
  });

  if (ps) {
    var pm2 = headerMap(ps);
    var pn = lastRow(ps);
    if (pn >= 2) ps.Range(ps.Cells(2, 1), ps.Cells(pn, pm2['entry_id'])).ClearContents();
    pool.forEach(function (o, i) {
      var vals = pickByHeader(o.v, m, pm2, 2);          // 跳过 分配状态、捞入支队
      ps.Range(ps.Cells(i + 2, 3), ps.Cells(i + 2, 3 + vals.length - 1)).Value2 = [vals];
      ps.Cells(i + 2, 1).Value2 = '淘汰';
    });
  }

  return 'reflow: 淘汰=' + moved + ', 捞起=' + picked + ', 分表重建=' + rebuilt + ', 淘汰池=' + pool.length;
}

/* ---------------- 金数据拉取（金山定时任务调用） ---------------- */

function respJson(r, path) {
  if (!r) throw new Error('金数据接口无响应：' + path);
  var code = (r.status != null ? r.status : r.statusCode);
  var isOk = (code === 200 || code === 201 || code == null);
  if (!isOk) throw new Error('金数据接口 ' + path + ' → HTTP ' + code);
  if (typeof r.json === 'function') return r.json();          // AirScript Response
  var txt = (typeof r.text === 'function') ? r.text() : r.text; // 兼容 fetch 风格
  if (txt && typeof txt === 'object') return txt;
  return JSON.parse(String(txt));
}

function jsjGet(path) {
  var r = HTTP.get('https://jinshuju.net/api/v1' + path, {
    headers: { Authorization: 'Bearer ' + CONFIG.token, Accept: 'application/json' },
    timeout: 15000
  });
  return respJson(r, path);
}

function ledgerRowOf(e) {
  var g = function (k) {
    var v = e[k];
    if (Object.prototype.toString.call(v) === '[object Array]') return v.join('、');
    return v == null ? '' : v;
  };
  var sorted = (e.field_16 || []).slice().sort(function (a, b) { return (a.rank || 99) - (b.rank || 99); });
  var vols = sorted.map(function (it) { return String(it.choice || '').trim(); });   // 志愿列写文本
  var firstNum = '';
  if (sorted.length) {
    var mm = String(sorted[0].choice || '').match(/^\s*(\d+)\./);
    if (mm) firstNum = mm[1];
  }
  return {
    id: e.token, sn: e.serial_number, assign: firstNum, vols: vols,
    name: g('field_1'), minority: g('field_2'), sid: g('field_3'), phone: g('field_4'),
    gender: g('field_5'), dept: g('field_18'), major: g('field_7'), grad: g('field_8'),
    degree: g('field_9'), bschool: g('field_10'), bmajor: g('field_11'), resume: g('field_12'),
    wantLeader: g('field_14'), hometown: g('field_15')
  };
}

function syncEntries() {
  if (!CONFIG.token || CONFIG.token.indexOf('<') === 0) throw new Error('CONFIG.token 未填写');
  var all = [], next = null, guard = 0;
  do {
    var q = next ? ('/forms/' + CONFIG.form + '/entries?next=' + encodeURIComponent(next) + '&include_answer=true')
                 : ('/forms/' + CONFIG.form + '/entries?limit=1000&include_answer=true');
    var d = jsjGet(q);
    (d.data || []).forEach(function (e) { all.push(e); });
    next = d.next; guard++;
  } while (next && guard < 20);
  var rows = all.map(ledgerRowOf);
  var msg = appendEntries(rows);
  return '同步 ' + rows.length + ' 条：' + msg;
}

function logRun(text) {
  var ws = ensureSheet('同步日志');
  if (String(fmt1(ws.Cells(1, 1).Value2)).trim() === '') {
    ws.Range('A1:B1').Value2 = [['时间', '结果']];
    ws.Rows(1).Font.Bold = true;
  }
  var r = lastRow(ws) + 1;
  var d = new Date();
  var p = function (n) { return (n < 10 ? '0' : '') + n; };
  var stamp = d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate()) + ' ' +
              p(d.getHours()) + ':' + p(d.getMinutes());
  ws.Range(ws.Cells(r, 1), ws.Cells(r, 2)).Value2 = [[stamp, String(text).slice(0, 400)]];
  if (r > 500) { try { ws.Range('A2:B' + (r - 400)).ClearContents(); } catch (e) {} }
}

/* 定时任务入口：建表(首次) + 拉答卷 + 归位 */
function main() {
  var out = [];
  try {
    if (CONFIG.teams && CONFIG.teams.length) out.push(ensureSheets(CONFIG.teams));
    out.push(syncEntries());
    out.push(reflow(true));
  } catch (e) {
    out.push('失败：' + (e && e.message ? e.message : e));
  }
  var msg = out.join(' | ');
  try { logRun(msg); } catch (e2) {}
  return msg;
}

/* ---------------- 入口 ---------------- */

if (action === 'ping') return 'pong';
if (action === 'ensure') {
  var tmsg = ensureSheets(argv.teams || []);
  var amsg = appendEntries(argv.rows || []);
  return tmsg + ' | ' + amsg;
}
if (action === 'appendEntries') return appendEntries(argv.rows || []);
if (action === 'reflow') return reflow();
if (action === 'main') return main();
return '未知 action: ' + action;
