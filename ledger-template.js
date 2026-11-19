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
 *   action = 'reflow'          归位：分表「淘汰」→ 调剂池；调剂池「捞入支队」→ 回分表
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
 *   调剂池         分配状态 | 捞入支队 | 序号 | 姓名 | … | entry_id
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

/* 手动开关：改成 true 跑一次 = 强制清空重建（跑完记得改回 false） */
var FORCE_RESET = false;

var argv = (typeof Context !== 'undefined' && Context && Context.argv) || {};
var action = argv.action || 'ping';

/* 浅色系：浅底 + 深字，打印/截图都清楚 */
var COLORS = [[220, 233, 249], [223, 243, 228], [255, 243, 209], [253, 226, 218], [232, 226, 250],
              [217, 242, 242], [251, 225, 238], [252, 235, 214], [233, 237, 242], [237, 245, 218],
              [240, 230, 221], [230, 232, 235]];
var TEXT_DARK = [55, 65, 81];
var HEAD_FILL = [234, 241, 251];
var HEAD_TEXT = [30, 58, 95];

var BASE_HEAD = ['序号', '姓名', '少干/定向', '学号', '手机号', '性别', '院系', '专业',
                 '拟毕业时间', '学历', '本科院校', '本科专业', '简历链接', '愿任支队长', '家乡'];

/* ---------------- 基础工具 ---------------- */

function sheet(name) {
  try { return Application.Sheets(name); } catch (e) { return null; }
}

function lastDataRow(ws) {
  var n = lastRow(ws);
  var ncol = Math.min(ws.UsedRange.Columns.Count, 30);
  for (var r = n; r >= 2; r--) {
    var has = false;
    for (var c = 1; c <= ncol; c++) {
      if (String(fmt1(ws.Cells(r, c).Value2)).trim() !== '') { has = true; break; }
    }
    if (has) return r;
  }
  return 1;
}

function lastRow(ws) {
  try {
    var r = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1;
    return r > 0 ? r : 1;
  } catch (e) { return 1; }
}

function headerMap(ws) {
  var cols = ws.UsedRange.Columns.Count;
  var map = {};
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
  ws.Rows(1).Interior.Color = RGB(HEAD_FILL[0], HEAD_FILL[1], HEAD_FILL[2]);
  ws.Rows(1).Font.Color = RGB(HEAD_TEXT[0], HEAD_TEXT[1], HEAD_TEXT[2]);
}

function colLetter(n) {
  var s2 = '';
  while (n > 0) { var m = (n - 1) % 26; s2 = String.fromCharCode(65 + m) + s2; n = Math.floor((n - 1) / 26); }
  return s2;
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
  var ncol = Math.max(ws.UsedRange.Columns.Count, headerMap(ws)['entry_id'] || 0);
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

/* 条件格式：按公式整行着色（失败静默跳过） */
/* 直接给总表 A 列涂色（不依赖条件格式，最稳） */
function paintAssignCol() {
  var ws = sheet('总表');
  if (!ws) return;
  var m = headerMap(ws);
  var cA = m['分配状态'] || 1;
  var cmap = {};
  (CONFIG.teams || []).forEach(function (t, i) {
    cmap[t.label] = COLORS[i % COLORS.length];
  });
  var n2 = lastRow(ws);
  if (n2 < 2) return;
  var colStr = colLetter(cA);
  try { ws.Range(colStr + '2:' + colStr + n2).Interior.ColorIndex = -4142; } catch (e0) {}
  // 逐队找到连续区间整块涂色（每队通常只有几段）
  Object.keys(cmap).forEach(function (lbl) {
    var p = cmap[lbl], start = 0;
    for (var r = 2; r <= n2 + 1; r++) {
      var hit = r <= n2 && String(fmt1(ws.Cells(r, cA).Value2)).trim() === lbl;
      if (hit && !start) start = r;
      if (!hit && start) {
        try {
          var rng = ws.Range(ws.Cells(start, cA), ws.Cells(r - 1, cA));
          rng.Interior.Color = RGB(p[0], p[1], p[2]);
          rng.Font.Color = RGB(TEXT_DARK[0], TEXT_DARK[1], TEXT_DARK[2]);
        } catch (e1) {}
        start = 0;
      }
    }
  });
}

/* 整列条件格式：单元格文本 = text 时整格上色；同一工作表只建一次（防重复规则） */
var _cfDone = {};
function addColColor(ws, colLetter, text, rgbArr) {
  var key = String(ws.Name) + '|' + colLetter + '|' + String(text);
  if (_cfDone[key]) return true;
  try {
    var rng = ws.Range(colLetter + ':' + colLetter);
    var fc = rng.FormatConditions;
    var c = null;
    try { c = fc.Add(1, 2, '"' + text + '"'); } catch (e1) {}          // 文本=等于
    if (!c) { try { c = fc.AddText(text); } catch (e2) {} }
    if (!c) { try { c = fc.Add(2, 0, '=' + colLetter + '1="' + text + '"'); } catch (e3) {} }
    if (!c) return false;
    c.Interior.Color = RGB(rgbArr[0], rgbArr[1], rgbArr[2]);
    try { c.Font.Color = RGB(55, 65, 81); } catch (e4) {}
    _cfDone[key] = true;
    return true;
  } catch (e) { return false; }
}

/* 旧函数保留为别名（整列 + 文本匹配） */
function addRowColor(ws, addr, formula, rgbArr) {
  var col = String(addr).split('2:')[0].split('1:')[0].replace(/[0-9]/g, '');
  var m = String(formula).match(/"(.+)"/);
  var text = m ? m[1] : String(formula).replace(/[=$A-Z0-9]/g, '');
  return addColColor(ws, col, text, rgbArr);
}

/* 下拉：优先内联（手机端可正常显示），超长才退回区域引用 */
function addSmartList(ws, addr, items, refFallback) {
  var joined = items.join(',');
  if (joined.length <= 240) return addList(ws, addr, items);
  if (refFallback) return addRangeList(ws, addr, refFallback);
  return false;
}

/* 下拉：引用某区域的文字列表（不受 255 字符限制） */
function addRangeList(ws, addr, ref) {
  try {
    var rng = ws.Range(addr);
    try { rng.Validation.Delete(); } catch (e0) {}
    rng.Validation.Add(3, 1, 1, '=' + ref);
    return true;
  } catch (e) { return false; }
}

function colorCell(ws, r, c, num) {
  var kk = Number(num);
  if (!isFinite(kk) || kk <= 0) return;     // 队名文本不手动着色（交给条件格式）
  var k = (kk - 1) % COLORS.length;
  var p = COLORS[k];
  if (!p) return;
  ws.Cells(r, c).Interior.Color = RGB(p[0], p[1], p[2]);
  ws.Cells(r, c).Font.Color = RGB(TEXT_DARK[0], TEXT_DARK[1], TEXT_DARK[2]);
}

/* ---------------- ensure：建表 ---------------- */

function ensureSheets(teams) {
  var n = teams.length;

  // 0. 选项（隐藏）：下拉数据源，避免内联字符串的引号/空格问题
  var opt = ensureSheet('选项');
  opt.Range('A1:C100').ClearContents();
  opt.Range('A1:C1').Value2 = [['录取状态', '支队名', '分配状态']];
  opt.Range('A2:A3').Value2 = [['录取'], ['淘汰']];
  for (var oi = 1; oi <= n; oi++) {
    opt.Cells(1 + oi, 2).Value2 = teams[oi - 1].label;
    opt.Cells(1 + oi, 3).Value2 = teams[oi - 1].label;
  }
  opt.Cells(n + 2, 3).Value2 = '淘汰';
  try { opt.Columns('A:C').Hidden = false; } catch (e0) {}       // 列不要隐藏（否则看起来是空表）
  try { opt.Visible = true; } catch (e1) {}                       // 保持可见：超长列表回退引用时要能读

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
    ws.Cells(r, 4).Formula = '=COUNTIF(总表!$A:$A,B' + r + ')';
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
  // 联络需求表留空（只建表头，内容人工补）
  if (lastRow(ws) >= 2) {
    try { ws.Range(ws.Cells(2, 1), ws.Cells(1000, 7)).ClearContents(); } catch (e0) {}
  }

  // 4. 开函统计
  ws = ensureSheet('开函统计');
  writeHeader(ws, ['批次', '序号', '院系', '单位名称', '对方单位是否审核完毕', '是否上传网盘',
                   '联系人', '联系方式', '上传时间', '已盖章', '院系接受后备注', '备注']);
  setWidths(ws, [8, 6, 24, 36, 14, 12, 10, 14, 12, 8, 18, 16]);
  // 开函统计留空（只建表头）
  if (lastRow(ws) >= 2) {
    try { ws.Range(ws.Cells(2, 1), ws.Cells(1000, 12)).ClearContents(); } catch (e1) {}
  }

  // 5. 总表
  ws = ensureSheet('总表');
  var head = ['分配状态'].concat(BASE_HEAD).concat(['录取状态']);
  for (var v = 1; v <= n; v++) head.push('志愿' + v);
  head.push('entry_id');
  writeHeader(ws, head);
  var widths = [34, 6, 10, 10, 14, 14, 6, 14, 18, 12, 8, 14, 18, 30, 10, 12];
  for (var x = 0; x < n; x++) widths.push(26);
  widths.push(12);
  setWidths(ws, widths);
  try { ws.Columns(head.length).Hidden = true; } catch (e) {}   // 隐藏 entry_id
  var assignOpts = teams.map(function (t) { return String(t.label); }).concat(['淘汰']);
  addSmartList(ws, 'A2:A1000', assignOpts, '选项!$C$2:$C$' + (n + 2));   // 分配状态下拉（内联优先）
  try { ws.Range('A2:A1000').FormatConditions.Delete(); } catch (e1) {}
  for (var ci = 1; ci <= n && ci <= COLORS.length; ci++) {
    var lbl = String(teams[ci - 1].label).replace(/"/g, '');
    addColColor(ws, 'A', lbl, COLORS[ci - 1]);                            // 整列按队名匹配
  }

  // 6. 调剂池
  ws = ensureSheet('调剂池');
  writeHeader(ws, ['分配状态', '捞入支队'].concat(BASE_HEAD).concat(['志愿1','志愿2','志愿3','志愿4','志愿5']).concat(['entry_id']));
  setWidths(ws, [10, 34, 6, 10, 10, 14, 14, 6, 14, 18, 12, 8, 14, 18, 30, 10, 12, 12]);
  var teamOpts = teams.map(function (t) { return String(t.label); });
  addSmartList(ws, 'B2:B1000', teamOpts, '选项!$B$2:$B$' + (n + 1));    // 捞入支队（内联优先）

  // 7. 各支队分表
  teams.forEach(function (t) {
    var name = String(t.label).substring(0, 31);
    var ts = ensureSheet(name);
    if (lastRow(ts) < 1 || String(fmt1(ts.Cells(1, 1).Value2)).trim() === '') {
      writeHeader(ts, ['录取状态'].concat(BASE_HEAD).concat(['志愿1','志愿2','志愿3','志愿4','志愿5']).concat(['entry_id']));
      setWidths(ts, [10, 6, 10, 10, 14, 14, 6, 14, 18, 12, 8, 14, 18, 30, 10, 12, 12]);
    }
    addList(ts, 'A2:A1000', ['录取', '淘汰']);                       // 录取状态下拉（内联）
    try { ts.Range('A2:A1000').FormatConditions.Delete(); } catch (e2) {}
    addColColor(ts, 'A', '淘汰', [237, 237, 237]);
    addColColor(ts, 'A', '录取', [214, 242, 222]);                      // 整列按状态匹配
  });

  return 'sheets ready: ' + (6 + teams.length) + ' 张';
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

  var nextNew = Object.keys(idx).length ? (lastDataRow(ws) + 1) : 2;
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
      rr = nextNew; nextNew++;
      ws.Cells(rr, cAssign).Value2 = row.assign || '';
      // 颜色由条件格式按队名自动渲染
      ws.Range(ws.Cells(rr, m['序号']), ws.Cells(rr, m['序号'] + b.length - 1)).Value2 = [b];
      if (volCols.length) {
        ws.Range(ws.Cells(rr, volCols[0]), ws.Cells(rr, volCols[volCols.length - 1])).Value2 = [vols];
      }
      ws.Cells(rr, cId).Value2 = row.id;
      idx[row.id] = rr;
      added++;
    }
  });

  paintAssignCol();
  return 'added=' + added + ', updated=' + updated + ', total=' + (lastRow(ws) - 1);
}

/* 整块重写总表（1 次写入，快）
   assignMap / statusMap：按 entry_id 保留的人工结果 */
function rebuildTotal(rows, assignMap, statusMap) {
  var ws = sheet('总表');
  if (!ws) throw new Error('找不到「总表」工作表');
  var m = headerMap(ws);
  var cId = m['entry_id'];
  if (!cId) throw new Error('总表缺少 entry_id 列');
  var width = cId;
  var volMax = 0;
  for (var v = 1; m['志愿' + v]; v++) volMax = v;

  // 先取旧的人工结果（避免整块重写时丢掉）
  var oldAssign = {}, oldStatus = {};
  var n = lastRow(ws);
  for (var r = 2; r <= n; r++) {
    var oid = String(fmt1(ws.Cells(r, cId).Value2)).trim();
    if (!oid) continue;
    var a = String(fmt1(ws.Cells(r, m['分配状态']).Value2)).trim();
    var st = m['录取状态'] ? String(fmt1(ws.Cells(r, m['录取状态']).Value2)).trim() : '';
    if (a) oldAssign[oid] = a;
    if (st) oldStatus[oid] = st;
  }

  var mat = [];
  rows.forEach(function (row) {
    if (!row || !row.id) return;
    var line = [];
    for (var k = 0; k < width; k++) line.push('');
    line[m['分配状态'] - 1] = (assignMap && assignMap[row.id]) || oldAssign[row.id] || row.assign || '';
    var keys = ['sn'].concat(BASE_HEAD.slice(1).map(function (h) { return HEAD2KEY[h]; }));
    var cols = ['序号'].concat(BASE_HEAD.slice(1));
    // 序号
    line[m['序号'] - 1] = row.sn == null ? '' : row.sn;
    // 其余基础字段
    BASE_HEAD.slice(1).forEach(function (h, i2) {
      var col = m[h];
      if (!col) return;
      var kk = HEAD2KEY[h];
      line[col - 1] = row[kk] == null ? '' : row[kk];
    });
    // 志愿
    var vols = row.vols || [];
    for (var v2 = 1; v2 <= volMax; v2++) {
      var cv = m['志愿' + v2];
      if (cv) line[cv - 1] = vols[v2 - 1] == null ? '' : vols[v2 - 1];
    }
    // 录取状态
    if (m['录取状态']) line[m['录取状态'] - 1] = (statusMap && statusMap[row.id]) || oldStatus[row.id] || '';
    // entry_id
    line[cId - 1] = row.id;
    mat.push(line);
  });

  var tn = lastRow(ws);
  if (tn >= 2) { try { ws.Range(ws.Cells(2, 1), ws.Cells(tn, width)).ClearContents(); } catch (eC) {} }
  if (mat.length) {
    ws.Range(ws.Cells(2, 1), ws.Cells(1 + mat.length, width)).Value2 = mat;
  }
  return 'added=' + mat.length + ', total=' + mat.length;
}

/* 字段名 → 答卷行里的键 */
var HEAD2KEY = {
  '序号': 'sn', '姓名': 'name', '少干/定向': 'minority', '学号': 'sid', '手机号': 'phone',
  '性别': 'gender', '院系': 'dept', '专业': 'major', '拟毕业时间': 'grad', '学历': 'degree',
  '本科院校': 'bschool', '本科专业': 'bmajor', '简历链接': 'resume', '愿任支队长': 'wantLeader', '家乡': 'hometown'
};

/* 比较「目标二维数组」与表内现有内容是否一致（字符串化比较，忽略类型差） */
function sameMatrix(ws, ncol, desired) {
  var n = lastRow(ws);
  var cur = n >= 2 ? ws.Range(ws.Cells(2, 1), ws.Cells(n, ncol)).Value2 : [];
  if ((cur || []).length !== desired.length) return false;
  for (var i = 0; i < desired.length; i++) {
    for (var j = 0; j < desired[i].length; j++) {
      var a = cur[i] ? String(fmt1(cur[i][j])).trim() : '';
      var b = String(fmt1(desired[i][j])).trim();
      if (a !== b) return false;
    }
  }
  return true;
}

/* 清空所有数据行（保留表头/样式/下拉） */
function clearAllData() {
  var names = ['总表', '调剂池'];
  for (var i = 1; i <= Application.Sheets.Count; i++) {
    var nm = Application.Sheets.Item(i).Name;
    if (names.indexOf(nm) < 0) names.push(nm);
  }
  var skip = ['支队对照', '支队概况总览', '企业联络需求表', '开函统计', '选项', '同步日志'];
  var cleared = 0;
  names.forEach(function (nm) {
    if (skip.indexOf(nm) >= 0) return;
    var ws = sheet(nm);
    if (!ws) return;
    // 支队分表（第一列表头是「录取状态」）不在这里清，交给 reflow 逐表重写
    try { if (String(fmt1(ws.Cells(1, 1).Value2)).trim() === '录取状态') return; } catch (eS) {}
    var n = lastDataRow(ws);
    if (n >= 2) {
      try {
        var rng = ws.Range(ws.Cells(2, 1), ws.Cells(n, Math.max(ws.UsedRange.Columns.Count, 20)));
        rng.ClearContents();
        cleared++;
      } catch (e) {
        try { ws.Range(ws.Cells(2, 1), ws.Cells(n, 20)).ClearContents(); cleared++; } catch (e2) {}
      }
    }
  });
  return cleared;
}

/* 表头是否为新版（有 录取状态 列） */
function headerOk() {
  var ws = sheet('总表');
  if (!ws) return false;
  var m = headerMap(ws);
  return !!m['录取状态'] && !!m['entry_id'];
}

/* 清空重建：保留每人的录取状态与分配，重建后恢复 */
function resetAll(hard, preRows, noReflow) {
  var ws = sheet('总表');
  if (ws && headerOk()) { try { reflow(true); } catch (e00) {} }   // 先把分表里的录取状态收回总表

  var keep = {};
  if (ws && headerOk()) {
    var m = headerMap(ws), cId = m['entry_id'];
    var st = m['录取状态'], as = m['分配状态'];
    for (var r = 2; r <= lastRow(ws); r++) {
      var id = String(fmt1(ws.Cells(r, cId).Value2)).trim();
      if (!id) continue;
      keep[id] = {
        status: String(fmt1(ws.Cells(r, st).Value2)).trim(),
        assign: String(fmt1(ws.Cells(r, as).Value2)).trim()
      };
    }
  }
  var cleared = 0;
  if (hard) {
    // 硬重建：直接删掉数据表（总表/调剂池/选项/各支队分表），再整表新建
    var del = ['总表', '调剂池', '选项'];
    var tz0 = sheet('支队对照');
    if (tz0) {
      var tzm0 = headerMap(tz0);
      for (var t0 = 2; t0 <= lastRow(tz0); t0++) {
        var nm0 = String(fmt1(tz0.Cells(t0, tzm0['支队名（=分表名）']).Value2)).trim();
        if (nm0) del.push(nm0);
      }
    }
    del.forEach(function (nmD) {
      var sd = sheet(nmD);
      if (!sd) return;
      try { sd.Delete(); cleared++; } catch (eD) {}
    });
    // 支队对照里的队名如果还没有（首次），用 CONFIG.teams 兜底
    if (!tz0 && CONFIG.teams) {
      CONFIG.teams.forEach(function (tD) {
        var sD = sheet(String(tD.label).substring(0, 31));
        if (sD) { try { sD.Delete(); cleared++; } catch (eD2) {} }
      });
    }
  } else {
    cleared = clearAllData();
  }
  var needEnsure = hard || !headerOk();
  var msg = needEnsure ? ensureSheets(CONFIG.teams || []) : 'sheets kept';
  var sync = preRows ? applyRows(preRows) : syncEntries();
  var ws2 = sheet('总表');
  if (ws2 && Object.keys(keep).length) {
    var m2 = headerMap(ws2), cId2 = m2['entry_id'], st2 = m2['录取状态'], as2 = m2['分配状态'];
    for (var r2 = 2; r2 <= lastRow(ws2); r2++) {
      var id2 = String(fmt1(ws2.Cells(r2, cId2).Value2)).trim();
      var k = keep[id2];
      if (!k) continue;
      if (k.status) ws2.Cells(r2, st2).Value2 = k.status;
      if (k.assign) { ws2.Cells(r2, as2).Value2 = k.assign; }
    }
  }
  var rf = '';
  if (!noReflow) {
    try { rf = ' | ' + reflow(true, true); } catch (eR) { rf = ' | reflow失败：' + (eR && eR.message ? eR.message : eR); }
  }
  try { paintAssignCol(); } catch (eP) {}
  return 'reset: 清空 ' + cleared + ' 张表 | ' + msg + ' | ' + sync + rf;
}

/* ---------------- reflow：淘汰 / 捞起归位 ---------------- */

function reflow(force, alwaysWrite) {
  // 防抖：自动化可能对每次编辑都触发，30 秒内重复调用直接跳过（main 里传 true 强制）
  try {
    var lock = ensureSheet('同步日志');
    var last = Number(lock.Cells(1, 4).Value2 || 0);
    var now = new Date().getTime();
    if (!force && last && now - last < 20000) return 'reflow skipped (20s 防抖)';
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
    if (name === '总表' || name === '调剂池' || name === '支队对照' || name === '支队概况总览' ||
        name === '企业联络需求表' || name === '开函统计') return;
    var ts = sheet(name);
    if (!ts) return;
    var tm = headerMap(ts);
    if (!tm['entry_id'] || !tm['录取状态']) return;
    readRows(ts).forEach(function (o) {
      var status = String(fmt1(o.v[0])).trim();
      var tid = String(fmt1(o.v[tm['entry_id'] - 1])).trim();
      if (tid && status === '淘汰') { assign[tid] = '淘汰'; statusOf[tid] = '淘汰'; moved++; }
      else if (tid && status) { statusOf[tid] = status; }
    });
  });

  // ② 调剂池「捞入支队」
  var ps = sheet('调剂池');
  var picked = 0;
  if (ps) {
    var pm = headerMap(ps);
    if (pm['entry_id'] && pm['捞入支队']) {
      readRows(ps).forEach(function (o) {
        var raw = String(fmt1(o.v[pm['捞入支队'] - 1])).trim();
        var tid = String(fmt1(o.v[pm['entry_id'] - 1])).trim();
        if (raw && tid) { assign[tid] = raw; picked++; }   // 直接存队名
      });
    }
  }

  // ③ 写回总表（分配状态 + 录取状态）
  var cStatus = m['录取状态'];
  if (n >= 2) {
    var ids = ws.Range(ws.Cells(2, cId), ws.Cells(n, cId)).Value2 || [];
    var aCol = [], sCol = [], changed = false;
    for (var r2 = 0; r2 < ids.length; r2++) {
      var id2 = String(fmt1(ids[r2] ? (ids[r2][0] !== undefined ? ids[r2][0] : ids[r2]) : '')).trim();
      var av = (id2 && assign[id2] !== undefined) ? assign[id2] : '';
      var sv = (id2 && statusOf[id2] !== undefined) ? statusOf[id2] : '';
      aCol.push([av]); sCol.push([sv]);
      if (av || sv) changed = true;
    }
    if (changed) {
      try { ws.Range(ws.Cells(2, cAssign), ws.Cells(1 + ids.length, cAssign)).Value2 = aCol; } catch (eA1) {}
      if (cStatus) { try { ws.Range(ws.Cells(2, cStatus), ws.Cells(1 + ids.length, cStatus)).Value2 = sCol; } catch (eA2) {} }
    }
  }

  // ④ 重建分表与调剂池
  // 先读队名列表（支队对照 B 列），用于把 A 列值映射到分表名
  var teamNames = [];
  var tz0 = sheet('支队对照');
  if (tz0) {
    var tzm0 = headerMap(tz0);
    for (var t0 = 2; t0 <= lastRow(tz0); t0++) {
      var nm0 = String(fmt1(tz0.Cells(t0, tzm0['支队名（=分表名）']).Value2)).trim();
      if (nm0) teamNames.push(nm0);
    }
  }

  var all = readRows(ws);
  var byTeam = {}, pool = [];
  all.forEach(function (o) {
    var a = String(fmt1(o.v[cAssign - 1])).trim();
    var id = String(fmt1(o.v[cId - 1])).trim();
    if (!id) return;
    if (a === '淘汰') { pool.push(o); return; }
    if (!a) return;
    var key = a;
    if (teamNames.indexOf(key) < 0) {
      // 兼容：A 列存序号或部分文本 → 匹配到完整队名
      for (var t2 = 0; t2 < teamNames.length; t2++) {
        var lbl = teamNames[t2];
        if (lbl === a || String(lbl).indexOf(a) === 0 || String(a).indexOf(String(lbl).split(' ')[0]) === 0) { key = lbl; break; }
      }
    }
    (byTeam[key] = byTeam[key] || []).push(o);
  });

  var rebuilt = 0;
  var dbg = [];
  Object.keys(byTeam).forEach(function (k) { dbg.push(String(k).slice(0, 14) + ':' + byTeam[k].length); });
  var dbgTxt = '分表分布: ' + (dbg.length ? dbg.join(' ') : '（无）');

  /* 队名 → 实际 sheet 名（容忍空格/截断/序号前缀差异） */
  function resolveTeamSheet(key) {
    var k = String(key).replace(/\s+/g, '');
    for (var i = 0; i < teamNames.length; i++) {
      var n = String(teamNames[i]).replace(/\s+/g, '');
      if (n === k) return teamNames[i];
      var p = n.match(/^(\d+)\./), q = k.match(/^(\d+)\./);
      if (p && q && p[1] === q[1]) return teamNames[i];
    }
    return null;
  }

  var wroteKeys = {};
  var writeFail = 0, lastErr = '';
  Object.keys(byTeam).forEach(function (key) {
    var nm = resolveTeamSheet(key);
    if (!nm) { console.log('未匹配到分表: ' + key); return; }
    wroteKeys[nm] = 1;
    var ts = sheet(nm);
    if (!ts) { return; }
    var tm = headerMap(ts);
    var desired = byTeam[key].map(function (o) {
      var oid = String(fmt1(o.v[m['entry_id'] - 1])).trim();
      return [statusOf[oid] || ''].concat(pickByHeader(o.v, m, tm, 1));
    });
    var tn = lastRow(ts);
    var ncolT = Math.max(tm['entry_id'] || 1, 17);
    if (tn >= 2) { try { ts.Range(ts.Cells(2, 1), ts.Cells(tn, ncolT)).ClearContents(); } catch (eC) {} }
    if (desired.length) {
      var width = ncolT;
      var mat = desired.map(function (r) {
        var line = [];
        for (var k = 0; k < width; k++) line.push(k < r.length ? r[k] : '');
        return line;
      });
      try { ts.Range(ts.Cells(2, 1), ts.Cells(1 + mat.length, width)).Value2 = mat; }
      catch (eW) { writeFail++; lastErr = (eW && eW.message) ? eW.message : String(eW); }
    }
    rebuilt++;
  });

  // 没有人的分表：清空数据行
  teamNames.forEach(function (nm) {
    if (wroteKeys[nm]) return;
    var ts = sheet(nm);
    if (!ts) return;
    var tn = lastRow(ts);
    if (tn >= 2) { try { ts.Range(ts.Cells(2, 1), ts.Cells(tn, 17)).ClearContents(); } catch (eC2) {} }
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

  paintAssignCol();
  var rowsTxt = [];
  teamNames.forEach(function (nm) {
    var ts = sheet(nm);
    if (ts) rowsTxt.push(String(nm).slice(0, 4) + (lastRow(ts) - 1));
  });
  return 'reflow: 淘汰=' + moved + ', 捞起=' + picked + ', 分表=' + rowsTxt.join(',') +
         ', 调剂池=' + pool.length + ' | ' + dbgTxt +
         (writeFail ? (' | 写入失败' + writeFail + '次: ' + lastErr) : '');
}

/* ---------------- 金数据拉取（金山定时任务调用） ---------------- */

function respJson(r, path) {
  if (!r) throw new Error('金数据接口无响应：' + path);
  var code = (r.status != null ? r.status : r.statusCode);
  var isOk = (code === 200 || code === 201 || code == null);
  if (code === 429) throw new Error('金数据限流（免费版 100 次/小时），请稍后再试');
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
  var firstLabel = sorted.length ? String(sorted[0].choice || '').trim() : '';
  return {
    id: e.token, sn: e.serial_number, assign: firstLabel, vols: vols,
    name: g('field_1'), minority: g('field_2'), sid: g('field_3'), phone: g('field_4'),
    gender: g('field_5'), dept: g('field_18'), major: g('field_7'), grad: g('field_8'),
    degree: g('field_9'), bschool: g('field_10'), bmajor: g('field_11'), resume: g('field_12'),
    wantLeader: g('field_14'), hometown: g('field_15')
  };
}

/* 只做网络请求，不写任何单元格 */
function fetchRows() {
  if (!CONFIG.token || CONFIG.token.indexOf('<') === 0) throw new Error('CONFIG.token 未填写');
  var all = [], next = null, guard = 0;
  do {
    var q = next ? ('/forms/' + CONFIG.form + '/entries?next=' + encodeURIComponent(next) + '&include_answer=true')
                 : ('/forms/' + CONFIG.form + '/entries?limit=1000&include_answer=true');
    var d = jsjGet(q);
    (d.data || []).forEach(function (e) { all.push(e); });
    next = d.next; guard++;
  } while (next && guard < 20);
  return all.map(ledgerRowOf);
}

/* 把行写进总表（含删除核对） */
function applyRows(rows) {

  // 核对：金数据里已删除的答卷 → 从总表移除
  var ws = sheet('总表');
  var m = headerMap(ws), cId = m['entry_id'];
  var wanted = {};
  rows.forEach(function (r) { if (r.id) wanted[r.id] = true; });
  var removed = 0;
  if (ws && cId) {
    var toDel = [];
    for (var r0 = 2; r0 <= lastRow(ws); r0++) {
      var id0 = String(fmt1(ws.Cells(r0, cId).Value2)).trim();
      if (id0 && !wanted[id0]) toDel.push(r0);
    }
    toDel.reverse().forEach(function (rr) {
      try { ws.Rows(rr).Delete(); removed++; } catch (e) {}
    });
  }

  var msg = rebuildTotal(rows);
  return '同步 ' + rows.length + ' 条：' + msg + (removed ? '；移除已删除 ' + removed : '');
}

/* 只同步总表（不归位），用于把长任务拆成两步 */
function fetchOnly() {
  var T0 = new Date().getTime();
  var rows;
  try { rows = fetchRows(); }
  catch (eNet) {
    var em = '失败（未改动表格）：' + (eNet && eNet.message ? eNet.message : eNet);
    try { logRun(em); } catch (e) {}
    return em;
  }
  var msg;
  try { msg = resetAll(FORCE_RESET, rows, true); }
  catch (e2) { msg = '写入异常：' + (e2 && e2.message ? e2.message : e2); }
  msg += ' | 共 ' + rows.length + ' 条 | 耗时 ' + ((new Date().getTime() - T0) / 1000).toFixed(1) + 's';
  try { logRun('fetchOnly: ' + msg); } catch (e3) {}
  return msg;
}

function syncEntries() {
  return applyRows(fetchRows());
}

function logRun(text) {
  var ws = ensureSheet('同步日志');
  if (String(fmt1(ws.Cells(1, 1).Value2)).trim() === '') {
    ws.Range('A1:B1').Value2 = [['时间', '结果']];
    ws.Rows(1).Font.Bold = true;
  }
  try { ws.Columns(1).NumberFormat = '@'; } catch (eF) {}   // 时间列按文本存，避免变成 46345.53
  var r = lastDataRow(ws) + 1;   // 追加到真正最后一条之后，不覆盖
  var d = new Date();
  var p = function (n) { return (n < 10 ? '0' : '') + n; };
  var stamp = d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate()) + ' ' +
              p(d.getHours()) + ':' + p(d.getMinutes());
  ws.Range(ws.Cells(r, 1), ws.Cells(r, 2)).Value2 = [[stamp, String(text).slice(0, 1200)]];
  if (r > 500) { try { ws.Range('A2:B' + (r - 400)).ClearContents(); } catch (e) {} }
}

/* 定时任务入口：建表(首次) + 拉答卷 + 归位
   用法：定时任务 A：每 10 分钟跑 main()（同步金数据）
         定时任务 B：每小时每分（即每分钟）跑 main()（也含归位，因为 reflow 内部有 20s 防抖）
   说明：main() 里 reflow 是同步执行的，所以不需要单独的"每分钟归位"任务；
         金数据 API 限额 100 次/小时，main() 每分钟跑 = 60 次/小时，够用且安全。 */
function reflowCheck() {
  var ws = sheet('总表');
  if (!ws) return '无总表';
  var m = headerMap(ws), cId = m['entry_id'], cA = m['分配状态'];
  var total = 0, pool = 0;
  for (var r = 2; r <= lastRow(ws); r++) {
    var id = String(fmt1(ws.Cells(r, cId).Value2)).trim();
    if (!id) continue;
    total++;
    if (String(fmt1(ws.Cells(r, cA).Value2)).trim() === '淘汰') pool++;
  }
  var team = 0, names = [];
  var tz = sheet('支队对照');
  if (tz) {
    var tzm = headerMap(tz);
    for (var t = 2; t <= lastRow(tz); t++) {
      var nm = String(fmt1(tz.Cells(t, tzm['支队名（=分表名）']).Value2)).trim();
      if (nm) names.push(nm);
    }
  }
  names.forEach(function (nm) {
    var ts = sheet(nm);
    if (ts) team += (lastRow(ts) - 1);
  });
  var expect = total - pool;
  return team === expect ? ('一致 分表' + team + '=总表' + expect) : ('不一致 分表' + team + '≠总表' + expect);
}

function main() {
  var T0 = new Date().getTime();
  // 防抖：60 秒内重复执行直接跳过，避免把金数据 100 次/小时额度打爆
  try {
    var lb = ensureSheet('同步日志');
    var lastRun = Number(lb.Cells(1, 5).Value2 || 0);
    var nowRun = new Date().getTime();
    if (lastRun && nowRun - lastRun < 60000) return '跳过：60 秒内已同步过（防抖）';
    lb.Cells(1, 5).Value2 = nowRun;
  } catch (eD) {}

  var out = [];
  try {
    // ① 先拉数据：失败（如 429 限流）就直接返回，绝不改动表格
    var rows;
    try {
      rows = fetchRows();
    } catch (eNet) {
      var em = '失败（**未改动任何表格**）：' + (eNet && eNet.message ? eNet.message : eNet);
      try { logRun(em); } catch (eL) {}
      return em;
    }

    // ② 数据到手，再做重建/写入
    if (!headerOk()) {
      out.push(resetAll(true, rows));
    } else {
      out.push(resetAll(FORCE_RESET, rows));
    }
    // 核对：分表总人数应与总表非淘汰人数一致；不一致就补一次归位
    try {
      var chk = reflowCheck();
      out.push('核对: ' + chk);
    } catch (eChk) { out.push('核对失败:' + (eChk && eChk.message)); }
    out.push('共 ' + (rows || []).length + ' 条答卷');
    out.push('耗时 ' + ((new Date().getTime() - T0) / 1000).toFixed(1) + 's');
  } catch (e) {
    out.push('失败：' + (e && e.message ? e.message : e));
  }
  var msg = out.join(' | ');
  try { logRun(msg); } catch (e2) {}
  return msg;
}

/* ---------------- verify：只读体检，不写任何数据 ---------------- */
function verify() {
  var issues = [], ok = 0;
  var ws = sheet('总表');
  if (!ws) return '❌ 缺少「总表」';
  var m = headerMap(ws);
  var cId = m['entry_id'], cAssign = m['分配状态'];
  var seen = {}, rows = 0, dup = 0, noId = 0, badAssign = 0, assignCount = {};

  var teamNames = [];
  var tz = sheet('支队对照');
  if (tz) {
    var tzm = headerMap(tz);
    for (var t = 2; t <= lastRow(tz); t++) {
      var nm = String(fmt1(tz.Cells(t, tzm['支队名（=分表名）']).Value2)).trim();
      if (nm) teamNames.push(nm);
    }
  }
  var valid = {};
  teamNames.forEach(function (n2) { valid[n2] = true; });
  valid['淘汰'] = true; valid[''] = true;

  var nAll = lastRow(ws);
  var block = nAll >= 2 ? (ws.Range(ws.Cells(2, 1), ws.Cells(nAll, Math.max(cId, 1))).Value2 || []) : [];
  for (var r = 0; r < block.length; r++) {
    var line = block[r] || [];
    var id = String(fmt1(line[cId - 1])).trim();
    var a = String(fmt1(line[cAssign - 1])).trim();
    var any = false;
    for (var c = 0; c < line.length; c++) {
      if (String(fmt1(line[c])).trim() !== '') { any = true; break; }
    }
    if (!any) continue;
    rows++;
    if (!id) noId++;
    else if (seen[id]) dup++; else seen[id] = 1;
    if (a && !valid[a]) badAssign++;
    if (a) assignCount[a] = (assignCount[a] || 0) + 1;
  }

  // 分表人数与总表一致（淘汰的应去调剂池）
  var poolCount = 0;
  var ps = sheet('调剂池');
  if (ps) {
    var pn2 = lastRow(ps);
    if (pn2 >= 2) {
      var pblock = ps.Range(ps.Cells(2, 1), ps.Cells(pn2, 3)).Value2 || [];
      pblock.forEach(function (ln) {
        var has = false;
        for (var q = 0; q < (ln || []).length; q++) if (String(fmt1(ln[q])).trim() !== '') { has = true; break; }
        if (has) poolCount++;
      });
    }
  }
  var teamRows = 0;
  teamNames.forEach(function (nm) {
    var ts = sheet(nm);
    if (!ts) {
      // 兜底：sheet 名可能被截断/有空格差异 → 按「序号. 」前缀匹配
      var pref = String(nm).match(/^(\d+)\./);
      if (pref) {
        for (var si = 1; si <= Application.Sheets.Count; si++) {
          var sn = Application.Sheets.Item(si).Name;
          if (sn === nm) { ts = sheet(sn); break; }
          if (String(sn).indexOf(pref[1] + '.') === 0) { ts = sheet(sn); break; }
        }
      }
    }
    if (!ts) return;
    var trows = lastRow(ts);
    if (trows >= 2) {
      var tblock = ts.Range(ts.Cells(2, 1), ts.Cells(trows, 3)).Value2 || [];
      tblock.forEach(function (ln) {
        var has = false;
        for (var q2 = 0; q2 < (ln || []).length; q2++) if (String(fmt1(ln[q2])).trim() !== '') { has = true; break; }
        if (has) teamRows++;
      });
    }
  });

  var expectedTeam = rows - (assignCount['淘汰'] || 0);
  var expectedPool = assignCount['淘汰'] || 0;

  if (noId) issues.push('有 ' + noId + ' 行缺 entry_id');
  else ok++;
  if (dup) issues.push('entry_id 重复 ' + dup + ' 处');
  else ok++;
  if (badAssign) issues.push('分配状态非法 ' + badAssign + ' 行');
  else ok++;
  if (teamRows !== expectedTeam) issues.push('分表人数 ' + teamRows + ' ≠ 总表非淘汰 ' + expectedTeam);
  else ok++;
  if (poolCount !== expectedPool) issues.push('调剂池 ' + poolCount + ' ≠ 总表淘汰 ' + expectedPool);
  else ok++;
  if (!m['录取状态'] || !m['entry_id']) issues.push('表头缺列（录取状态/entry_id）');
  else ok++;

  return (issues.length ? '❌ ' : '✅ 体检通过 ') + ok + ' 项正常' +
         (issues.length ? '，问题：' + issues.join('；') : '（总表 ' + rows + ' 行，分表 ' + teamRows + '，调剂池 ' + poolCount + '）');
}

/* ---------------- debug：只读返回台账状态 ---------------- */
function debug() {
  var out = { sheets: [] };
  for (var i = 1; i <= Application.Sheets.Count; i++) out.sheets.push(Application.Sheets.Item(i).Name);

  var ws = sheet('总表');
  if (ws) {
    var m = headerMap(ws);
    out.totalHeader = Object.keys(m);
    out.totalRows = lastRow(ws) - 1;
    out.assign = {};
    for (var r = 2; r <= lastRow(ws); r++) {
      var a = String(fmt1(ws.Cells(r, m['分配状态']).Value2)).trim();
      if (a) out.assign[a] = (out.assign[a] || 0) + 1;
    }
    out.firstAssign = String(fmt1(ws.Cells(2, m['分配状态']).Value2)).trim();
    out.firstId = String(fmt1(ws.Cells(2, m['entry_id']).Value2)).trim();
  }
  var tz = sheet('支队对照');
  if (tz) {
    var tzm = headerMap(tz);
    out.teams = [];
    for (var t = 2; t <= lastRow(tz); t++) {
      var nm = String(fmt1(tz.Cells(t, tzm['支队名（=分表名）']).Value2)).trim();
      if (nm) {
        var ts = sheet(nm);
        out.teams.push({ name: nm, exists: !!ts, rows: ts ? lastRow(ts) - 1 : -1, header: ts ? Object.keys(headerMap(ts)).slice(0, 5) : [] });
      }
    }
  }
  return JSON.stringify(out);
}

/* ---------------- perms：只读检查工作表保护/可见性/筛选 ---------------- */
function perms() {
  var out = [];
  for (var i = 1; i <= Application.Sheets.Count; i++) {
    var ws = Application.Sheets.Item(i);
    var rec = { name: ws.Name };
    try { rec.visible = ws.Visible; } catch (e1) {}
    try { rec.protectContents = ws.ProtectContents; } catch (e2) {}
    try { rec.protectDrawing = ws.ProtectDrawingObjects; } catch (e3) {}
    try { rec.filterMode = ws.FilterMode; } catch (e4) {}
    try { rec.autoFilter = ws.AutoFilterMode; } catch (e5) {}
    out.push(rec);
  }
  return JSON.stringify(out);
}

/* ---------------- status：只读台账状态（不写数据） ---------------- */
function status() {
  var out = { sheets: [], teams: [], log: [] };
  for (var i = 1; i <= Application.Sheets.Count; i++) out.sheets.push(Application.Sheets.Item(i).Name);

  var ws = sheet('总表');
  if (ws) {
    var m = headerMap(ws);
    out.total = { rows: lastRow(ws) - 1, headers: Object.keys(m) };
    out.assign = {};
    for (var r = 2; r <= lastRow(ws); r++) {
      var a = String(fmt1(ws.Cells(r, m['分配状态'] || 1).Value2)).trim();
      if (a) out.assign[a] = (out.assign[a] || 0) + 1;
    }
  }

  var tz = sheet('支队对照');
  if (tz) {
    var tzm = headerMap(tz);
    for (var t = 2; t <= lastRow(tz); t++) {
      var nm = String(fmt1(tz.Cells(t, tzm['支队名（=分表名）']).Value2)).trim();
      if (!nm) continue;
      var ts = sheet(nm);
      out.teams.push({
        name: nm,
        exists: !!ts,
        rows: ts ? lastRow(ts) - 1 : -1,
        header: ts ? Object.keys(headerMap(ts)).slice(0, 6) : []
      });
    }
  }

  var ps = sheet('调剂池');
  if (ps) out.pool = { rows: lastRow(ps) - 1, headers: Object.keys(headerMap(ps)) };

  var lg = sheet('同步日志');
  if (lg) {
    var n = lastRow(lg);
    for (var l = Math.max(2, n - 4); l <= n; l++) {
      out.log.push([String(fmt1(lg.Cells(l, 1).Value2)), String(fmt1(lg.Cells(l, 2).Value2)).slice(0, 200)]);
    }
  }
  return JSON.stringify(out);
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
if (action === 'reset') return resetAll();
if (action === 'verify') return verify();
if (action === 'debug') return debug();
if (action === 'status') return status();
return '未知 action: ' + action;
