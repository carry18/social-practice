/**
 * 打款院系表 → 拨付说明（Word 下载 + 金山同步）
 * ------------------------------------------------------------------
 * 流程：从金山打款汇总表格读取数据（AirScript action=export）
 *       → 浏览器内用 docx.js 生成正式《经费拨付事宜的说明》Word 并下载
 *       → 同步调用 AirScript（action=syncNote）在同一工作簿新建
 *         「拨付说明」工作表存档。
 * 依赖：页面已加载 docx@8 UMD（window.docx）。
 * 本文件独立于 app.js / payout.js，所有 DOM id 带 note- 前缀。
 * ------------------------------------------------------------------
 */
(function () {
  'use strict';

  var DEFAULT_HOOK = 'https://1386267309-3bh6z757na.ap-shanghai.tencentscf.com';
  var $ = function (id) { return document.getElementById(id); };
  if (!$('note-gen')) return; // 页面没有该板块时静默退出

  /* ================= 代理调用 ================= */
  function callKdocs(argv, opts) {
    var payload = { Context: { argv: argv } };
    if (opts && opts.note) payload.target = 'note';   // 由代理路由到「拨付说明」文件的脚本
    return fetch(DEFAULT_HOOK, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    }).then(function (resp) {
      return resp.json().then(function (data) { return { ok: resp.ok, data: data }; });
    }).then(function (r) {
      if (!r.ok || r.data.error) throw new Error(r.data.error || ('HTTP ' + r.status));
      return r.data.data && r.data.data.result;
    });
  }

  /* ================= 工作表列表 ================= */
  function refreshList() {
    var btn = $('note-refresh');
    btn.disabled = true;
    callKdocs({ action: 'list' }).then(function (allNames) {
      var sel = $('note-sheet');
      var prev = sel.value;
      // 只保留「像打款表」的工作表：排除决算表、说明表、系统预留表
      var names = (allNames || []).filter(function (n) {
        return !/决算|说明|WpsReserved|Sheet\d*$/i.test(n);
      });
      sel.innerHTML = '';
      names.forEach(function (n) {
        var op = document.createElement('option');
        op.value = n; op.textContent = n;
        if (n === prev) op.selected = true;
        sel.appendChild(op);
      });
      if (!names.length) {
        var op0 = document.createElement('option');
        op0.value = ''; op0.textContent = '（未读到工作表，请点刷新）'; sel.appendChild(op0);
      }
      // 优先保留原选择；否则自动选中最后一个「打款表」
      if (prev && (names || []).indexOf(prev) >= 0) {
        sel.value = prev;
      } else {
        var hits = (names || []).filter(function (n) { return /打款/.test(n); });
        if (hits.length) sel.value = hits[hits.length - 1];
      }
      var hidden = (allNames || []).length - names.length;
      setStatus('✅ 打款表列表已刷新（' + names.length + ' 个可选用' +
        (hidden ? '，已忽略 ' + hidden + ' 个非打款表' : '') + '）', '#1a7f37');
    }).catch(function (e) {
      setStatus('❌ 获取工作表列表失败：' + e.message, '#b3261e');
    }).finally(function () { btn.disabled = false; });
  }
  $('note-refresh').onclick = refreshList;
  refreshList(); // 打开页面即自动拉一次

  /* ================= 文案拼装 ================= */
  /** 从工作表名解析年份学期：2026暑期院系打款表 → {y:'2026', tShort:'暑期', tFull:'暑假'} */
  function parseYearTerm(sheetName) {
    var ym = String(sheetName).match(/(20\d{2})/);
    var y = ym ? ym[1] : String(new Date().getFullYear());
    var s = String(sheetName);
    var tShort = /寒/.test(s) ? '寒假' : (/暑/.test(s) ? '暑期' : (/春/.test(s) ? '春季' : (/秋/.test(s) ? '秋季' : '寒假')));
    var tFull = tShort === '暑期' ? '暑假' : tShort;
    return { y: y, tShort: tShort, tFull: tFull };
  }

  function fmtAmount(n) {
    return Number.isInteger(n) ? String(n) : (Math.round(n * 100) / 100).toFixed(2);
  }

  function esc(s) {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }

  function setStatus(msg, color) {
    var el = $('note-status');
    el.innerHTML = msg;
    el.style.color = color || '';
  }

  /* ================= docx 生成 ================= */
  function buildDocx(meta, tableRows, totalStr) {
    var D = window.docx;
    var FANGSONG = { ascii: '仿宋', hAnsi: '仿宋', eastAsia: '仿宋' };       // 正文 仿宋（与样例一致）
    var CELLFONT = { ascii: '仿宋_GB2312', hAnsi: '仿宋_GB2312', eastAsia: '仿宋_GB2312' }; // 表格 仿宋_GB2312
    var TITLEFONT = { ascii: '方正小标宋_GBK', hAnsi: '方正小标宋_GBK', eastAsia: '方正小标宋_GBK' };

    var titlePara = function (text) {
      return new D.Paragraph({
        alignment: D.AlignmentType.CENTER,
        spacing: { line: 302, lineRule: D.LineRuleType.AUTO },
        children: [new D.TextRun({ text: text, font: TITLEFONT, size: 32 })] // 16pt
      });
    };
    var bodyPara = function (text, indent) {
      return new D.Paragraph({
        alignment: D.AlignmentType.JUSTIFIED,
        spacing: { before: 91, line: 303, lineRule: D.LineRuleType.AUTO }, // 样例：段前91缇/行距303
        indent: indent ? { firstLine: 550 } : undefined, // 首行缩进 2 字符（样例 550 缇）
        children: [new D.TextRun({ text: text, font: FANGSONG, size: 28, characterSpacing: -5 })] // 14pt
      });
    };

    // 明细表：列宽 dxa（1cm≈567）——与样例一致
    // 列宽：证号/金额两列较样例加宽，保证 10 位证号与带小数的金额不折行
    var COLW = [743, 1383, 1225, 880, 748, 950, 2009, 1128];
    var border = { style: D.BorderStyle.SINGLE, size: 2, color: '000000' };
    var cell = function (text, opts) {
      opts = opts || {};
      return new D.TableCell({
        width: { size: COLW[opts.col || 0], type: D.WidthType.DXA },
        columnSpan: opts.span,
        verticalMerge: opts.vMerge,
        verticalAlign: D.VerticalAlign.CENTER,
        children: [new D.Paragraph({
          alignment: D.AlignmentType.CENTER,
          spacing: { before: 237, line: 151, lineRule: D.LineRuleType.EXACT }, // 样例：段前237缇/行距151缇
          children: [new D.TextRun({ text: String(text), font: CELLFONT, size: 14, bold: true, characterSpacing: -6 })] // 7pt 加粗
        })]
      });
    };

    var HEADERS = ['编号', '地点', '打款院系', '打款金额', '项目负责人', '工作证号', '项目名称', '项目编号'];
    var mkRow = function (cells, isHeader) {
      return new D.TableRow({
        height: { value: 567, rule: D.HeightRule.EXACT }, // 样例行高 567 缇
        tableHeader: !!isHeader,                         // 跨页时自动重复表头
        children: cells
      });
    };
    var rows = [mkRow(HEADERS.map(function (h, i) { return cell(h, { col: i }); }), true)];
    tableRows.forEach(function (r, idx) {
      var isCont = !!r[8];
      var nextCont = idx + 1 < tableRows.length && !!tableRows[idx + 1][8];
      var cells = [];
      for (var i = 0; i < 8; i++) {
        var vMerge = null;
        if (i === 0 || i === 1) {                    // 编号、地点两列支持纵向合并
          if (isCont) vMerge = D.VerticalMergeType.CONTINUE;
          else if (nextCont) vMerge = D.VerticalMergeType.RESTART;
        }
        cells.push(cell(isCont && (i === 0 || i === 1) ? '' : r[i], { col: i, vMerge: vMerge }));
      }
      rows.push(mkRow(cells));
    });
    // 总计行：前 3 列合并「总计」+ 金额 + 后 4 列合并空白
    rows.push(mkRow([
      cell('总计', { col: 0, span: 3 }),
      cell(totalStr, { col: 3 }),
      cell('', { col: 4, span: 4 })
    ]));

    var table = new D.Table({
      width: { size: 9066, type: D.WidthType.DXA },
      columnWidths: COLW,
      layout: D.TableLayoutType.FIXED,
      alignment: D.AlignmentType.CENTER,
      borders: {
        top: border, bottom: border, left: border, right: border,
        insideHorizontal: border, insideVertical: border
      },
      rows: rows
    });

    return new D.Document({
      styles: { default: { document: { run: { font: { ascii: 'Arial', hAnsi: 'Arial', eastAsia: '仿宋_GB2312' }, size: 21 } } } },
      sections: [{
        properties: {
          page: {
            size: { width: 11906, height: 16838 }, // A4
            margin: { top: 1431, bottom: 0, left: 1476, right: 1355, header: 0, footer: 0 } // 与样例一致（下边距0，保证单页）
          }
        },
        children: [
          titlePara(meta.title1),
          titlePara(meta.title2),
          bodyPara(''),
          bodyPara(meta.body, true),
          bodyPara(''),
          bodyPara(meta.attach, false),
          bodyPara(''),
          table
        ]
      }]
    });
  }

  function downloadDocx(doc, filename) {
    return window.docx.Packer.toBlob(doc).then(function (blob) {
      var a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 4000);
    });
  }

  /* ================= 按院系汇总 ================= */
  var SELFPAY_RE2 = /自己报销|自行报销|自付|不拨付|自行处理/;

  /** 两级聚合：院系 → 支队；返回 [{dept,total,squads:[{place,amount,leader,cert,proj,projId,selfPay}]}] */
  function buildDeptRows(rows) {
    var map = {}, order = [];
    rows.forEach(function (r) {
      var dept = String(r.dept || '').trim() || '（未填院系）';
      var place = String(r.place || '').trim() || '（未填地点）';
      var amt = parseFloat(String(r.amount).replace(/[^\d.\-]/g, '')) || 0;
      var isSelf = SELFPAY_RE2.test(String(r.leader || '') + String(r.proj || ''));

      if (!map[dept]) { map[dept] = { dept: dept, total: 0, squads: {}, order: [] }; order.push(dept); }
      var g = map[dept];
      g.total += amt;

      if (!g.squads[place]) {
        g.squads[place] = { place: place, amount: 0, leader: '', cert: '', proj: '', projId: '', selfPay: false };
        g.order.push(place);
      }
      var sq = g.squads[place];
      sq.amount += amt;
      if (!sq.leader && r.leader) sq.leader = String(r.leader).trim();
      if (!sq.cert && r.cert) sq.cert = String(r.cert).trim();
      if (!sq.proj && r.proj) sq.proj = String(r.proj).trim();
      if (!sq.projId && r.projId) sq.projId = String(r.projId).trim();
      if (isSelf) sq.selfPay = true;
    });

    var list = order.map(function (k) { return map[k]; });
    list.forEach(function (g) {
      g.items = g.order.map(function (p) { return g.squads[p]; });
      g.items.sort(function (a, b) { return b.amount - a.amount; });
    });
    list.sort(function (a, b) { return b.total - a.total; });
    return list;
  }

  function summarySheetName(sheet) {
    var yt = parseYearTerm(sheet);
    return yt.y + yt.tShort + '院系汇总';
  }

  /** 表一：明细（院系分组 + 支队明细，同值列自动合并） */
  function buildDetailGrid(list, sheet) {
    var HEAD = ['打款院系', '打款总金额', '项目负责人', '工作证号', '项目名称', '项目编号', '支队', '对应金额'];
    var cells = [], merges = [], put = function (r, c, v) {
      if (v !== '' && v !== null && v !== undefined) cells.push([r, c, v]);
    };
    put(0, 0, summarySheetName(sheet) + '（院系 × 支队）');
    merges.push([0, 0, 0, HEAD.length - 1]);
    HEAD.forEach(function (h, i) { put(1, i, h); });

    var row = 2, sum = 0;
    list.forEach(function (g) {
      var startRow = row, items = g.items || [];
      var sameOf = function (key) {
        var first = String(items[0][key] || '');
        return items.every(function (it) { return String(it[key] || '') === first; }) ? first : null;
      };
      var sl = sameOf('leader'), sc = sameOf('cert'), sp = sameOf('proj'), si = sameOf('projId');

      items.forEach(function (it, idx) {
        put(row, 0, idx === 0 ? g.dept : '');
        put(row, 1, idx === 0 ? Math.round(g.total * 100) / 100 : '');
        put(row, 2, (sl !== null && idx > 0) ? '' : (it.leader || ''));
        put(row, 3, (sc !== null && idx > 0) ? '' : (it.cert || ''));
        put(row, 4, (sp !== null && idx > 0) ? '' : (it.proj || ''));
        put(row, 5, (si !== null && idx > 0) ? '' : (it.projId || ''));
        put(row, 6, it.place + (it.selfPay ? '（自己报销）' : ''));
        put(row, 7, Math.round(it.amount * 100) / 100);
        row++;
      });
      if (items.length > 1) {
        merges.push([startRow, 0, row - 1, 0]);
        merges.push([startRow, 1, row - 1, 1]);
        if (sl !== null) merges.push([startRow, 2, row - 1, 2]);
        if (sc !== null) merges.push([startRow, 3, row - 1, 3]);
        if (sp !== null) merges.push([startRow, 4, row - 1, 4]);
        if (si !== null) merges.push([startRow, 5, row - 1, 5]);
      }
      sum += g.total;
    });

    var totalR = row;
    put(totalR, 0, '合计');
    put(totalR, 1, Math.round(sum * 100) / 100);
    merges.push([totalR, 2, totalR, 6]);

    return {
      sheetName: summarySheetName(sheet),
      ncols: HEAD.length, cells: cells, merges: merges,
      cols: [22, 14, 14, 15, 32, 14, 34, 14],
      rows: [26, 28],
      styles: [
        { s: { n: '仿宋_GB2312', sz: 12, b: 1, h: 'center', v: 'center' }, addr: 'A1' },
        { s: { n: '仿宋_GB2312', sz: 10, b: 1, h: 'center', v: 'center', w: 1, fg: 'FFD9E1F2' }, addr: 'A2:H2' },
        { s: { n: '仿宋_GB2312', sz: 10, h: 'center', v: 'center', w: 1 }, addr: 'A3:H' + (totalR + 1) },
        { s: { n: '仿宋_GB2312', sz: 10, b: 1, h: 'center', v: 'center' }, addr: 'A' + (totalR + 1) + ':H' + (totalR + 1) }
      ]
    };
  }

  function generateSummary() {
    var sheet = $('note-sheet').value.trim();
    if (!sheet) { alert('请先选择「打款工作表」'); return; }
    var btn = $('note-summary');
    btn.disabled = true;
    setStatus('⏳ 正在读取「' + esc(sheet) + '」并按院系整理…');

    var g;
    callKdocs({ action: 'export', sheet: sheet }).then(function (res) {
      var rows = res.rows || [];
      if (!rows.length) throw new Error('表里读不到数据行');
      g = buildDetailGrid(buildDeptRows(rows), sheet);
      setStatus('⏳ 正在写入「' + esc(g.sheetName) + '」…');
      return callKdocs({
        action: 'syncBudgetGrid', sheetName: g.sheetName, cells: g.cells,
        merges: g.merges, styles: g.styles, cols: g.cols, rows: g.rows, ncols: g.ncols
      });
    }).then(function () {
      setStatus('✅ 已生成「' + g.sheetName + '」：按院系分组，每个支队一行，相同内容自动合并。\n原始打款表未改动。', '#1a7f37');
    }).catch(function (e) {
      setStatus('❌ 汇总失败：' + esc(e.message), '#b3261e');
    }).finally(function () {
      btn.disabled = false;
    });
  }

  /* ================= 主流程 ================= */
  function generate() {
    var sheet = $('note-sheet').value.trim();
    if (!sheet) { // 未选时尝试沿用上一张卡片的打款表名
      var pt = $('payout-title');
      if (pt && pt.value.trim()) sheet = pt.value.trim();
    }
    var topic = $('note-topic').value.trim() || '企业主题实践专项';
    if (!sheet) { alert('请先在「打款工作表」下拉框选择一个工作表（可点刷新列表）'); return; }

    var btn = $('note-gen');
    btn.disabled = true;
    setStatus('⏳ 第 1/3 步：正在从金山读取「' + esc(sheet) + '」…');

    var rowsData, meta, tableRows, totalStr;

    callKdocs({ action: 'export', sheet: sheet }).then(function (res) {
      rowsData = res.rows || [];
      if (!rowsData.length) throw new Error('表里读不到数据行');

      // ① 识别「自己报销」类行（负责人/项目名称里注明院系自行报销的，不进入拨付清单）
      var SELFPAY_RE = /自己报销|自行报销|自付|不拨付|自行处理/;
      var isSelfPay = function (r) {
        return SELFPAY_RE.test(String(r.leader || '') + String(r.proj || ''));
      };
      var sumOf = function (list) {
        return list.reduce(function (t, r) {
          return t + (parseFloat(String(r.amount).replace(/[^\d.\-]/g, '')) || 0);
        }, 0);
      };
      var selfPaid = rowsData.filter(isSelfPay);
      var tableTotalAll = sumOf(rowsData);
      var pool = rowsData;
      if (selfPaid.length) {
        var exAmt = sumOf(selfPaid);
        var exList = selfPaid.map(function (r) {
          return '· ' + (r.place || '') + ' / ' + (r.dept || '') + '：' + fmtAmount(parseFloat(String(r.amount).replace(/[^\d.\-]/g, '')) || 0) + ' 元';
        }).join('\n');
        var keep = window.confirm('检测到 ' + selfPaid.length + ' 行由院系「自己报销」，不应计入拨付清单：\n\n' + exList +
          '\n\n剔除后说明总计 = ' + fmtAmount(tableTotalAll - exAmt) + ' 元' +
          '（打款表总金额 ' + fmtAmount(tableTotalAll) + ' 元，差额 ' + fmtAmount(exAmt) + ' 元即自己报销部分）。' +
          '\n\n点「确定」剔除这些行；点「取消」全部保留（会出现空白行）。');
        if (keep) pool = rowsData.filter(function (r) { return !isSelfPay(r); });
      }
      // 剔除后重算续行标记：同一支队的第一行 cont=false，其余 cont=true
      var lastSquad = null;
      pool.forEach(function (r) {
        r.cont = (r.squad === lastSquad);
        lastSquad = r.squad;
      });
      rowsData = pool;

      // ② 缺项检查（已剔除的行不再提示）
      var miss = [];
      rowsData.forEach(function (r, i) {
        if (!r.dept || !r.leader || !r.cert || !r.projId) miss.push(i + 1);
      });
      if (miss.length) {
        var go = window.confirm('⚠️ 第 ' + miss.join('、') + ' 行的院系/负责人/工作证号/项目编号还没填全，' +
          '生成的说明会有空白。\n\n点「确定」继续生成（之后可重新生成覆盖），点「取消」先去金山补全。');
        if (!go) throw new Error('__cancel__');
      }

      // 拼装文案
      var yt = parseYearTerm(sheet);
      var n = rowsData.filter(function (r) { return !r.cont; }).length;
      meta = {
        title1: '关于' + yt.y + '年' + yt.tFull + '“清华学子地域行”实践',
        title2: topic + '经费拨付事宜的说明',
        body: '根据工作部署，' + yt.y + '年度' + yt.tFull + '实践活动中，' + n +
              '支专项实践支队由相关院系具体承办。依据相关规定，相关经费需统一拨付至承办院系指定对公账户，' +
              '严格执行财务报销制度。（具体拨付明细详见附件）',
        attach: '附件：' + yt.y + '年' + yt.tFull + '社会实践专项经费拨付明细表（含项目名称、承办院系、金额及账户信息）。'
      };

      // 明细行：编号/地点/院系/金额/负责人/证号/项目名/项目编号 + 合并标记
      // 同一支队拆成多个院系行时，cont=true 的行不重复编号，Word 里合并编号/地点单元格
      var total = sumOf(rowsData), squadNo = 0, squads = 0;
      tableRows = rowsData.map(function (r) {
        var amt = parseFloat(String(r.amount).replace(/[^\d.\-]/g, '')) || 0;
        if (!r.cont) { squadNo += 1; squads += 1; }
        return [squadNo, r.place || '', r.dept, fmtAmount(amt), r.leader, r.cert, r.proj, r.projId, !!r.cont];
      });
      totalStr = fmtAmount(total);
      meta.body = meta.body.replace(/\d+\s*支专项实践支队/, squads + '支专项实践支队');

      setStatus('⏳ 第 2/3 步：正在生成 Word（共 ' + n + ' 行明细，总计 ' + totalStr + ' 元）…');
      var doc = buildDocx(meta, tableRows, totalStr);
      var fname = meta.title1 + '-' + meta.title2 + '.docx';
      return downloadDocx(doc, fname);
    }).then(function () {
      setStatus('⏳ 第 3/3 步：Word 已下载，正在同步到金山工作簿…');
      var yt = parseYearTerm(sheet);
      // ① 写进「拨付说明」文件里它自己的脚本（那个脚本运行在自己的文档中，无需跨文件权限）
      return callKdocs({
        action: 'syncNote',
        name: yt.y + yt.tShort + '拨付说明',
        sheetName: yt.y + yt.tShort + '拨付说明',
        separate: false,
        note: {
          title1: meta.title1, title2: meta.title2, body: meta.body, attach: meta.attach,
          rows: tableRows, total: totalStr
        }
      }, { note: true });
    }).then(function (res) {
      // ② 顺手清掉打款表工作簿里历史遗留的说明表（这是另一个文档，走主脚本）
      return callKdocs({ action: 'cleanupNotes' }).then(function (removed) {
        return { note: res, removed: removed || [] };
      }, function () { return { note: res, removed: [] }; });
    }).then(function (pack) {
      var res = pack.note, removed = pack.removed;
      var where;
      if (res && res.mode === 'separate') {
        where = '② 金山云文档「' + esc(res.fileName) + '」里已写入工作表「' + esc(res.sheet) + '」' +
          (res.removed && res.removed.length ? '\n   已清理打款表工作簿里的旧说明表：' + esc(res.removed.join('、')) : '');
      } else {
        where = '② 已在「拨付说明」文件里写入工作表「' + esc(String((res && res.sheet) || res)) + '」';
      }
      if (removed && removed.length) where += '\n   已清理打款表里的旧说明表：' + esc(removed.join('、'));
      setStatus('✅ 全部完成！\n① Word 说明已下载到本地（报送财务用）\n' + where + '\n' +
        '共 ' + tableRows.length + ' 行明细，总计 ' + totalStr + ' 元。', '#1a7f37');
    }).catch(function (e) {
      if (e.message === '__cancel__') { setStatus('已取消。请先在金山表格补全信息，再重新生成。'); return; }
      setStatus('❌ ' + esc(e.message), '#b3261e');
    }).finally(function () {
      btn.disabled = false;
    });
  }
  $('note-gen').onclick = generate;
  $('note-summary').onclick = generateSummary;
})();
