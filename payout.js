/**
 * 决算表 → 打款院系工作表（嵌入预算表生成器页面）
 * ------------------------------------------------------------------
 * 流程：上传决算 docx → mammoth 浏览器内解析 → 提取「院系联合组织」支队
 *       → POST 腾讯云函数代理（令牌存在函数环境变量中，页面零秘密）
 *       → 金山 AirScript 在打款汇总表格中新建工作表。
 * 兜底：代理不可达时可下载本地 xlsx（复用页面已加载的 xlsx-js-style）。
 * 本文件独立于 app.js，所有 DOM id 带 payout- 前缀，不影响原有功能。
 * ------------------------------------------------------------------
 */
(function () {
  'use strict';

  // 内置默认代理地址（腾讯云函数 URL，非机密）
  var DEFAULT_HOOK = 'https://1386267309-3bh6z757na.ap-shanghai.tencentscf.com';

  var $ = function (id) { return document.getElementById(id); };
  var parsedRows = [];

  /* ================= docx 解析 ================= */
  function htmlTables(html) {
    var tables = [];
    var tableRe = /<table[\s\S]*?<\/table>/gi;
    var m;
    while ((m = tableRe.exec(html))) {
      var rows = [];
      var trRe = /<tr[\s\S]*?<\/tr>/gi;
      var r;
      while ((r = trRe.exec(m[0]))) {
        var cells = [];
        var tdRe = /<t[dh][^>]*>([\s\S]*?)<\/t[dh]>/gi;
        var c;
        while ((c = tdRe.exec(r[0]))) {
          var t = c[1]
            .replace(/<br\s*\/?>/gi, '\n')
            .replace(/<\/p>\s*<p[^>]*>/gi, '\n')
            .replace(/<[^>]+>/g, '')
            .replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&')
            .replace(/&lt;/g, '<').replace(/&gt;/g, '>');
          t = t.split('\n').map(function (s) { return s.trim(); }).filter(Boolean).join('\n');
          cells.push(t);
        }
        if (cells.length) rows.push(cells);
      }
      tables.push(rows);
    }
    return tables;
  }

  function toNum(s) {
    var v = parseFloat(String(s).replace(/[^\d.]/g, ''));
    return isFinite(v) ? v : 0;
  }

  /** 提取需打款支队：企业支队表中备注含「院系」的行；剔除「就协全部」 */
  function extractPayouts(tables) {
    var payouts = [];
    for (var ti = 0; ti < tables.length; ti++) {
      var rows = tables[ti];
      var headerIdx = -1;
      for (var i = 0; i < rows.length; i++) {
        if (rows[i][0] === '序号' && rows[i].some(function (c) { return c.indexOf('目的地') >= 0; })) {
          headerIdx = i; break;
        }
      }
      if (headerIdx < 0) continue;
      for (var j = headerIdx + 1; j < rows.length; j++) {
        var r = rows[j];
        if (r.length < 11) continue;                     // 跳过子表头
        if (r[0] === '合计' || r[1] === '合计' || !r[1]) continue;
        var remark = r[r.length - 1] || '';
        if (remark.indexOf('就协全部') >= 0) continue;   // 仅就协单独承担 → 删除
        if (remark.indexOf('院系') < 0) continue;        // 非院系联合组织 → 不列入
        payouts.push({ squad: r[1], amount: toNum(r[r.length - 2]) });
      }
    }
    return payouts;
  }

  /* ================= 交互 ================= */
  var drop = $('payout-drop'), fileInput = $('payout-file');
  if (!drop) return; // 页面没有该板块时静默退出

  drop.onclick = function () { fileInput.click(); };
  drop.ondragover = function (e) { e.preventDefault(); drop.classList.add('hover'); };
  drop.ondragleave = function () { drop.classList.remove('hover'); };
  drop.ondrop = function (e) {
    e.preventDefault(); drop.classList.remove('hover');
    if (e.dataTransfer.files[0]) handleFile(e.dataTransfer.files[0]);
  };
  fileInput.onchange = function () { if (fileInput.files[0]) handleFile(fileInput.files[0]); };

  function handleFile(file) {
    if (!/\.docx$/i.test(file.name)) { alert('请上传 .docx 文件'); return; }
    file.arrayBuffer().then(function (buf) {
      return mammoth.convertToHtml({ arrayBuffer: buf });
    }).then(function (result) {
      parsedRows = extractPayouts(htmlTables(result.value));
      if (!parsedRows.length) { alert('未解析到需要打款的院系联合组织支队'); return; }
      var base = file.name.replace(/\.docx$/i, '').replace(/\s*v?\d+(\.\d+)*\s*$/i, '');
      if (!$('payout-title').value) $('payout-title').value = base + ' 打款院系';
      renderPreview();
      $('payout-result').style.display = 'block';
      $('payout-result').scrollIntoView({ behavior: 'smooth' });
      generate(); // 上传即触发
    }).catch(function (e) {
      alert('解析失败：' + e.message);
    });
  }

  function renderPreview() {
    var total = 0;
    var h = '<tr><th style="width:45%">支队（地点 / 院系）</th><th>打款总金额（决算总计）</th></tr>';
    for (var i = 0; i < parsedRows.length; i++) {
      total += parsedRows[i].amount;
      h += '<tr><td class="squad">' + esc(parsedRows[i].squad) + '</td><td>' + parsedRows[i].amount + '</td></tr>';
    }
    h += '<tr><th>合计（' + parsedRows.length + ' 支）</th><th>' + (Math.round(total * 100) / 100) + '</th></tr>';
    $('payout-preview').innerHTML = h;
  }

  function esc(s) {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }

  function setStatus(msg, cls) {
    var el = $('payout-status');
    el.className = cls || '';
    el.innerHTML = msg;
  }

  /* ================= 调用代理 → 金山 AirScript ================= */
  function generate() {
    var btn = $('payout-gen');
    btn.disabled = true;
    setStatus('正在调用接口新建工作表…');
    fetch(DEFAULT_HOOK, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        Context: {
          argv: {
            title: $('payout-title').value.trim() || '实践打款院系',
            rows: parsedRows
          }
        }
      })
    }).then(function (resp) {
      return resp.json().then(function (data) { return { ok: resp.ok, data: data }; });
    }).then(function (r) {
      if (!r.ok || r.data.error) throw new Error(r.data.error || ('HTTP ' + (r.ok ? 200 : 'error')));
      var result = r.data.data && r.data.data.result;
      var isLink = /^https?:\/\//.test(String(result));
      setStatus(isLink
        ? '✅ 已创建在线 Excel：<a href="' + esc(result) + '" target="_blank">' + esc(result) + '</a>\n请打开补全：打款院系 / 打款金额 / 负责人 / 工作证号 / 项目名称 / 项目编号。'
        : '✅ ' + esc(String(result)) + '\n请到打款汇总的金山文档表格查看新工作表，并补全：打款院系 / 打款金额 / 负责人 / 工作证号 / 项目名称 / 项目编号。', 'ok');
    }).catch(function (e) {
      setStatus('❌ 调用失败：' + esc(e.message) +
        '\n可先点击「下载本地 xlsx 备份」，再手动上传到金山文档。', 'err');
    }).finally(function () {
      btn.disabled = false;
    });
  }
  $('payout-gen').onclick = generate;

  /* ================= 本地 xlsx 兜底（复用页面已加载的 XLSX） ================= */
  $('payout-xlsx').onclick = function () {
    var headers = ['支队', '打款总金额', '打款院系', '打款金额', '项目负责人', '工作证号', '项目名称', '项目编号（需要是“100002/行政类”）'];
    var aoa = [headers];
    for (var i = 0; i < parsedRows.length; i++) {
      aoa.push([parsedRows[i].squad, parsedRows[i].amount, '', '', '', '', '', '']);
    }
    aoa.push(['', '', '', '', '', '', '', '']); // 合计行
    var ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['B' + (parsedRows.length + 2)] = { f: 'SUM(B2:B' + (parsedRows.length + 1) + ')', t: 'n' };
    ws['!cols'] = [{ wch: 26 }, { wch: 12 }, { wch: 18 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 26 }, { wch: 30 }];
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '打款院系');
    XLSX.writeFile(wb, ($('payout-title').value.trim() || '实践打款院系') + '.xlsx');
  };
})();
