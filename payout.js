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

  /* ================= 接口自检 ================= */
  /**
   * 静默检测全链路：网页 → 腾讯云函数 → 令牌 → 金山脚本。
   * 不读取、不显示任何令牌；仅在失败时指出故障环节和修复方向。
   */
  function runCheck(manual) {
    var el = $('payout-health');
    if (!el) return;
    el.textContent = '🔍 正在检测接口…';
    el.style.color = '';
    fetch(DEFAULT_HOOK, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ Context: { argv: { ping: true } } })
    }).then(function (resp) {
      return resp.text().then(function (t) { return { status: resp.status, text: t }; });
    }).then(function (r) {
      var data = null;
      try { data = JSON.parse(r.text); } catch (e) {}
      var errText = (data && data.error ? String(data.error) : '') + r.text;
      var fail = function (msg) { el.textContent = '❌ ' + msg; el.style.color = '#b3261e'; };

      if (r.status === 401 || r.status === 403 || /token|令牌|unauthor/i.test(errText)) {
        return fail('金山脚本令牌无效或已过期（默认 180 天）→ 管理员：金山脚本编辑器重新生成令牌，更新腾讯云函数环境变量 AIRSCRIPT_TOKEN');
      }
      if (/server misconfigured/.test(errText)) {
        return fail('腾讯云函数缺少环境变量 → 管理员：函数配置中补齐 KDOCS_WEBHOOK / AIRSCRIPT_TOKEN');
      }
      if (/rows 为空/.test(errText)) {
        // 老版本脚本不支持 ping，但报错本身证明链路是通的
        el.textContent = '✅ 接口正常（代理 → 令牌 → 金山脚本全链路连通）';
        el.style.color = '#1a7f37'; return;
      }
      if (data && data.data && data.data.result === 'pong') {
        el.textContent = '✅ 接口正常（代理 → 令牌 → 金山脚本全链路连通）';
        el.style.color = '#1a7f37'; return;
      }
      if (data && data.error) {
        return fail('金山脚本执行报错：' + data.error + ' → 管理员：检查脚本编辑器里的代码是否为最新版本');
      }
      fail('返回内容无法识别：' + r.text.slice(0, 120));
    }).catch(function () {
      el.textContent = '❌ 无法连接腾讯云函数 → 检查函数 URL 是否失效，或函数环境变量 ALLOWED_ORIGIN 是否包含本站域名';
      el.style.color = '#b3261e';
    });
    if (manual) setStatus(''); // 手动点击时清掉旧状态
  }

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

  /* ================= 自动命名 ================= */
  /**
   * 根据文件名 / 文档首段自动识别工作表名：
   *   「2026年暑假“清华学子地域行”实践决算 0918 v3.docx」→「2026暑期院系打款表」
   * 识别不到年份时退化为「<文件名> 打款院系」。
   */
  function autoTitle(fileName, html) {
    var sources = [fileName.replace(/\.docx$/i, '')];
    var pm = /<p[^>]*>([\s\S]*?)<\/p>/.exec(html); // 文档第一个段落（通常是标题）
    if (pm) sources.push(pm[1].replace(/<[^>]+>/g, ''));
    var text = sources.join(' ');

    var yearM = text.match(/(20\d{2})\s*年?/);
    if (yearM) {
      var term = '';
      if (/暑假|暑期/.test(text)) term = '暑期';
      else if (/寒假|冬季/.test(text)) term = '寒假';
      else if (/春季|春学期/.test(text)) term = '春季';
      else if (/秋季|秋学期/.test(text)) term = '秋季';
      else if (/学期中/.test(text)) term = '学期中';
      return yearM[1] + term + '院系打款表';
    }
    return sources[0] + ' 打款院系';
  }

  /* ================= 交互 ================= */
  var drop = $('payout-drop'), fileInput = $('payout-file');
  if (!drop) return; // 页面没有该板块时静默退出

  // 页面加载即静默自检一次；也可点「接口自检」按钮手动复测
  var checkBtn = $('payout-check');
  if (checkBtn) checkBtn.onclick = function () { runCheck(true); };
  runCheck(false);

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
      $('payout-title').value = autoTitle(file.name, result.value); // 自动命名，可手动改
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

  /* ================= 本地 xlsx 兜底（复用页面已加载的 XLSX，含样式） ================= */
  $('payout-xlsx').onclick = function () {
    var headers = ['支队', '打款总金额', '打款院系', '打款金额', '项目负责人', '工作证号', '项目名称', '项目编号（需要是“100002/行政类”）'];
    // 结构与金山在线版一致：第 1 行表头；第 2-3 行红色斜体示例；第 4 行起正式数据；末行合计
    var demo = [
      ['例子支队', 7300, 'A学院', 4000, 'xx', 'xxxxxx', '研究生暑期社会实践', 100002029],
      ['', '', 'B学院', 3300, 'xx', 'xxxxxx', '研究生社会实践专项经费', 100002089]
    ];
    var aoa = [headers].concat(demo);
    for (var i = 0; i < parsedRows.length; i++) {
      aoa.push([parsedRows[i].squad, parsedRows[i].amount, '', '', '', '', '', '']);
    }
    aoa.push(['', '', '', '', '', '', '', '']); // 合计行
    var ws = XLSX.utils.aoa_to_sheet(aoa);

    var n = parsedRows.length, last = 3 + n, totalR = last + 1; // 1-based 行号
    ws['B' + totalR] = { f: 'SUM(B4:B' + last + ')', t: 'n' };
    // 样式 1:1 复刻《2026年春季学期暑期实践打款院系.xlsx》实测格式
    ws['!cols'] = [{ wch: 14.3 }, { wch: 15.2 }, { wch: 17.5 }, { wch: 9.7 },
                   { wch: 10.7 }, { wch: 11.8 }, { wch: 37 }, { wch: 36.5 }];
    ws['!merges'] = [ // 示例行 A/B 列纵向合并（0-based）
      { s: { r: 1, c: 0 }, e: { r: 2, c: 0 } },
      { s: { r: 1, c: 1 }, e: { r: 2, c: 1 } }
    ];
    ws['!rows'] = [];
    for (var rr = 3; rr < last; rr++) ws['!rows'][rr] = { hpt: 39 }; // 数据行高 39（0-based r=3 即第 4 行）

    var CENTER = { horizontal: 'center', vertical: 'center' };
    var mkStyle = function (name, sz, extra) {
      var o = { font: { name: name, sz: sz }, alignment: CENTER };
      if (extra) for (var k in extra) o.font[k] = extra[k];
      return o;
    };
    var ref = XLSX.utils.decode_range(ws['!ref']);
    for (var R = ref.s.r; R <= ref.e.r; R++) {
      for (var C = ref.s.c; C <= ref.e.c; C++) {
        var addr = XLSX.utils.encode_cell({ r: R, c: C });
        if (!ws[addr]) ws[addr] = { t: 's', v: '' };
        var st;
        if (R === 0) { // 表头：A1 宋体8 / B1 宋体11 / C:G 仿宋9加粗蓝底 / H 黄底红字
          if (C === 0) st = mkStyle('宋体', 8);
          else if (C === 1) st = mkStyle('宋体', 11);
          else {
            st = mkStyle('仿宋_GB2312', 9, { bold: true });
            st.fill = { patternType: 'solid', fgColor: { rgb: 'FF00B0F0' } };
            if (C === 7) { st.font.color = { rgb: 'FFFF0000' }; st.fill.fgColor.rgb = 'FFFFFF00'; }
          }
        } else if (R <= 2) { // 示例行：红色斜体；A列8号、B列11号、其余10号
          st = mkStyle('宋体', C === 0 ? 8 : (C === 1 ? 11 : 10), { italic: true, color: { rgb: 'FFFF0000' } });
        } else if (R === totalR - 1) { // 合计行
          st = mkStyle('宋体', 11);
        } else { // 正式数据：支队列宋体8自动换行，B列10号，其余11号
          st = mkStyle('宋体', C === 0 ? 8 : (C === 1 ? 10 : 11));
          if (C === 0) st.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
        }
        ws[addr].s = st;
      }
    }

    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '打款院系');
    XLSX.writeFile(wb, ($('payout-title').value.trim() || '实践打款院系') + '.xlsx');
  };
})();
