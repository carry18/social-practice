/* 报名推送：期次 → 创建/复用报名表 → 二维码 → 推送 Word
 * ------------------------------------------------------------------
 * 链接只有两个来源：云函数建表结果 / 高级设置里手动粘贴；按期次分别记忆
 * （localStorage「sj_signup_links」，key 形如「2026年秋季学期中」）。
 * 二维码参数与 自动化/web/tools/qr_into_docx.py 对齐：纠错 M、白边 3 模块、900px、放大 1.05。
 * ------------------------------------------------------------------ */
(function () {
  'use strict';

  var FN_URL = 'https://1386267309-knom92xprw.ap-shanghai.tencentscf.com';  // 建表云函数（地址公开，无口令）
  var QR_BORDER = 3;      // 白边模块数（标准是 4，3 是为了把码印得大一点）
  var QR_PX = 900;
  var SCALE = 1.05;       // 图片显示尺寸倍数（与 Python 工具一致，码面约 4.0cm）

  var $ = function (id) { return document.getElementById(id); };
  var current = null;     // { dataUrl, base64 }
  var currentUrl = '';    // 当前期次生效的报名表链接
  var lastKey = null;     // 上一次生效的期次 key，避免 change 事件循环

  /* ================= 支队数据 ================= */
  function teamsFromBudget() {
    if (window.BudgetBridge && window.BudgetBridge.getTeams) {
      return (window.BudgetBridge.getTeams() || []).filter(function (t) {
        return String(t.B || '').trim();
      });
    }
    return [];
  }

  function renderTeams() {
    var tb = $('signup-teams-body');
    if (!tb) return;
    var rows = teamsFromBudget();
    tb.innerHTML = '';
    if (!rows.length) {
      tb.innerHTML = '<tr><td colspan="5" class="hint">还没有支队：先到「预算编制」录入支队信息</td></tr>';
      refreshButtons();
      return;
    }
    rows.forEach(function (t, i) {
      var tr = document.createElement('tr');
      tr.innerHTML = '<td>' + (i + 1) + '</td>' +
        '<td>' + String(t.B || '').replace(/</g, '&lt;') + '</td>' +
        '<td>' + String(t.E || '') + '</td>' +
        '<td>' + String(t.C || '').replace(/</g, '&lt;') + '</td>' +
        '<td>' + String(t.D || '') + '</td>';
      tb.appendChild(tr);
    });
    refreshButtons();
  }

  /* ================= 二维码 ================= */
  function makeQr(text) {
    if (!window.qrcode) throw new Error('二维码库未加载（qrcode-generator 的 CDN 可能被拦了）');
    var qr = window.qrcode(0, 'M');
    qr.addData(text);
    qr.make();

    var n = qr.getModuleCount();
    var total = n + QR_BORDER * 2;
    var scale = Math.max(1, Math.round(QR_PX / total));
    var size = total * scale;

    var canvas = document.createElement('canvas');
    canvas.width = canvas.height = size;
    var ctx = canvas.getContext('2d');
    ctx.fillStyle = '#fff';
    ctx.fillRect(0, 0, size, size);
    ctx.fillStyle = '#000';
    for (var r = 0; r < n; r++) {
      for (var c = 0; c < n; c++) {
        if (qr.isDark(r, c)) {
          ctx.fillRect((c + QR_BORDER) * scale, (r + QR_BORDER) * scale, scale, scale);
        }
      }
    }
    var dataUrl = canvas.toDataURL('image/jpeg', 0.95);
    return { dataUrl: dataUrl, base64: dataUrl.split(',')[1], modules: n, px: size };
  }

  /* ================= 期次与链接记忆 ================= */
  function termVal() {
    return { year: ($('signup-year') || {}).value || '', term: ($('signup-term') || {}).value || '' };
  }
  function termKey(y, t) { return (y && t) ? y + '年' + t : ''; }
  function formName(y, t) { return y + '年' + t + '社会实践报名'; }

  function loadLinks() {
    try { return JSON.parse(localStorage.getItem('sj_signup_links') || '{}'); } catch (e) { return {}; }
  }
  function saveLink(key, url, src) {
    if (!key) return;
    var m = loadLinks();
    m[key] = { url: url, src: src };
    try { localStorage.setItem('sj_signup_links', JSON.stringify(m)); } catch (e) {}
  }
  function getLink(key) {
    return key ? (loadLinks()[key] || null) : null;
  }

  /* ================= 展示 ================= */
  function setStatus(id, msg, bad) {
    var el = $(id);
    if (!el) return;
    el.textContent = msg || '';
    el.style.color = bad ? '#b3261e' : '#1a7f37';
  }

  function updateNamePreview() {
    var v = termVal();
    var el = $('signup-name-preview');
    if (el) el.textContent = termKey(v.year, v.term) ? '将生成：' + formName(v.year, v.term) : '先选期次';
  }

  function renderLink() {
    var textEl = $('signup-url-text');
    if (currentUrl) {
      textEl.textContent = currentUrl;
      textEl.style.color = '';
      $('signup-copy-url').hidden = false;
    } else {
      textEl.textContent = '（本期还没有报名表，点上方「创建/复用报名表」）';
      textEl.style.color = '#999';
      $('signup-copy-url').hidden = true;
    }
    var src = '';
    var v = termVal();
    var rec = getLink(termKey(v.year, v.term));
    if (currentUrl && rec && rec.url === currentUrl && rec.src) src = '（' + rec.src + '）';
    $('signup-url-src').textContent = src;
  }

  function showQr(url) {
    current = makeQr(url);
    $('signup-qr-img').src = current.dataUrl;
    $('signup-qr-dl').href = current.dataUrl;
    $('signup-qr-box').style.display = 'block';
  }

  function hideQr() {
    current = null;
    $('signup-qr-box').style.display = 'none';
  }

  function refreshButtons() {
    var v = termVal();
    var hasTerm = !!termKey(v.year, v.term);
    if ($('signup-create')) $('signup-create').disabled = !hasTerm;
    if ($('signup-push')) $('signup-push').disabled = !(hasTerm && current && teamsFromBudget().length);
  }

  /* 期次切换：保存旧期次链接 → 恢复新期次链接（或清空） */
  function onTermChange() {
    var v = termVal();
    var key = termKey(v.year, v.term);
    if (key === lastKey) return;         // 同步事件来回触发时防抖
    lastKey = key;

    // 写回「预算编制」的下拉（推送文件名、公文标题都用它们）
    if ($('doc-year') && v.year) $('doc-year').value = v.year;
    if ($('doc-term')) $('doc-term').value = v.term;
    updateNamePreview();

    if (!key) { currentUrl = ''; hideQr(); renderLink(); refreshButtons(); return; }
    var rec = getLink(key);
    if (rec && rec.url) {
      currentUrl = rec.url;
      showQr(rec.url);
      setStatus('signup-create-status', '', false);
    } else {
      currentUrl = '';
      hideQr();
      setStatus('signup-create-status', '本期还没有报名表，点左侧按钮创建', true);
    }
    renderLink();
    refreshButtons();
  }

  function applyLink(url, src) {
    var v = termVal();
    currentUrl = url;
    saveLink(termKey(v.year, v.term), url, src);
    showQr(url);
    renderLink();
    refreshButtons();
  }

  /* ================= 年份下拉 ================= */
  function fillSignupYears() {
    var y = $('signup-year'), src = $('doc-year');
    if (!y || y.options.length) return;
    var opts = [];
    if (src && src.options.length) {
      Array.prototype.forEach.call(src.options, function (o) { opts.push([o.value, o.textContent]); });
    } else {
      var now = new Date().getFullYear();
      for (var i = now - 2; i <= now + 4; i++) opts.push([String(i), i + '年']);
    }
    opts.forEach(function (pair) {
      var op = document.createElement('option');
      op.value = pair[0]; op.textContent = pair[1];
      y.appendChild(op);
    });
    y.value = (src && src.value) || String(new Date().getFullYear());
  }

  /* 「预算编制」里改期次 → 报名面板跟随 */
  function syncSignupFromDoc() {
    if (!$('signup-term')) return;
    var changed = false;
    if ($('doc-term') && $('signup-term').value !== $('doc-term').value) {
      $('signup-term').value = $('doc-term').value;
      changed = true;
    }
    if ($('doc-year') && $('doc-year').value && $('signup-year').value !== $('doc-year').value) {
      $('signup-year').value = $('doc-year').value;
      changed = true;
    }
    if (changed) onTermChange();
  }

  /* ================= 建表 ================= */
  function createForm() {
    var v = termVal();
    if (!termKey(v.year, v.term)) return;   // 按钮已置灰，双保险
    var rows = teamsFromBudget();
    if (!rows.length) { setStatus('signup-create-status', '「预算编制」里还没有支队', true); return; }

    var payload = {
      action: 'createForm',
      name: formName(v.year, v.term),
      teams: rows.map(function (r, i) {
        return {
          label: (i + 1) + '. ' + String(r.B || '') + (r.E ? ' ' + r.E : '') + (r.D ? ' ' + r.D : ''),
          city_full: String(r.B || ''),
          date: String(r.E || ''),
          dept: String(r.D || '')
        };
      })
    };
    var key = ($('signup-fn-key').value || '').trim();
    try { localStorage.setItem('sj_signup_key', key); } catch (e) {}
    var headers = { 'Content-Type': 'application/json' };
    if (key) headers['x-signup-key'] = key;

    setStatus('signup-create-status', '正在建表…');
    fetch(FN_URL, { method: 'POST', headers: headers, body: JSON.stringify(payload) })
      .then(function (r) { return r.json().then(function (d) { return { status: r.status, data: d }; }); })
      .then(function (r) {
        var d = r.data || {};
        if (!d.url) throw new Error(d.error || ('HTTP ' + r.status));
        applyLink(d.url, '来自建表结果');
        setStatus('signup-create-status',
          (r.status === 409 ? '✅ 本期报名表已存在，直接使用' : '✅ 建表成功') +
          (d.fields ? '（' + d.fields + ' 个字段）' : ''));
      })
      .catch(function (e) { setStatus('signup-create-status', '❌ ' + (e.message || e), true); });
  }

  /* ================= 下载推送 ================= */
  function downloadPush() {
    var rows = teamsFromBudget();
    if (!rows.length || !current) return;   // 按钮已置灰，双保险
    if (!window.SJ_PUSH) { setStatus('signup-status', '❌ app.js 未暴露 SJ_PUSH，请刷新页面', true); return; }

    setStatus('signup-status', '正在生成推送…');
    var v = termVal();
    window.SJ_PUSH.generate({
      list: rows,
      qrJpeg: current.base64,
      scale: SCALE,
      filename: '【就协】' + formName(v.year, v.term).replace('报名', '招募推送') + rows.length + '支.docx'
    }).then(function () {
      setStatus('signup-status', '✅ 推送已下载（二维码已嵌入）');
    }).catch(function (e) {
      setStatus('signup-status', '❌ ' + (e && e.message ? e.message : e), true);
    });
  }

  /* ================= 初始化 ================= */
  function init() {
    if (!$('view-signup')) return;

    fillSignupYears();
    renderTeams();

    // 口令记忆
    try { $('signup-fn-key').value = localStorage.getItem('sj_signup_key') || ''; } catch (e) {}

    // 旧版单一链接（sj_signup_url）迁移：当前期次没有记忆时用它顶上
    try {
      var legacy = localStorage.getItem('sj_signup_url');
      if (legacy) {
        var v = termVal();
        var key = termKey(v.year, v.term);
        if (key && !getLink(key)) saveLink(key, legacy, '来自建表结果');
        localStorage.removeItem('sj_signup_url');
        localStorage.removeItem('sj_signup_url_src');
      }
    } catch (e) {}

    syncSignupFromDoc();   // 预算编制里已选过期次则跟随
    updateNamePreview();
    onTermChange();        // 恢复本期链接（若有）

    $('signup-year').addEventListener('change', onTermChange);
    $('signup-term').addEventListener('change', onTermChange);
    ['doc-year', 'doc-term'].forEach(function (id) {
      if ($(id)) $(id).addEventListener('change', syncSignupFromDoc);
    });

    $('signup-create').addEventListener('click', createForm);
    $('signup-push').addEventListener('click', downloadPush);

    $('signup-copy-url').addEventListener('click', function () {
      if (!currentUrl || !navigator.clipboard) return;
      navigator.clipboard.writeText(currentUrl).then(function () {
        $('signup-copy-url').textContent = '已复制';
        setTimeout(function () { $('signup-copy-url').textContent = '复制'; }, 1500);
      });
    });

    $('signup-use-manual').addEventListener('click', function () {
      var url = ($('signup-url-manual').value || '').trim();
      var hint = $('signup-manual-hint');
      if (!/^https?:\/\//.test(url)) {
        hint.textContent = '链接要以 http(s):// 开头';
        hint.style.color = '#b3261e';
        return;
      }
      var v = termVal();
      if (!termKey(v.year, v.term)) {
        hint.textContent = '先在上方选期次';
        hint.style.color = '#b3261e';
        return;
      }
      hint.textContent = '';
      $('signup-url-manual').value = '';
      applyLink(url, '手动粘贴');
      setStatus('signup-create-status', '✅ 已使用粘贴的链接');
    });

    // 预算编制页③的「前往报名推送」跳转按钮
    var go = $('goto-signup');
    if (go) go.addEventListener('click', function () {
      var nav = document.querySelector('aside.side button[data-view="signup"]');
      if (nav) nav.click();
    });

    window.addEventListener('budget:render', renderTeams);
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
  else init();
})();
