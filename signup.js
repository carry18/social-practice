/* 报名推送：报名表链接 → 二维码 → 含二维码的招募推送 Word
 * ------------------------------------------------------------------
 * 依赖：qrcode-generator（CDN）、PizZip、app.js 暴露的 window.SJ_PUSH。
 * 二维码参数与 自动化/web/tools/qr_into_docx.py 对齐：纠错 M、白边 3 模块、900px、图片放大 1.15。
 * ------------------------------------------------------------------ */
(function () {
  'use strict';

  var DEFAULT_FN = 'https://1386267309-knom92xprw.ap-shanghai.tencentscf.com';  // 建表云函数（地址公开，无口令）
  var QR_BORDER = 3;      // 白边模块数（标准是 4，3 是为了把码印得大一点）
  var QR_PX = 900;
  var SCALE = 1.05;       // 图片显示尺寸倍数（与 Python 工具一致，码面约 4.0cm）

  var $ = function (id) { return document.getElementById(id); };
  var current = null;     // { dataUrl, base64, modules }

  function teamsFromBudget() {
    if (window.BudgetBridge && window.BudgetBridge.getTeams) {
      return (window.BudgetBridge.getTeams() || []).filter(function (t) {
        return String(t.B || '').trim();
      });
    }
    return [];
  }

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

  function showQr() {
    var url = ($('signup-url').value || '').trim();
    var status = $('signup-status');
    if (!/^https?:\/\//.test(url)) {
      alert('报名表链接要以 http(s):// 开头');
      return;
    }
    try {
      current = makeQr(url);
    } catch (e) {
      alert(e.message);
      return;
    }
    try { localStorage.setItem('sj_signup_url', url); } catch (e) {}
    $('signup-qr-img').src = current.dataUrl;
    $('signup-qr-dl').href = current.dataUrl;
    $('signup-qr-box').style.display = 'block';
    $('signup-qr-note').textContent = current.modules + '×' + current.modules +
      ' 模块，' + current.px + 'px，白边 ' + QR_BORDER + ' 模块';
    status.textContent = '✅ 二维码已生成，可下载推送';
    status.style.color = '#1a7f37';
  }

  function renderTeams() {
    var tb = $('signup-teams-body');
    if (!tb) return;
    var rows = teamsFromBudget();
    tb.innerHTML = '';
    if (!rows.length) {
      tb.innerHTML = '<tr><td colspan="5" class="hint">还没有支队：先到「预算编制」录入支队信息</td></tr>';
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
  }

  function downloadPush() {
    var rows = teamsFromBudget();
    var status = $('signup-status');
    if (!rows.length) { alert('「预算编制」里还没有支队信息，推送表格会是空的'); return; }
    if (!current) { alert('先生成二维码'); return; }
    if (!window.SJ_PUSH) { alert('app.js 未暴露 SJ_PUSH，请刷新页面'); return; }

    status.textContent = '正在生成推送…';
    status.style.color = '';
    var term = (($('doc-term') || {}).value || '').trim();
    var year = (($('doc-year') || {}).value || String(new Date().getFullYear()));

    window.SJ_PUSH.generate({
      list: rows,
      qrJpeg: current.base64,
      scale: SCALE,
      filename: '【就协】' + year + (term || '') + '社会实践招募推送' + rows.length + '支.docx'
    }).then(function () {
      status.textContent = '✅ 推送已下载（二维码已嵌入）';
      status.style.color = '#1a7f37';
    }).catch(function (e) {
      status.textContent = '❌ ' + (e && e.message ? e.message : e);
      status.style.color = '#b3261e';
    });
  }

  function savedConfig() {
    var url = '', key = '';
    try {
      url = localStorage.getItem('sj_signup_fn') || '';
      key = localStorage.getItem('sj_signup_key') || '';
    } catch (e) {}
    return { url: url, key: key };
  }

  function rememberConfig() {
    try {
      localStorage.setItem('sj_signup_fn', ($('signup-fn-url').value || '').trim());
      localStorage.setItem('sj_signup_key', ($('signup-fn-key').value || '').trim());
    } catch (e) {}
  }

  function needTerm() {
    var term = (($('doc-term') || {}).value || '').trim();
    if (!term) throw new Error('请先在「预算编制」里选择期次');
    var year = (($('doc-year') || {}).value || String(new Date().getFullYear()));
    return { year: year, term: term };
  }

  function setCreateStatus(msg, bad) {
    var el = $('signup-create-status');
    el.textContent = msg;
    el.style.color = bad ? '#b3261e' : '#1a7f37';
  }

  function createForm() {
    var url = ($('signup-fn-url').value || '').trim();
    if (!url) { setCreateStatus('先填云函数地址（见 自动化/报名/cloudfn/README.md）', true); return; }
    var rows = teamsFromBudget();
    if (!rows.length) { setCreateStatus('「预算编制」里还没有支队', true); return; }
    rememberConfig();

    var t;
    try { t = needTerm(); } catch (e) { setCreateStatus(e.message, true); return; }

    var payload = {
      action: 'createForm',
      name: t.year + '年' + t.term + '社会实践报名',
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
    var headers = { 'Content-Type': 'application/json' };
    if (key) headers['x-signup-key'] = key;

    setCreateStatus('正在建表…');
    fetch(url, { method: 'POST', headers: headers, body: JSON.stringify(payload) })
      .then(function (r) { return r.json().then(function (d) { return { ok: r.ok, status: r.status, data: d }; }); })
      .then(function (r) {
        var d = r.data || {};
        if (!d.url) throw new Error(d.error || ('HTTP ' + r.status));
        $('signup-url').value = d.url;
        showQr();
        setCreateStatus((r.status === 409 ? '⚠️ 同名表已存在，已复用：' : '✅ 建表成功：') +
          d.url + ' ｜ token ' + d.token + (d.fields ? ' ｜ ' + d.fields + ' 个字段' : ''));
      })
      .catch(function (e) { setCreateStatus('❌ ' + (e.message || e), true); });
  }

  function buildCommand() {
    var term = (($('doc-term') || {}).value || '秋季学期中');
    var year = (($('doc-year') || {}).value || new Date().getFullYear());
    var cmd = 'python3 create_form.py --plan "../../学期中实践点位安排整合表.xlsx" \\\n' +
      '  --year ' + year + ' --term ' + term + ' --live';
    $('signup-cmd').textContent = cmd;
  }

  function init() {
    if (!$('view-signup')) return;
    renderTeams();
    buildCommand();

    var cfg = savedConfig();
    $('signup-fn-url').value = cfg.url || DEFAULT_FN;
    if (cfg.key) $('signup-fn-key').value = cfg.key;

    $('signup-create').addEventListener('click', createForm);

    // 链接一改就自动出码（防抖 300ms），不再需要单独点按钮
    var timer = null;
    $('signup-url').addEventListener('input', function () {
      clearTimeout(timer);
      timer = setTimeout(function () {
        var v = ($('signup-url').value || '').trim();
        if (!v) { $('signup-qr-box').style.display = 'none'; return; }
        showQr();
      }, 300);
    });

    // 上次用过的链接：刷新后自动回填并出码
    try {
      var last = localStorage.getItem('sj_signup_url');
      if (last) { $('signup-url').value = last; showQr(); }
    } catch (e) {}
    $('signup-push').addEventListener('click', downloadPush);
    $('signup-copy').addEventListener('click', function () {
      var txt = $('signup-cmd').textContent;
      if (navigator.clipboard) navigator.clipboard.writeText(txt).then(function () {
        $('signup-copy').textContent = '已复制';
        setTimeout(function () { $('signup-copy').textContent = '复制'; }, 1500);
      });
    });
    ['doc-year', 'doc-term'].forEach(function (id) {
      if ($(id)) $(id).addEventListener('change', buildCommand);
    });
    window.addEventListener('budget:render', renderTeams);
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
  else init();
})();
