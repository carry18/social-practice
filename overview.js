/**
 * 总览监控：从金山「决算」类工作表读取预算与决算数据，做对比统计
 * 依赖 AirScript action = exportBudget（打款汇总表里的脚本）
 */
(function () {
  var DEFAULT_HOOK = 'https://1386267309-3bh6z757na.ap-shanghai.tencentscf.com';

  function $(id) { return document.getElementById(id); }

  function callKdocs(argv) {
    return fetch(DEFAULT_HOOK, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ Context: { argv: argv } })
    }).then(function (r) { return r.json(); }).then(function (d) {
      if (d.error) throw new Error(d.error);
      return d.data && d.data.result;
    });
  }

  function num(v) {
    var n = parseFloat(String(v == null ? '' : v).replace(/[^\d.\-]/g, ''));
    return isFinite(n) ? n : 0;
  }

  function fmt(n) {
    if (!isFinite(n)) return '—';
    var s = Math.round(n * 100) / 100;
    return s.toLocaleString('zh-CN', { maximumFractionDigits: 2 });
  }

  function pct(actual, budget) {
    if (!budget) return '—';
    var p = (actual - budget) / budget * 100;
    return (p > 0 ? '+' : '') + p.toFixed(1) + '%';
  }

  function esc(s) {
    return String(s == null ? '' : s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }

  function sum(rows, key) {
    return rows.reduce(function (t, r) { return t + num(r[key]); }, 0);
  }

  function render(data) {
    var srcNames = data.map(function (x) { return x.sheet; });
    $('ov-src').textContent = srcNames.length ? srcNames.join('、') : '暂无';

    var all = [];
    data.forEach(function (x) { all = all.concat(x.rows || []); });

    var budget = sum(all, 'budget');
    var actual = sum(all, 'actual');
    var diff = actual - budget;

    $('ov-stats').innerHTML =
      '<div class="stat"><div class="k">期次</div><div class="v">' + data.length + '</div></div>' +
      '<div class="stat"><div class="k">支队数</div><div class="v">' + all.length + '</div></div>' +
      '<div class="stat"><div class="k">预算合计</div><div class="v">' + fmt(budget) + '</div></div>' +
      '<div class="stat"><div class="k">决算合计</div><div class="v">' + fmt(actual) + '</div></div>' +
      '<div class="stat"><div class="k">差额（决算−预算）</div><div class="v ' +
        (diff > 0 ? 'neg' : 'pos') + '">' + (diff > 0 ? '+' : '') + fmt(diff) + '　' + pct(actual, budget) + '</div></div>';

    if (!data.length) {
      $('ov-body').innerHTML = '<p class="hint">金山里还没有名字含「决算」的工作表。</p>';
      return;
    }

    var html = '';
    data.forEach(function (group, gi) {
      var rows = (group.rows || []).filter(function (r) { return num(r.actual) || num(r.budget); });
      var b = sum(rows, 'budget'), a = sum(rows, 'actual');
      html += '<div style="margin-top:18px">' +
        '<h3 style="font-size:14px;margin:0 0 6px;color:var(--purple)">' + esc(group.sheet) +
        '　<span class="hint" style="font-weight:normal">' + rows.length + ' 支队 · 预算 ' + fmt(b) +
        ' · 决算 ' + fmt(a) + ' · 差额 ' + (a - b > 0 ? '+' : '') + fmt(a - b) + '</span></h3>' +
        '<table class="ov"><thead><tr>' +
        '<th>支队</th><th>人数</th><th>预算额</th><th>决算金额</th><th>差额</th><th>差额率</th><th>备注</th>' +
        '</tr></thead><tbody>';
      rows.filter(function (r) { return num(r.actual) || num(r.budget); }).forEach(function (r) {
        var d = num(r.actual) - num(r.budget);
        html += '<tr>' +
          '<td>' + esc(r.place) + (r.squad ? '　<span class="hint">' + esc(r.squad) + (r.date ? ' ' + esc(r.date) : '') + '</span>' : '') + '</td>' +
          '<td>' + esc(r.people) + '</td>' +
          '<td>' + fmt(num(r.budget)) + '</td>' +
          '<td>' + fmt(num(r.actual)) + '</td>' +
          '<td class="' + (d > 0 ? 'neg' : 'pos') + '">' + (d > 0 ? '+' : '') + fmt(d) + '</td>' +
          '<td>' + pct(num(r.actual), num(r.budget)) + '</td>' +
          '<td style="text-align:left">' + esc(r.note) + '</td>' +
          '</tr>';
      });
      html += '<tr><td><b>合计</b></td><td></td><td><b>' + fmt(b) + '</b></td><td><b>' + fmt(a) +
        '</b></td><td><b>' + (a - b > 0 ? '+' : '') + fmt(a - b) + '</b></td><td><b>' + pct(a, b) +
        '</b></td><td></td></tr>';
      html += '</tbody></table></div>';
    });
    $('ov-body').innerHTML = html;
  }

  var loading = false;
  function load() {
    if (loading) return;
    loading = true;
    var msg = $('ov-msg');
    msg.textContent = '读取中…';
    callKdocs({ action: 'exportBudget' }).then(function (data) {
      if (!Array.isArray(data)) throw new Error('返回格式异常，请确认金山脚本已更新到最新版');
      render(data);
      msg.textContent = '更新于 ' + new Date().toLocaleTimeString('zh-CN');
    }).catch(function (e) {
      msg.textContent = '读取失败：' + e.message +
        '（若提示未更新，请把最新 create_payout_sheet.js 贴到打款表脚本里）';
      $('ov-body').innerHTML = '';
      $('ov-stats').innerHTML = '';
    }).finally(function () { loading = false; });
  }

  window.SJ_OVERVIEW = { load: load };

  var btn = $('ov-refresh');
  if (btn) btn.onclick = load;

  // 首屏若停在总览，自动加载一次
  var view = $('view-overview');
  if (view && !view.hidden) load();
})();
