/**
 * 总览监控：预算文档「预算汇总」+ 决算文档「决算汇总」→ 年×期次表 + 趋势图。
 * 不再读取原始 v1 决算明细工作表。
 */
(function () {
  var FINAL_HOOK = 'https://1386267309-3bh6z757na.ap-shanghai.tencentscf.com';
  var BUDGET_HOOK = 'https://1386267309-3bh6z757na.ap-shanghai.tencentscf.com';

  function $(id) { return document.getElementById(id); }

  function callHook(hook, target, argv) {
    var body = { Context: { argv: argv } };
    if (target) body.target = target;
    return fetch(hook, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body)
    }).then(function (resp) {
      return resp.json().then(function (data) { return { ok: resp.ok, data: data }; });
    }).then(function (r) {
      if (!r.ok || r.data.error) throw new Error(r.data.error || ('HTTP ' + (r.ok ? 200 : 'error')));
      return (r.data.data && r.data.data.result) || [];
    });
  }

  function callFinal(argv) { return callHook(FINAL_HOOK, '', argv); }
  function callBudget(argv) { return callHook(BUDGET_HOOK, 'budget', argv); }

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

  function safe(p) {
    return p.then(function (v) { return { v: v }; })
      .catch(function (e) { return { e: e }; });
  }

  function combine(budgetRows, finalRows) {
    var map = {};
    function add(rows, kind) {
      (rows || []).forEach(function (r) {
        var year = String(r.year || '未识别');
        var term = String(r.term || '其他');
        var key = year + '|' + term;
        var c = map[key] || (map[key] = {
          year: year, term: term, b: 0, a: 0, tb: 0, ta: 0, hasB: false, hasA: false
        });
        if (kind === 'b') {
          c.b += num(r.amount); c.tb += num(r.teams); c.hasB = true;
        } else {
          c.a += num(r.amount); c.ta += num(r.teams); c.hasA = true;
        }
      });
    }
    add(budgetRows, 'b');
    add(finalRows, 'a');
    return Object.keys(map).map(function (k) { return map[k]; }).sort(function (x, y) {
      if (x.year === '未识别') return 1;
      if (y.year === '未识别') return -1;
      return y.year.localeCompare(x.year);
    });
  }

  function renderStats(rows) {
    var budget = 0, actual = 0, teams = 0;
    rows.forEach(function (r) {
      budget += r.b;
      actual += r.a;
      teams += Math.max(r.tb, r.ta);
    });
    var diff = actual - budget;
    $('ov-stats').innerHTML =
      '<div class="stat"><div class="k">期次</div><div class="v">' + rows.length + '</div></div>' +
      '<div class="stat"><div class="k">支队数</div><div class="v">' + teams + '</div></div>' +
      '<div class="stat"><div class="k">预算合计</div><div class="v">' + fmt(budget) + '</div></div>' +
      '<div class="stat"><div class="k">决算合计</div><div class="v">' + fmt(actual) + '</div></div>' +
      '<div class="stat"><div class="k">差额（决算−预算）</div><div class="v ' +
        (diff > 0 ? 'neg' : 'pos') + '">' + (diff > 0 ? '+' : '') + fmt(diff) +
        '　' + pct(actual, budget) + '</div></div>';
  }

  var TERMS = ['春季学期中', '秋季学期中', '寒假', '暑假'];

  function renderSummary(rows) {
    var el = $('ov-summary');
    if (!el) return;
    if (!rows.length) {
      el.innerHTML = '<p class="hint">暂无预算/决算汇总数据。请先在预算/决算文档里保存一次，或确认金山脚本已更新。</p>';
      return;
    }
    var hasOther = rows.some(function (r) { return r.term === '其他'; });
    var terms = TERMS.slice();
    if (hasOther) terms.push('其他');
    var byYear = {};
    rows.forEach(function (r) {
      if (!byYear[r.year]) byYear[r.year] = {};
      byYear[r.year][r.term] = r;
    });
    var years = Object.keys(byYear).sort(function (a, b) {
      if (a === '未识别') return 1;
      if (b === '未识别') return -1;
      return b.localeCompare(a);
    });
    var totals = {}, grandB = 0, grandA = 0;
    terms.forEach(function (t) { totals[t] = { b: 0, a: 0 }; });
    var head = '<tr><th rowspan="2">年份</th>' +
      terms.map(function (t) { return '<th colspan="2">' + t + '</th>'; }).join('') +
      '<th colspan="2">合计</th></tr>';
    var sub = '<tr>' + terms.map(function () { return '<th>预算</th><th>决算</th>'; }).join('') +
      '<th>预算</th><th>决算</th></tr>';
    var body = years.map(function (y) {
      var row = byYear[y], tb = 0, ta = 0;
      var cells = terms.map(function (t) {
        var c = row[t];
        if (c) {
          if (c.hasB) { tb += c.b; totals[t].b += c.b; }
          if (c.hasA) { ta += c.a; totals[t].a += c.a; }
        }
        return '<td>' + (c && c.hasB ? fmt(c.b) : '—') + '</td>' +
               '<td>' + (c && c.hasA ? fmt(c.a) : '—') + '</td>';
      }).join('');
      grandB += tb; grandA += ta;
      return '<tr><td>' + esc(y) + '</td>' + cells +
        '<td>' + fmt(tb) + '</td><td>' + fmt(ta) + '</td></tr>';
    }).join('');
    var foot = '<tr><td>合计</td>' +
      terms.map(function (t) { return '<td>' + fmt(totals[t].b) + '</td><td>' + fmt(totals[t].a) + '</td>'; }).join('') +
      '<td>' + fmt(grandB) + '</td><td>' + fmt(grandA) + '</td></tr>';
    el.innerHTML = '<table class="ov ov-summary"><thead>' + head + sub + '</thead><tbody>' +
      body + '</tbody><tfoot>' + foot + '</tfoot></table>';
  }

  var lastRows = [];
  var chart = null;

  function chartTerm() {
    var sel = $('ov-chart-term');
    return sel ? sel.value : '全部';
  }

  function renderChart(rows) {
    var el = $('ov-chart');
    if (!el) return;
    var term = chartTerm();
    var note = $('ov-chart-note');
    if (note) note.textContent = term === '全部' ? '按年度合计：预算 vs 决算' : '仅看 ' + term;
    if (!window.echarts) {
      el.innerHTML = '<p class="hint">图表库未加载，表格仍可查看。</p>';
      return;
    }
    var map = {};
    (rows || []).forEach(function (r) {
      if (term !== '全部' && r.term !== term) return;
      if (!map[r.year]) map[r.year] = { b: 0, a: 0 };
      map[r.year].b += r.b;
      map[r.year].a += r.a;
    });
    var years = Object.keys(map).filter(function (y) { return y !== '未识别'; })
      .sort(function (a, b) { return Number(a) - Number(b); });
    if (map['未识别']) years.push('未识别');
    if (!years.length) {
      if (chart) { chart.dispose(); chart = null; }
      el.innerHTML = '<p class="hint">暂无数据。</p>';
      return;
    }
    var budget = [], actual = [], diff = [];
    years.forEach(function (y) {
      budget.push(map[y].b);
      actual.push(map[y].a);
      diff.push(Math.round((map[y].a - map[y].b) * 100) / 100);
    });
    if (!chart) {
      el.innerHTML = '';
      chart = window.echarts.init(el);
    }
    chart.setOption({
      tooltip: {
        trigger: 'axis',
        axisPointer: { type: 'shadow' },
        valueFormatter: function (v) { return fmt(v) + ' 元'; }
      },
      legend: { data: ['预算', '决算', '差额'], top: 0 },
      grid: { left: 64, right: 28, top: 42, bottom: 34 },
      xAxis: { type: 'category', data: years, axisLabel: { interval: 0 } },
      yAxis: {
        type: 'value', name: '元',
        axisLabel: {
          formatter: function (v) {
            return Math.abs(v) >= 10000 ? (v / 10000).toFixed(1) + '万' : v;
          }
        }
      },
      series: [
        { name: '预算', type: 'bar', data: budget, barMaxWidth: 42,
          itemStyle: { color: '#6f2d91' } },
        { name: '决算', type: 'bar', data: actual, barMaxWidth: 42,
          itemStyle: { color: '#2f9e6f' } },
        { name: '差额', type: 'line', data: diff, symbol: 'circle', symbolSize: 7,
          lineStyle: { color: '#c77c1e' }, itemStyle: { color: '#c77c1e' } }
      ]
    }, true);
  }

  function render(budgetRows, finalRows, bErr, fErr) {
    var rows = combine(budgetRows, finalRows);
    lastRows = rows;
    $('ov-src').textContent = '预算汇总 + 决算汇总';
    renderStats(rows);
    renderSummary(rows);
    renderChart(rows);
  }

  var loading = false;
  function load() {
    if (loading) return;
    loading = true;
    var msg = $('ov-msg');
    if (msg) msg.textContent = '读取中…';
    Promise.all([
      safe(callFinal({ action: 'exportSummary', kind: 'final' })),
      safe(callBudget({ action: 'exportSummary', kind: 'budget' }))
    ]).then(function (res) {
      var f = res[0].v, b = res[1].v;
      if (res[0].e && res[1].e) {
        throw new Error('预算汇总/决算汇总都读取失败：' + res[0].e.message + '；' + res[1].e.message);
      }
      render(b || [], f || [], res[1].e, res[0].e);
      var errs = [];
      if (res[0].e) errs.push('决算汇总');
      if (res[1].e) errs.push('预算汇总');
      if (msg) msg.textContent = '更新于 ' + new Date().toLocaleTimeString('zh-CN') +
        (errs.length ? '；' + errs.join('/') + '未读取（脚本可能未更新）' : '');
    }).catch(function (e) {
      if (msg) msg.textContent = '读取失败：' + e.message;
      $('ov-stats').innerHTML = '';
      if ($('ov-summary')) $('ov-summary').innerHTML = '<p class="hint">读取失败，无法汇总。</p>';
      if (chart) { chart.dispose(); chart = null; }
      if ($('ov-chart')) $('ov-chart').innerHTML = '<p class="hint">读取失败，无法绘图。</p>';
    }).finally(function () { loading = false; });
  }

  window.SJ_OVERVIEW = { load: load };

  var btn = $('ov-refresh');
  if (btn) btn.onclick = load;
  var termSel = $('ov-chart-term');
  if (termSel) termSel.onchange = function () { renderChart(lastRows); };
  window.addEventListener('resize', function () { if (chart) chart.resize(); });

  var view = $('view-overview');
  if (view && !view.hidden) load();
})();
