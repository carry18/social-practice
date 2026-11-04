/* 实践地图：ECharts 中国地图 + 金山文档统计数据
 * 数据：map_data.json（{updated, source, provinces: {省: {市: ["25年01月（2支）", ...]}}}）
 * 更新：页面内拖入新统计表 xlsx 即时预览，或用 tools/build_map_data.py 重建 map_data.json
 * 展示：默认「PPT 连线标注」模式（仿 实践地图.pptx：锚点圆点 + 折线 + 两侧标注），
 *       可切换到「侧栏列表」模式。
 */
(function () {
  'use strict';

  var $ = function (id) { return document.getElementById(id); };

  /* 主要城市经纬度（用于城市圆点；未收录的城市只出现在标注/明细里） */
  var CITY_COORDS = {
    '北京': [116.41, 39.90], '上海': [121.47, 31.23], '天津': [117.20, 39.13], '重庆': [106.55, 29.56],
    '哈尔滨': [126.53, 45.80], '长春': [125.32, 43.89], '沈阳': [123.43, 41.80], '大连': [121.61, 38.91],
    '石家庄': [114.51, 38.04], '太原': [112.55, 37.87], '呼和浩特': [111.75, 40.84],
    '济南': [117.00, 36.65], '青岛': [120.38, 36.07], '威海': [122.12, 37.51], '烟台': [121.45, 37.46],
    '南京': [118.78, 32.07], '苏州': [120.62, 31.30], '无锡': [120.30, 31.57], '徐州': [117.18, 34.26],
    '杭州': [120.15, 30.28], '宁波': [121.55, 29.88], '温州': [120.70, 28.00],
    '合肥': [117.28, 31.86], '芜湖': [118.38, 31.33],
    '福州': [119.30, 26.08], '厦门': [118.09, 24.48], '泉州': [118.68, 24.87],
    '南昌': [115.86, 28.68], '赣州': [114.93, 25.83],
    '郑州': [113.62, 34.75], '洛阳': [112.45, 34.62],
    '武汉': [114.30, 30.59], '宜昌': [111.29, 30.69], '襄阳': [112.12, 32.01],
    '长沙': [112.94, 28.23], '株洲': [113.13, 27.83],
    '广州': [113.26, 23.13], '深圳': [114.06, 22.55], '珠海': [113.58, 22.27], '佛山': [113.12, 23.02],
    '东莞': [113.75, 23.02], '汕头': [116.68, 23.35],
    '南宁': [108.32, 22.82], '桂林': [110.29, 25.27], '柳州': [109.42, 24.33],
    '海口': [110.20, 20.04], '三亚': [109.51, 18.25],
    '成都': [104.07, 30.57], '绵阳': [104.68, 31.47], '宜宾': [104.64, 28.75],
    '贵阳': [106.63, 26.65], '遵义': [106.93, 27.73],
    '昆明': [102.83, 24.88], '大理': [100.23, 25.59],
    '拉萨': [91.11, 29.66],
    '西安': [108.95, 34.27], '延安': [109.49, 36.60],
    '兰州': [103.82, 36.06], '酒泉': [98.52, 39.73], '敦煌': [94.66, 40.14],
    '西宁': [101.78, 36.62], '银川': [106.23, 38.49],
    '乌鲁木齐': [87.62, 43.83], '喀什': [75.99, 39.47]
  };

  var MUNICIPALITIES = ['北京市', '上海市', '天津市', '重庆市'];

  var chart = null;
  var geoReady = false;
  var DATA = null;          // 当前生效数据
  var PUBLISHED = null;     // 页面加载时的已发布数据（用于“还原”）
  var CITY_PROVINCE = {};   // 城市 → 省份 对照表
  var PROV_CENTER = {};     // 省份 → 几何中心 [lng, lat]（来自 china.json）
  var CALLOUT_LAYOUT = {};  // 省份 → 标注布局（每次 render 重算）

  /* ---------- 数据工具 ---------- */

  // "25年01月（2支）" → {yy:25, mm:1, n:2}
  function parseLabel(label) {
    var m = /^(\d{2})年(\d{2})月(?:（(\d+)支）)?$/.exec(label);
    if (!m) return null;
    return { yy: +m[1], mm: +m[2], n: m[3] ? +m[3] : 1 };
  }

  function inFilter(p, fy, fm) {
    if (fy && p.yy !== fy) return false;
    if (fm && p.mm !== fm) return false;
    return true;
  }

  // 按当前筛选聚合：{ prov: { total, cities: { city: {total, labels[]} } } }
  function aggregate(data, fy, fm) {
    var out = {};
    Object.keys(data.provinces).forEach(function (prov) {
      var cities = data.provinces[prov];
      Object.keys(cities).forEach(function (city) {
        var labels = [], total = 0;
        cities[city].forEach(function (label) {
          var p = parseLabel(label);
          if (p && inFilter(p, fy, fm)) { labels.push(label); total += p.n; }
        });
        if (labels.length) {
          out[prov] = out[prov] || { total: 0, cities: {} };
          out[prov].cities[city] = { total: total, labels: labels };
          out[prov].total += total;
        }
      });
    });
    return out;
  }

  function currentFilter() {
    var fy = $('f-year').value ? +$('f-year').value : 0;
    var fm = $('f-month').value ? +$('f-month').value : 0;
    return { fy: fy, fm: fm };
  }

  function calloutMode() {
    return $('callout-mode').checked;
  }

  /* ---------- PPT 连线标注布局 ---------- */

  var LINE_H = 15;        // 标注行高
  var WRAP_CHARS = 22;    // 标注每行最大字符数
  // 连线模式：地图放大居中，标注压在西部/海上的空白区（仿 PPT）
  var GEO_CALLOUT = { left: '13%', right: '23%', top: '2%', bottom: '2%' };
  var LABEL_X_L = 0.225, LABEL_X_R = 0.765;
  // 强制放左侧的省（仿 PPT：广东虽在东经112.5以东，标注仍放左下）
  var FORCE_LEFT = { '广东省': 1 };

  function wrapText(str, maxChars) {
    var parts = String(str).split('、'), lines = [], cur = '';
    parts.forEach(function (p, i) {
      var seg = (i ? '、' : '') + p;
      if (cur && (cur + seg).length > maxChars) {
        lines.push(cur);
        cur = seg.replace(/^、/, '　'); // 续行用全角空格开头对齐
      } else {
        cur += seg;
      }
    });
    if (cur) lines.push(cur);
    return lines;
  }

  function computeCalloutLayout(agg) {
    CALLOUT_LAYOUT = {};
    if (!chart) return;
    var H = chart.getHeight();
    var items = [];
    Object.keys(agg).forEach(function (prov) {
      var center = PROV_CENTER[prov];
      if (!center) return;
      var info = agg[prov];
      var lines = [{ t: prov, bold: true }];
      Object.keys(info.cities).forEach(function (city) {
        var text = city + '　' + info.cities[city].labels.join('、');
        wrapText(text, WRAP_CHARS).forEach(function (t) {
          lines.push({ t: t, bold: false });
        });
      });
      items.push({ prov: prov, center: center, lines: lines, h: lines.length * LINE_H + 6 });
    });
    // 按经度分左右两侧（仿 PPT：西部省份标左边，东部标右边）
    var left = items.filter(function (i) { return i.center[0] < 112.5 || FORCE_LEFT[i.prov]; })
                    .sort(function (a, b) { return b.center[1] - a.center[1]; });
    var right = items.filter(function (i) { return !(i.center[0] < 112.5 || FORCE_LEFT[i.prov]); })
                     .sort(function (a, b) { return b.center[1] - a.center[1]; });
    // 各侧按省份纬度自然落位：纬度映射到地图实际纵向范围（而非整个画布），
    // 使标注贴着各自省份的高度；重叠时向下最小推开，底部溢出再向上回顶
    [['L', left], ['R', right]].forEach(function (pair) {
      var side = pair[0], list = pair[1];
      if (!list.length) return;
      var topPad = 10, bottomPad = 10, minGap = 3;
      var W2 = chart.getWidth();
      var geoW = W2 * (1 - 0.13 - 0.23), geoH = H * 0.96;
      var mapH = Math.min(geoH, geoW / 1.33);   // 中国地图宽高比约 1.33
      var yTop = H * 0.02 + (geoH - mapH) / 2;  // geo 上下居中
      list.forEach(function (it) {
        it.y = yTop + (53.5 - it.center[1]) / (53.5 - 17.8) * mapH - it.h * 0.35;
        it.y = Math.max(topPad, Math.min(it.y, H - bottomPad - it.h));
      });
      var y = topPad;
      list.forEach(function (it) {
        it.y = Math.max(it.y, y);
        y = it.y + it.h + minGap;
      });
      var overflow = y - minGap - (H - bottomPad);
      if (overflow > 0) {
        var yb = H - bottomPad;
        for (var i = list.length - 1; i >= 0; i--) {
          var it = list[i];
          it.y = Math.min(it.y, yb - it.h);
          yb = it.y - minGap;
        }
      }
      list.forEach(function (it) {
        CALLOUT_LAYOUT[it.prov] = { side: side, y: it.y, lines: it.lines, h: it.h };
      });
    });
  }

  // custom series：为每个省绘制 锚点圆点 + 标注文字；
  // 折线默认不画，悬停/点击钉住该省时才显示（renderItem 读取 HOVER_PROV/PINNED_PROV）
  var CALLOUT_DATA = [];  // custom series 数据（renderItem 用 dataIndex 取省名）
  var HOVER_PROV = null, PINNED_PROV = null, SHOW_ALL_LINES = false;

  function lineVisible(prov) {
    return SHOW_ALL_LINES || prov === HOVER_PROV || prov === PINNED_PROV;
  }

  function calloutRenderItem(params, api) {
    var prov = CALLOUT_DATA[params.dataIndex] && CALLOUT_DATA[params.dataIndex].name;
    var lay = prov && CALLOUT_LAYOUT[prov];
    if (!lay) return { type: 'group', children: [] };
    var anchor = api.coord([api.value(0), api.value(1)]);
    var W = chart.getWidth();
    var isL = lay.side === 'L';
    var textX = isL ? W * LABEL_X_L : W * LABEL_X_R;
    var attachY = lay.y + 8; // 折线接到省份名那一行
    // 转折通道距标注边缘的距离随省份略有变化，避免所有折线一个模子
    var vary = 12 + (prov.charCodeAt(0) % 3) * 9;
    var elbowX = textX + (isL ? vary : -vary);
    var lineEndX = textX + (isL ? 5 : -5);
    var active = lineVisible(prov);

    var children = [];
    if (active) {
      children.push({
        type: 'polyline', silent: true, z: 6,
        shape: { points: [anchor, [elbowX, anchor[1]], [elbowX, attachY], [lineEndX, attachY]] },
        style: { stroke: '#d63384', lineWidth: 1.6, fill: 'none', lineJoin: 'round' }
      });
    }
    children.push({
      type: 'circle', z: 4,
      shape: { cx: anchor[0], cy: anchor[1], r: active ? 5 : 4 },
      style: { fill: active ? '#d63384' : '#fff', stroke: '#d63384',
               lineWidth: 1.6, opacity: active ? 1 : 0.85 }
    });
    lay.lines.forEach(function (line, i) {
      children.push({
        type: 'text', z: 5,
        style: {
          x: textX, y: lay.y + i * LINE_H,
          text: line.t,
          fontSize: line.bold ? 12.5 : 10.5,
          fontWeight: line.bold || active ? 'bold' : 'normal',
          fill: line.bold ? '#6f2d91' : (active ? '#6f2d91' : '#444'),
          align: isL ? 'right' : 'left',
          verticalAlign: 'top'
        }
      });
    });
    return { type: 'group', children: children };
  }

  // 只重绘 custom 系列（悬停状态变化时调用）
  function refreshCallouts() {
    if (!chart || !calloutMode()) return;
    chart.setOption({ series: [{ id: 'callout', data: CALLOUT_DATA.slice() }] });
  }

  /* ---------- 渲染 ---------- */

  function render() {
    if (!geoReady || !DATA) return;
    var f = currentFilter();
    var agg = aggregate(DATA, f.fy, f.fm);
    var callout = calloutMode();

    var maxV = 1, provTotal = 0, citySet = {}, squadTotal = 0;
    Object.keys(agg).forEach(function (prov) {
      maxV = Math.max(maxV, agg[prov].total);
      provTotal++;
      squadTotal += agg[prov].total;
      Object.keys(agg[prov].cities).forEach(function (c) { citySet[c] = 1; });
    });

    var mapSeries = Object.keys(agg).map(function (prov) {
      return { name: prov, value: agg[prov].total };
    });

    var scatter = [];
    Object.keys(citySet).forEach(function (city) {
      var key = city.replace(/市$/, '');
      if (CITY_COORDS[key]) scatter.push({ name: city, value: CITY_COORDS[key] });
    });

    // 连线标注模式：隐藏侧栏、加高画布、地图居中留两侧边距
    var mapEl = $('map');
    mapEl.style.height = callout ? '780px' : '640px';
    $('side').style.display = callout ? 'none' : '';
    chart.resize();
    if (callout) computeCalloutLayout(agg);

    var geoOpt = callout
      ? GEO_CALLOUT
      : { zoom: 1.05, top: '8%' };

    var series = [
      { type: 'map', map: 'china', geoIndex: 0, data: mapSeries },
      {
        type: 'scatter', coordinateSystem: 'geo',
        symbolSize: 8,
        itemStyle: { color: '#d63384', borderColor: '#fff', borderWidth: 1.5,
                     shadowColor: 'rgba(214,51,132,.5)', shadowBlur: 4 },
        emphasis: { label: { show: !callout, position: 'right', fontSize: 11,
                             color: '#333', formatter: '{b}' },
                    scale: 1.4 },
        label: { show: false },
        tooltip: { show: false },
        zlevel: 2,
        data: scatter
      }
    ];
    if (callout) {
      CALLOUT_DATA = Object.keys(agg)
        .filter(function (p) { return PROV_CENTER[p]; })
        .map(function (p) { return { name: p, value: PROV_CENTER[p] }; });
      series.push({
        id: 'callout',
        type: 'custom', coordinateSystem: 'geo',
        renderItem: calloutRenderItem,
        z: 6,
        data: CALLOUT_DATA
      });
    }

    chart.setOption({
      backgroundColor: '#fff',
      tooltip: {
        trigger: 'item',
        formatter: function (p) {
          var prov = p.name;
          var info = agg[prov];
          if (p.seriesType === 'effectScatter' || p.seriesType === 'scatter') return p.name;
          if (!info) return prov + '<br><span style="color:#999">筛选范围内暂无实践记录</span>';
          var html = '<b>' + prov + '</b>（' + info.total + ' 支次）';
          Object.keys(info.cities).forEach(function (city) {
            var c = info.cities[city];
            html += '<br><b>' + city + '</b>　' + c.labels.join('、');
          });
          return html;
        }
      },
      visualMap: Object.assign({
        min: 0, max: maxV, calculable: true,
        seriesIndex: 0,
        text: ['支队次多', '少'],
        inRange: { color: ['#f3eaf8', '#cfa9e3', '#9b59b6', '#6f2d91'] },
        outOfRange: { color: ['#f0f0f0'] },
        textStyle: { fontSize: 11, color: '#666' }
      }, callout ? { orient: 'horizontal', left: '31%', bottom: 6 }
                 : { left: 12, bottom: 12 }),
      geo: Object.assign({
        map: 'china', roam: true, scaleLimit: { min: 0.8, max: 6 },
        itemStyle: { areaColor: '#f0f0f0', borderColor: '#fff', borderWidth: 1 },
        emphasis: { label: { show: !callout, fontSize: 12, color: '#333' },
                    itemStyle: { areaColor: '#b37fc9' } },
        label: { show: false }
      }, geoOpt),
      series: series
    }, { notMerge: true });

    /* 侧栏明细（连线模式下隐藏但内容照算，切换即用） */
    var side = Object.keys(agg).sort(function (a, b) { return agg[b].total - agg[a].total; })
      .map(function (prov) {
        var rows = Object.keys(agg[prov].cities).map(function (city) {
          var c = agg[prov].cities[city];
          return '<div class="city"><b>' + city + '</b>　<span class="months">' +
                 c.labels.join('、') + '</span></div>';
        }).join('');
        return '<h3>' + prov + '（' + agg[prov].total + ' 支次）</h3>' + rows;
      }).join('');
    $('side').innerHTML = side || '<p class="hint">筛选范围内暂无记录</p>';

    var ymRange = (f.fy ? f.fy + ' 年' : '全部年份') + (f.fm ? ' ' + f.fm + ' 月' : '');
    $('summary').textContent = ymRange + '：覆盖 ' + provTotal + ' 个省级行政区、' +
      Object.keys(citySet).length + ' 个城市，共 ' + squadTotal + ' 支次';
  }

  /* ---------- 筛选器 ---------- */

  function buildFilters() {
    var years = {}, months = {};
    Object.keys(DATA.provinces).forEach(function (prov) {
      Object.keys(DATA.provinces[prov]).forEach(function (city) {
        DATA.provinces[prov][city].forEach(function (label) {
          var p = parseLabel(label);
          if (p) { years[p.yy] = 1; months[p.mm] = 1; }
        });
      });
    });
    var ys = $('f-year');
    ys.innerHTML = '<option value="">全部年份</option>' + Object.keys(years).sort().map(function (y) {
      return '<option value="' + y + '">20' + (y < 10 ? '0' + y : y) + ' 年</option>';
    }).join('');
    var ms = $('f-month');
    ms.innerHTML = '<option value="">全部月份</option>' + Object.keys(months).sort(function (a, b) { return a - b; }).map(function (m) {
      return '<option value="' + m + '">' + m + ' 月</option>';
    }).join('');
  }

  /* ---------- xlsx 上传解析（与 tools/build_map_data.py 同口径） ----------
   * 支持金山《工作统计》工作簿：
   *  - 「YYYY年…-寒假/学期中/暑期」支队实践表：地点（多城市）+ 参访时间（出发月），
   *    「实践取消」行不计，一行多城市每城记 1 支；
   *  - 「开放日/行业之声」表：企业名称 + 开放日时间，每个活动记北京市 1 支；
   *  - 其他表（如数据统计）自动跳过。 */

  // 省份名与城市同名会误判的城市
  var AMBIGUOUS_CITIES = { '吉林': 1 };
  var CITY_NAMES = [];

  function extractCities(text) {
    var s = String(text == null ? '' : text).replace(/\s+/g, '');
    if (!s) return [];
    var found = [], i = 0;
    outer: while (i < s.length) {
      for (var k = 0; k < CITY_NAMES.length; k++) {
        var name = CITY_NAMES[k];
        if (s.substr(i, name.length) === name) { found.push(name); i += name.length; continue outer; }
      }
      i++;
    }
    var seen = {}, out = [];
    found.forEach(function (c) { if (!seen[c]) { seen[c] = 1; out.push(c); } });
    return out;
  }

  function parseStartMonth(v) {
    if (v == null || v === '') return 0;
    if (v instanceof Date) return v.getMonth() + 1;
    if (typeof v === 'number') { // Excel 序列日期
      var d = new Date(Math.round((v - 25569) * 86400000));
      return isNaN(d) ? 0 : d.getUTCMonth() + 1;
    }
    var m = /(\d{1,2})\s*[月.]/.exec(String(v));
    if (!m) return 0;
    var mm = +m[1];
    return mm >= 1 && mm <= 12 ? mm : 0;
  }

  function parseFullDate(v) {
    if (v instanceof Date) {
      var y = v.getFullYear();
      return (y >= 2024 && y <= 2035) ? { yy: y % 100, mm: v.getMonth() + 1 } : null;
    }
    if (typeof v === 'number') {
      var d = new Date(Math.round((v - 25569) * 86400000));
      if (isNaN(d)) return null;
      var y2 = d.getUTCFullYear();
      return (y2 >= 2024 && y2 <= 2035) ? { yy: y2 % 100, mm: d.getUTCMonth() + 1 } : null;
    }
    if (v == null) return null;
    var m = /(20\d{2})[-/.年](\d{1,2})/.exec(String(v));
    return m ? { yy: +m[1] % 100, mm: +m[2] } : null;
  }

  function findCol(headerRow) {
    var cols = {};
    headerRow.forEach(function (cell, idx) {
      var t = String(cell == null ? '' : cell);
      if (/^地点$|实践地点|地点/.test(t) && cols.loc == null) cols.loc = idx;
      if (/参访时间|实践时间|时间/.test(t) && cols.time == null) cols.time = idx;
      if (/进度|状态/.test(t) && cols.prog == null) cols.prog = idx;
      if (/企业名称/.test(t) && cols.ent == null) cols.ent = idx;
      if (/开放日时间/.test(t) && cols.otime == null) cols.otime = idx;
    });
    return cols;
  }

  function parsePracticeSheet(name, aoa, records, warnings) {
    var ym = /^(\d{4})年/.exec(aoa.name);
    if (!ym) return;
    var year = +ym[1];
    var cols = findCol(aoa.rows[0] || []);
    if (cols.loc == null || cols.time == null) {
      warnings.push('工作表「' + aoa.name + '」缺少地点/时间列，已跳过');
      return;
    }
    for (var r = 1; r < aoa.rows.length; r++) {
      var row = aoa.rows[r] || [];
      if (cols.prog != null && row[cols.prog] && String(row[cols.prog]).indexOf('取消') >= 0) continue;
      var loc = String(row[cols.loc] == null ? '' : row[cols.loc]).trim();
      var mm = parseStartMonth(row[cols.time]);
      if (!loc || !mm) continue;
      var cities = extractCities(loc);
      if (!cities.length) {
        // 纯数字的是汇总尾行（如 '', '5', '65'），不告警
        if (!/^[\d.\s]*$/.test(loc))
          warnings.push('「' + aoa.name + '」第' + (r + 1) + '行：地点「' + loc + '」未识别出已知城市');
        continue;
      }
      cities.forEach(function (city) {
        var prov = CITY_PROVINCE[city];
        var cname = MUNICIPALITIES.indexOf(prov) >= 0 ? prov : city + '市';
        records.push({ yy: year % 100, mm: mm, city: cname, prov: prov, n: 1 });
      });
    }
  }

  function parseOpendaySheet(aoa, records, warnings) {
    var cols = null, start = 1;
    for (var i = 0; i < Math.min(3, aoa.rows.length); i++) {
      var c = findCol(aoa.rows[i] || []);
      if (c.ent != null && c.otime != null) { cols = c; start = i + 1; break; }
    }
    if (!cols) {
      warnings.push('工作表「' + aoa.name + '」缺少企业名称/开放日时间列，已跳过');
      return;
    }
    for (var r = start; r < aoa.rows.length; r++) {
      var row = aoa.rows[r] || [];
      var name = String(row[cols.ent] == null ? '' : row[cols.ent]).trim();
      if (!name || name === '企业名称') continue;
      var t = parseFullDate(row[cols.otime]);
      if (!t) {
        warnings.push('「' + aoa.name + '」第' + (r + 1) + '行：「' + name + '」无有效开放日时间，未计入');
        continue;
      }
      records.push({ yy: t.yy, mm: t.mm, city: '北京市', prov: '北京市', n: 1 });
    }
  }

  function buildProvinces(records) {
    var aggMap = {}, order = [];
    records.forEach(function (r) {
      var key = r.prov + '|' + r.city + '|' + r.yy + '|' + r.mm;
      if (!aggMap[key]) { aggMap[key] = { prov: r.prov, city: r.city, yy: r.yy, mm: r.mm, n: 0 }; order.push(key); }
      aggMap[key].n += r.n;
    });
    order.sort();
    var provinces = {};
    order.forEach(function (key) {
      var a = aggMap[key];
      var label = (a.yy < 10 ? '0' + a.yy : a.yy) + '年' + (a.mm < 10 ? '0' + a.mm : a.mm) + '月' +
                  (a.n > 1 ? '（' + a.n + '支）' : '');
      provinces[a.prov] = provinces[a.prov] || {};
      provinces[a.prov][a.city] = provinces[a.prov][a.city] || [];
      provinces[a.prov][a.city].push(label);
    });
    Object.keys(provinces).forEach(function (p) {
      Object.keys(provinces[p]).forEach(function (c) { provinces[p][c].sort(); });
    });
    return provinces;
  }

  function handleXlsx(file) {
    var status = $('upload-status');
    status.className = ''; status.textContent = '解析中…';
    var reader = new FileReader();
    reader.onload = function (e) {
      try {
        var wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array', cellDates: true });
        var records = [], warnings = [];
        wb.SheetNames.forEach(function (name) {
          var aoa = {
            name: name,
            rows: XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: null })
          };
          if (/开放日|行业之声/.test(name)) parseOpendaySheet(aoa, records, warnings);
          else if (/^\d{4}年/.test(name)) parsePracticeSheet(name, aoa, records, warnings);
          else warnings.push('工作表「' + name + '」不是实践/开放日表，已跳过');
        });
        if (!records.length) throw new Error('没有解析出任何记录');

        DATA = {
          updated: new Date().toISOString().slice(0, 10),
          source: file.name,
          provinces: buildProvinces(records)
        };
        buildFilters();
        render();
        $('upload-actions').style.display = 'block';
        status.className = 'ok';
        status.textContent = '✅ 已按「' + file.name + '」重绘：' + records.length + ' 条记录' +
          (warnings.length ? '\n⚠ ' + warnings.length + ' 条提示：\n' + warnings.slice(0, 10).join('\n') : '') +
          '\n确认无误后点「下载 map_data.json」，替换仓库里的 map_data.json 并提交，即可正式发布。';
      } catch (err) {
        status.className = 'err';
        status.textContent = '❌ ' + err.message;
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function downloadJson() {
    var blob = new Blob([JSON.stringify(DATA, null, 2)], { type: 'application/json' });
    var a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'map_data.json';
    a.click();
    URL.revokeObjectURL(a.href);
  }

  /* ---------- 初始化 ---------- */

  function init() {
    chart = echarts.init($('map'));
    Promise.all([
      fetch('china.json').then(function (r) { return r.json(); }),
      fetch('map_data.json').then(function (r) { return r.json(); }),
      fetch('city_province.json').then(function (r) { return r.json(); })
    ]).then(function (res) {
      var geojson = res[0];
      echarts.registerMap('china', geojson);
      geojson.features.forEach(function (f) {
        var p = f.properties || {};
        if (p.name && (p.centroid || p.center)) PROV_CENTER[p.name] = p.centroid || p.center;
      });
      geoReady = true;
      DATA = PUBLISHED = res[1];
      CITY_PROVINCE = res[2];
      CITY_NAMES = Object.keys(CITY_PROVINCE)
        .filter(function (k) { return !AMBIGUOUS_CITIES[k]; })
        .sort(function (a, b) { return b.length - a.length; });
      buildFilters();
      render();
      $('summary').textContent += '　｜　数据更新：' + (DATA.updated || '未知') +
        '（' + (DATA.source || '') + '）';
    }).catch(function (err) {
      $('map').innerHTML = '<p class="hint" style="padding:40px">⚠ 数据加载失败：' + err.message +
        '<br>请通过 http(s) 访问本页（file:// 直开无法读取 JSON），并确认 china.json / map_data.json 与本页同目录。</p>';
    });

    $('f-year').addEventListener('change', render);
    $('f-month').addEventListener('change', render);
    $('callout-mode').addEventListener('change', render);
    $('all-lines').addEventListener('change', function () {
      SHOW_ALL_LINES = $('all-lines').checked;
      refreshCallouts();
    });

    /* 悬停省份（地图区域或四周标注）→ 显示连线；点击 → 钉住/取消钉住 */
    var clearTimer = null;
    function provFromEvent(params) {
      if (params.seriesType === 'map') return params.name;
      if (params.seriesType === 'custom')
        return CALLOUT_DATA[params.dataIndex] && CALLOUT_DATA[params.dataIndex].name;
      return null;
    }
    chart.on('mouseover', function (params) {
      var prov = provFromEvent(params);
      if (!prov || !CALLOUT_LAYOUT[prov]) return;
      clearTimeout(clearTimer);
      if (HOVER_PROV !== prov) { HOVER_PROV = prov; refreshCallouts(); }
    });
    chart.on('mouseout', function (params) {
      if (!provFromEvent(params)) return;
      clearTimeout(clearTimer);
      clearTimer = setTimeout(function () {
        HOVER_PROV = null; refreshCallouts();
      }, 120);
    });
    chart.getZr().on('globalout', function () {
      clearTimeout(clearTimer);
      HOVER_PROV = null; refreshCallouts();
    });
    chart.on('click', function (params) {
      var prov = provFromEvent(params);
      if (!prov || !CALLOUT_LAYOUT[prov]) return;
      PINNED_PROV = (PINNED_PROV === prov) ? null : prov;
      refreshCallouts();
    });

    $('export-png').addEventListener('click', function () {
      var url = chart.getDataURL({ pixelRatio: 2, backgroundColor: '#fff' });
      var a = document.createElement('a');
      var f = currentFilter();
      a.href = url;
      a.download = '实践地图' + (f.fy ? '_20' + (f.fy < 10 ? '0' + f.fy : f.fy) + '年' : '') +
                   (f.fm ? '_' + f.fm + '月' : '') + '.png';
      a.click();
    });

    var drop = $('drop'), fileInput = $('file');
    drop.addEventListener('click', function () { fileInput.click(); });
    drop.addEventListener('dragover', function (e) { e.preventDefault(); drop.classList.add('hover'); });
    drop.addEventListener('dragleave', function () { drop.classList.remove('hover'); });
    drop.addEventListener('drop', function (e) {
      e.preventDefault(); drop.classList.remove('hover');
      if (e.dataTransfer.files.length) handleXlsx(e.dataTransfer.files[0]);
    });
    fileInput.addEventListener('change', function () {
      if (fileInput.files.length) handleXlsx(fileInput.files[0]);
      fileInput.value = '';
    });
    $('download-json').addEventListener('click', downloadJson);
    $('revert').addEventListener('click', function () {
      DATA = PUBLISHED;
      buildFilters();
      render();
      $('upload-actions').style.display = 'none';
      var status = $('upload-status');
      status.className = ''; status.textContent = '已还原为已发布数据。';
    });

    var resizeTimer = null;
    window.addEventListener('resize', function () {
      clearTimeout(resizeTimer);
      resizeTimer = setTimeout(render, 200); // 重算标注布局
    });
  }

  document.addEventListener('DOMContentLoaded', init);
})();
