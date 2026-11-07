/* 企业参考侧栏：跟随预算表当前行，按地点筛选历史合作单位，一键写入 C 列。 */
(function () {
  if (typeof window === "undefined") return;

  const state = {
    open: false,
    row: -1,
    data: null,
    loading: false,
    error: "",
    q: "",
    group: "",
    view: "all",
    cities: [],
    province: "",
    base: [],
    fallback: false,
  };

  let panel = null;
  let listEl = null;
  let cityEl = null;
  let countEl = null;
  let subEl = null;
  let rowLabelEl = null;
  let subTextEl = null;
  let qEl = null;
  let groupEl = null;
  let toastEl = null;
  let toastTimer = null;

  const LEVEL_ORDER = { "集团本部": 0, "成员单位": 1, "分支机构": 2, "科研机构": 3, "独立单位": 4, "部队": 5 };

  function esc(s) {
    return String(s == null ? "" : s)
      .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  function bridge() { return window.BudgetBridge || null; }

  function build() {
    if (panel) return;
    panel = document.createElement("aside");
    panel.id = "ent-panel";
    panel.innerHTML =
      '<div class="ent-head">' +
        '<div class="ent-title">🏢 <span id="ent-city">全部企业</span> ' +
          '<span class="ent-n" id="ent-count"></span>' +
          '<button class="ent-close" title="关闭">✕</button></div>' +
        '<div class="ent-sub" id="ent-sub"><button class="ent-nav" data-dir="-1" title="上一行">‹</button>' +
          '<span id="ent-row-label">第 1 行</span>' +
          '<button class="ent-nav" data-dir="1" title="下一行">›</button>' +
          '<span id="ent-sub-text"></span></div>' +
      '</div>' +
      '<div class="ent-search"><span>🔍</span>' +
        '<input id="ent-q" placeholder="搜企业 / 集团 / 城市"></div>' +
      '<div class="ent-chips">' +
        '<button class="ent-chip on" data-view="all">全部</button>' +
        '<button class="ent-chip mil" data-view="defense">军工</button>' +
        '<button class="ent-chip" data-view="troops">部队</button>' +
        '<select class="ent-group" id="ent-group"><option value="">全部集团</option></select>' +
      '</div>' +
      '<div class="ent-list" id="ent-list"></div>' +
      '<div class="ent-foot-bar"><span class="ent-live"></span> 合作单位库 · 更新 2026-11-07 · 不含联系人</div>';
    document.body.appendChild(panel);

    listEl = panel.querySelector("#ent-list");
    cityEl = panel.querySelector("#ent-city");
    countEl = panel.querySelector("#ent-count");
    subEl = panel.querySelector("#ent-sub");
    rowLabelEl = panel.querySelector("#ent-row-label");
    subTextEl = panel.querySelector("#ent-sub-text");
    qEl = panel.querySelector("#ent-q");
    groupEl = panel.querySelector("#ent-group");

    panel.querySelector(".ent-close").addEventListener("click", close);
    qEl.addEventListener("input", () => { state.q = qEl.value.trim(); renderList(); });
    groupEl.addEventListener("change", () => { state.group = groupEl.value; renderList(); });
    subEl.addEventListener("click", function (e) {
      const btn = e.target.closest(".ent-nav");
      if (!btn) return;
      const next = state.row + Number(btn.getAttribute("data-dir"));
      const b = bridge();
      if (!b || next < 0 || next >= b.rowCount()) return;
      state.row = next;
      state.q = "";
      state.group = "";
      state.view = "all";
      if (qEl) qEl.value = "";
      panel.querySelectorAll(".ent-chip").forEach(function (x) {
        x.classList.toggle("on", x.getAttribute("data-view") === "all");
      });
      render();
    });
    panel.querySelectorAll(".ent-chip").forEach(function (b) {
      b.addEventListener("click", function () {
        state.view = b.getAttribute("data-view");
        panel.querySelectorAll(".ent-chip").forEach(function (x) {
          x.classList.toggle("on", x === b);
        });
        renderList();
      });
    });
    window.addEventListener("keydown", function (e) {
      if (e.key === "Escape" && state.open) close();
    });
  }

  function toast(msg) {
    if (!toastEl) {
      toastEl = document.createElement("div");
      toastEl.className = "ent-toast";
      document.body.appendChild(toastEl);
    }
    toastEl.textContent = msg;
    toastEl.classList.add("show");
    clearTimeout(toastTimer);
    toastTimer = setTimeout(function () { toastEl.classList.remove("show"); }, 1800);
  }

  function load() {
    if (state.data || state.loading) return Promise.resolve();
    state.loading = true;
    state.error = "";
    return fetch("enterprise_index.json", { cache: "no-cache" })
      .then(function (r) {
        if (!r.ok) throw new Error("HTTP " + r.status);
        return r.json();
      })
      .then(function (json) {
        state.data = json;
        state.loading = false;
        render();
      })
      .catch(function (e) {
        state.loading = false;
        state.error = "企业库加载失败：" + (e && e.message ? e.message : e);
        render();
      });
  }

  function parseCities(B) {
    const out = { cities: [], province: "" };
    if (!state.data) return out;
    const raw = String(B || "").trim();
    if (!raw) return out;
    const known = Object.keys(state.data.cities || {});
    const parts = raw.split(/[、,，/;；\s]+/).filter(Boolean);
    parts.forEach(function (p) {
      const q = p.replace(/市$/, "");
      let hit = known.find(function (c) { return c === p || c === q || c.replace(/市$/, "") === q; });
      if (!hit) hit = known.find(function (c) { return p.indexOf(c) >= 0 || q.indexOf(c) >= 0; });
      if (hit && out.cities.indexOf(hit) < 0) out.cities.push(hit);
    });
    if (!out.province) {
      const provs = {};
      Object.keys(state.data.cities || {}).forEach(function (c) {
        const p = state.data.cities[c].province;
        if (p) provs[p] = true;
      });
      Object.keys(provs).forEach(function (p) {
        const short = p.replace(/省|市|自治区|特别行政区/g, "");
        if (short && raw.indexOf(short) >= 0) out.province = p;
      });
    }
    if (!out.province && out.cities.length) {
      out.province = state.data.cities[out.cities[0]].province || "";
    }
    return out;
  }

  function baseList() {
    if (!state.data) return [];
    const all = state.data.companies || [];
    const picked = parseCities((bridge() && bridge().getRow(state.row) || {}).B);
    state.cities = picked.cities;
    state.province = picked.province;
    state.fallback = false;
    let base = all;
    if (picked.cities.length) {
      base = all.filter(function (c) { return picked.cities.indexOf(c.city) >= 0; });
      if (!base.length && picked.province) {
        base = all.filter(function (c) { return c.province === picked.province; });
        state.fallback = true;
      }
    } else if (picked.province) {
      base = all.filter(function (c) { return c.province === picked.province; });
    }
    state.base = base;
    return base;
  }

  function filtered() {
    let list = state.base.slice();
    const q = state.q.toLowerCase();
    if (q) {
      list = list.filter(function (c) {
        return (c.name + " " + c.raw + " " + c.group + " " + c.city + " " + c.province).toLowerCase().indexOf(q) >= 0;
      });
    }
    if (state.view === "defense") list = list.filter(function (c) { return c.defense; });
    if (state.view === "troops") list = list.filter(function (c) { return c.troops; });
    if (state.group) list = list.filter(function (c) { return c.group === state.group; });
    list.sort(function (a, b) {
      const la = LEVEL_ORDER[a.level] == null ? 9 : LEVEL_ORDER[a.level];
      const lb = LEVEL_ORDER[b.level] == null ? 9 : LEVEL_ORDER[b.level];
      if (la !== lb) return la - lb;
      if (b.group_count !== a.group_count) return b.group_count - a.group_count;
      return a.name.localeCompare(b.name, "zh");
    });
    return list;
  }

  function sourceLabel(s) {
    if (!s || s === "空" || s === "已有") return "";
    const map = {
      "企业台账": "台账匹配", "院所字典": "资料", "省会推断": "省会推断",
      "城市反推": "城市反推", "集团总部": "总部",
      "联网确认(CETC官网)": "官网已核", "联网确认": "已核"
    };
    return "地点：" + (map[s] || s);
  }

  function currentC() {
    const r = bridge() && bridge().getRow(state.row);
    return String((r && r.C) || "");
  }

  function isAdded(c) {
    const parts = currentC().split(/[、,，]/).map(function (s) { return s.trim(); }).filter(Boolean);
    return parts.indexOf(c.write_name) >= 0 || parts.indexOf(c.raw) >= 0 || parts.indexOf(c.name) >= 0;
  }

  function insert(c) {
    const b = bridge();
    if (!b) return;
    const r = b.getRow(state.row);
    if (!r) return;
    if (isAdded(c)) { toast("已在本行"); return; }
    const parts = currentC().split(/[、,，]/).map(function (s) { return s.trim(); }).filter(Boolean);
    parts.push(c.write_name);
    b.setPoint(state.row, parts.join("、"));
    toast("已写入第 " + (state.row + 1) + " 行：" + c.write_name);
    renderList();
  }

  function render() {
    if (!panel || !state.open) return;
    if (state.error) {
      cityEl.textContent = "企业库";
      countEl.textContent = "";
      rowLabelEl.textContent = "";
      subTextEl.textContent = state.error;
      listEl.innerHTML = '<div class="ent-empty">' + esc(state.error) + '</div>';
      return;
    }
    if (!state.data) {
      cityEl.textContent = "加载中…";
      countEl.textContent = "";
      rowLabelEl.textContent = "";
      subTextEl.textContent = "正在读取企业台账";
      listEl.innerHTML = '<div class="ent-empty">加载中…</div>';
      return;
    }

    const base = baseList();
    const row = bridge() && bridge().getRow(state.row);
    const place = (row && String(row.B || "").trim()) || "未填地点";
    let label = "全部企业";
    if (state.cities.length) label = state.cities.join("、");
    else if (state.province) label = state.province;
    if (state.fallback) label += " · 同省";
    cityEl.textContent = label;
    rowLabelEl.textContent = "第 " + (state.row + 1) + " 行";
    subTextEl.textContent = " · 地点「" + place + "」" +
      (state.fallback ? " · 该市暂无记录，显示同省" : "");

    // 集团下拉：只列当前地点/省出现的集团，按数量排序
    const counts = {};
    base.forEach(function (c) { counts[c.group] = (counts[c.group] || 0) + 1; });
    const groups = Object.keys(counts).sort(function (a, b) { return counts[b] - counts[a] || a.localeCompare(b, "zh"); }).slice(0, 80);
    if (state.group && groups.indexOf(state.group) < 0) state.group = "";
    groupEl.innerHTML = '<option value="">全部集团</option>' +
      groups.map(function (g) {
        return '<option value="' + esc(g) + '">' + esc(g) + " (" + counts[g] + ")</option>";
      }).join("");
    groupEl.value = state.group;

    renderList();
  }

  function renderList() {
    if (!listEl || !state.data) return;
    const list = filtered();
    countEl.textContent = "· " + list.length + " 家";
    if (!list.length) {
      listEl.innerHTML = '<div class="ent-empty">' +
        (state.q ? "没找到「" + esc(state.q) + "」" : "该地点暂无合作记录") +
        '<br><span style="font-size:11px;color:#bbb">可换个城市，或直接手动填 C 列</span></div>';
      return;
    }
    const max = 80;
    listEl.innerHTML = list.slice(0, max).map(function (c) {
      const added = isAdded(c);
      const badge = c.defense ? '<span class="ent-badge">军工</span>' :
        (c.troops ? '<span class="ent-badge troop">部队</span>' : "");
      const region = [c.province, c.city].filter(Boolean).join(" · ");
      return '<div class="ent-card">' +
        '<div class="ent-name"><span class="ent-nm">' + esc(c.name) + '</span>' + badge + '</div>' +
        '<div class="ent-meta">' + esc(c.group) + " · " + esc(c.level) +
          (region ? " · " + esc(region) : "") + '</div>' +
        '<div class="ent-foot"><span class="ent-src">' +
          esc(sourceLabel(c.source)) + '</span>' +
          '<button class="ent-add' + (added ? " done" : "") + '" data-id="' + esc(c.id) + '">' +
          (added ? "已添加" : "写入点位") + '</button></div>' +
      '</div>';
    }).join("") + (list.length > max ? '<div class="ent-empty">仅显示前 ' + max + ' 家，输入关键词继续缩小范围</div>' : "");

    listEl.querySelectorAll(".ent-add").forEach(function (btn) {
      btn.addEventListener("click", function () {
        const c = state.data.companies.find(function (x) { return x.id === btn.getAttribute("data-id"); });
        if (c) insert(c);
      });
    });
  }

  function open(rowIndex) {
    if (!bridge()) { alert("预算表还没准备好，请刷新页面"); return; }
    if (rowIndex < 0 || rowIndex >= bridge().rowCount()) return;
    state.row = rowIndex;
    state.open = true;
    state.q = "";
    state.group = "";
    state.view = "all";
    state.base = [];
    build();
    panel.classList.add("open");
    if (qEl) qEl.value = "";
    panel.querySelectorAll(".ent-chip").forEach(function (x) {
      x.classList.toggle("on", x.getAttribute("data-view") === "all");
    });
    if (state.data) render(); else load();
  }

  function close() {
    state.open = false;
    if (panel) panel.classList.remove("open");
  }

  window.EnterprisePanel = { open: open, close: close, refresh: render };

  window.addEventListener("budget:row-change", function (e) {
    if (!state.open || !e.detail) return;
    if (e.detail.i === state.row && e.detail.c === "B") render();
  });
  window.addEventListener("budget:render", function () {
    if (!state.open) return;
    if (!bridge() || state.row >= bridge().rowCount()) close();
    else render();
  });
})();
