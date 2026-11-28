/* 前端口令门禁（GitHub Pages 纯静态站点没有后端，只能做到"挡人"这一层）
 *
 * 说明：
 *   1. 这里存的是口令的 SHA-256，不是明文；拿到文件也看不到原口令。
 *   2. 登录状态存在 sessionStorage，默认 7 天（AUTH.days），关标签页不丢。
 *   3. 重要：这只能挡住随手打开链接的人，**挡不住直接访问 fares.json / ledger-map.json 等文件**。
 *      需要真正的安全，得把敏感数据放到需要 token 的云函数后面。
 *
 * 改口令：在本地跑
 *   python3 -c "import hashlib;print(hashlib.sha256(b'你的新口令').encode().hexdigest())"
 * 把结果填进下面的 hash。
 */
(function () {
  var AUTH = {
    hash: "658658fdc132da4f7f7b46cc828bb85bd504bd0d392ce5f176748dac45c6db35", // 当前口令：shijianbu2026
    days: 7,
    key: "sj_auth_exp",
    title: "实践部经费助手",
    sub: "请输入访问口令",
  };

  var CSS = [
    "body.sj-locked { overflow: hidden; }",
    "body.sj-locked > *:not(#sj-login) { filter: blur(6px); pointer-events: none; user-select: none; }",
    "#sj-login { position: fixed; inset: 0; z-index: 99999; display: flex; align-items: center;",
    "  justify-content: center; background: rgba(250,250,252,.72); backdrop-filter: blur(2px);",
    "  font-family: -apple-system,'PingFang SC','Microsoft YaHei',sans-serif; }",
    "#sj-login .card { width: 340px; max-width: calc(100vw - 32px); background: #fff; border-radius: 16px;",
    "  box-shadow: 0 18px 50px rgba(60,20,90,.18); padding: 26px 24px 20px; border: 1px solid #eee; }",
    "#sj-login h2 { margin: 0 0 6px; font-size: 18px; color: #4a1f63; }",
    "#sj-login p.sub { margin: 0 0 16px; font-size: 13px; color: #888; }",
    "#sj-login input { width: 100%; padding: 11px 12px; border: 1px solid #ddd; border-radius: 10px;",
    "  font-size: 15px; outline: none; font-family: inherit; }",
    "#sj-login input:focus { border-color: #6f2d91; box-shadow: 0 0 0 3px rgba(111,45,145,.12); }",
    "#sj-login button { width: 100%; margin-top: 12px; padding: 11px; border: none; border-radius: 10px;",
    "  background: #6f2d91; color: #fff; font-size: 15px; cursor: pointer; font-family: inherit; }",
    "#sj-login button:disabled { background: #bba; cursor: wait; }",
    "#sj-login .err { color: #c0392b; font-size: 12.5px; margin-top: 10px; min-height: 17px; }",
    "#sj-login .foot { margin-top: 12px; font-size: 11.5px; color: #aaa; text-align: center; line-height: 1.6; }",
    "#sj-logout { color: #fff; text-decoration: underline; cursor: pointer; }",
  ].join("\n");

  function sha256(str) {
    var enc = new TextEncoder().encode(str);
    return crypto.subtle.digest("SHA-256", enc).then(function (buf) {
      return Array.prototype.map.call(new Uint8Array(buf), function (b) {
        return b.toString(16).padStart(2, "0");
      }).join("");
    });
  }

  function authed() {
    try {
      var exp = Number(sessionStorage.getItem(AUTH.key) || 0);
      return exp > Date.now();
    } catch (e) { return false; }
  }

  function markAuthed() {
    try {
      sessionStorage.setItem(AUTH.key, String(Date.now() + AUTH.days * 86400000));
    } catch (e) {}
  }

  function buildOverlay() {
    if (document.getElementById("sj-login")) return;
    var wrap = document.createElement("div");
    wrap.id = "sj-login";
    wrap.innerHTML = '' +
      '<div class="card">' +
      '  <h2>' + AUTH.title + '</h2>' +
      '  <p class="sub">' + AUTH.sub + '</p>' +
      '  <input id="sj-pw" type="password" autocomplete="current-password" placeholder="访问口令">' +
      '  <button id="sj-go">进入</button>' +
      '  <div class="err" id="sj-err"></div>' +
      '  <div class="foot">仅限实践部内部使用<br>问题联系：实践部</div>' +
      '</div>';
    document.body.appendChild(wrap);

    var pw = wrap.querySelector("#sj-pw");
    var btn = wrap.querySelector("#sj-go");
    var err = wrap.querySelector("#sj-err");

    function submit() {
      var v = pw.value;
      if (!v) { err.textContent = "请输入口令"; return; }
      btn.disabled = true;
      sha256(v).then(function (h) {
        if (h === AUTH.hash) {
          markAuthed();
          unlock();
        } else {
          err.textContent = "口令不对";
          pw.value = "";
          pw.focus();
        }
      }).catch(function () {
        err.textContent = "浏览器不支持安全校验（请用 https 访问）";
      }).then(function () { btn.disabled = false; });
    }

    btn.addEventListener("click", submit);
    pw.addEventListener("keydown", function (e) { if (e.key === "Enter") submit(); });
    setTimeout(function () { pw.focus(); }, 50);
  }

  function unlock() {
    document.body.classList.remove("sj-locked");
    var o = document.getElementById("sj-login");
    if (o) o.remove();
  }

  function addLogout() {
    var p = document.querySelector("header p");
    if (!p || document.getElementById("sj-logout")) return;
    var a = document.createElement("span");
    a.id = "sj-logout";
    a.textContent = "退出";
    a.style.marginLeft = "10px";
    a.addEventListener("click", function () {
      try { sessionStorage.removeItem(AUTH.key); } catch (e) {}
      location.reload();
    });
    p.appendChild(a);
  }

  function init() {
    var st = document.createElement("style");
    st.textContent = CSS;
    document.head.appendChild(st);

    if (authed()) { addLogout(); return; }
    document.body.classList.add("sj-locked");
    buildOverlay();
  }

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", init);
  else init();
})();
