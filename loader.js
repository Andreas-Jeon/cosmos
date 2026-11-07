/* Cosmos Mobile loader: OneDrive login (MSAL) -> data.js from OneDrive -> start the galaxy screen */
(function () {
  "use strict";
  var CFG = {
    clientId: "40855125-e565-46a2-b87c-0b74733a1276",
    redirectUri: "https://andreas-jeon.github.io/cosmos/",
    dataPath: "My workspace/Cosmos/_app/data.js",
    scopes: ["Files.Read"]
  };
  var LOCAL = /^(localhost|127\.0\.0\.1)$/.test(location.hostname);
  if (location.hostname !== "andreas-jeon.github.io") CFG.redirectUri = location.origin + location.pathname.replace(/index\.html$/, "");

  var panel = document.getElementById("panel");
  var esc = function (s) { return String(s).replace(/[&<>"']/g, function (c) { return {"&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"}[c]; }); };
  function show(title, body, btn, fn) {
    panel.innerHTML = "<h2>" + esc(title) + "</h2><p class=\"empty\">" + body + "</p>" + (btn ? "<button class=\"more\" id=\"ldbtn\">" + esc(btn) + "</button>" : "");
    if (btn) document.getElementById("ldbtn").onclick = fn;
  }
  function banner(text, fn) {
    var b = document.getElementById("ldbanner");
    if (!b) { b = document.createElement("button"); b.id = "ldbanner"; b.className = "ldbanner"; document.body.appendChild(b); }
    b.textContent = text; b.onclick = function () { b.remove(); if (fn) fn(); };
  }

  /* IndexedDB: keeps the last index so the app opens at once and also offline */
  function idb() {
    return new Promise(function (ok, no) {
      var r = indexedDB.open("cosmos", 1);
      r.onupgradeneeded = function () { r.result.createObjectStore("kv"); };
      r.onsuccess = function () { ok(r.result); }; r.onerror = function () { no(r.error); };
    });
  }
  function kv(mode, key, val) {
    return idb().then(function (db) {
      return new Promise(function (ok, no) {
        var st = db.transaction("kv", mode).objectStore("kv");
        var q = mode === "readonly" ? st.get(key) : st.put(val, key);
        q.onsuccess = function () { ok(q.result); }; q.onerror = function () { no(q.error); };
      });
    }).catch(function () { return null; });
  }
  var get = function (k) { return kv("readonly", k); }, put = function (k, v) { return kv("readwrite", k, v); };

  var parse = function (t) { return JSON.parse(t.slice(t.indexOf("{"), t.lastIndexOf("}") + 1)); };
  var started = false;
  function start(text) {
    if (started) return; started = true;
    try { window.COSMOS_DATA = parse(text); } catch (e) { started = false; show("색인을 읽지 못함", esc(e.message)); return; }
    var s = document.createElement("script");
    s.textContent = document.getElementById("cosmos-main").textContent;
    document.body.appendChild(s);
  }

  /* login */
  var pca = window.msal ? new msal.PublicClientApplication({
    auth: {clientId: CFG.clientId, authority: "https://login.microsoftonline.com/consumers", redirectUri: CFG.redirectUri},
    cache: {cacheLocation: "localStorage"}
  }) : null;
  function account() { return pca && (pca.getActiveAccount() || pca.getAllAccounts()[0]) || null; }
  function token(interactive) {
    var acc = account();
    if (!acc) { if (interactive) pca.loginRedirect({scopes: CFG.scopes, prompt: "select_account"}); return Promise.resolve(null); }
    pca.setActiveAccount(acc);
    return pca.acquireTokenSilent({scopes: CFG.scopes, account: acc}).then(function (r) { return r.accessToken; }, function () {
      if (interactive) pca.acquireTokenRedirect({scopes: CFG.scopes, account: acc});
      return null;
    });
  }
  var enc = function (p) { return p.split("/").map(encodeURIComponent).join("/"); };
  function graph(path, tk) {
    return fetch("https://graph.microsoft.com/v1.0" + path, {headers: {Authorization: "Bearer " + tk}}).then(function (r) {
      if (r.ok) return r.json();
      return r.text().then(function (t) { var e = new Error("OneDrive " + r.status); e.status = r.status; e.detail = t.slice(0, 300); throw e; });
    });
  }

  /* data.js from OneDrive; downloads only when it changed */
  function refresh(tk) {
    return graph("/me/drive/root:/" + enc(CFG.dataPath), tk).then(function (meta) {
      return get("etag").then(function (old) {
        if (old && old === meta.eTag) return null;
        var url = meta["@microsoft.graph.downloadUrl"];
        var viaGraph = function () { return fetch("https://graph.microsoft.com/v1.0/me/drive/items/" + meta.id + "/content", {headers: {Authorization: "Bearer " + tk}}).then(function (r) { if (!r.ok) throw new Error("download " + r.status); return r.text(); }); };
        var p = url ? fetch(url).then(function (r) { if (!r.ok) throw new Error("download " + r.status); return r.text(); }).catch(viaGraph) : viaGraph();
        return p.then(function (text) { return put("text", text).then(function () { return put("etag", meta.eTag); }).then(function () { return text; }); });
      });
    });
  }

  /* open a TMI file (C:\...\OneDrive\...) in OneDrive */
  window.cosmosOpen = function (full, btn) {
    var m = String(full).replace(/\\/g, "/").match(/\/OneDrive[^/]*\/(.+)$/i);
    if (!m) { alert("OneDrive 폴더 밖의 파일이라 폰에서 열 수 없음"); return; }
    var w = window.open("about:blank", "_blank");
    var label = btn && btn.firstChild; if (label) label.textContent = "여는 중…";
    token(true).then(function (tk) {
      if (!tk) { if (w) w.close(); return; }
      return graph("/me/drive/root:/" + enc(m[1]) + "?select=webUrl", tk).then(function (it) {
        if (w) w.location.href = it.webUrl; else location.href = it.webUrl;
        if (label) label.textContent = "OneDrive에서 열기";
      });
    }).catch(function (e) {
      if (w) w.close();
      if (label) label.textContent = "OneDrive에서 열기";
      alert(e.status === 404 ? "OneDrive에서 파일을 찾지 못함. 아직 동기화 전이거나 이름·위치가 바뀌었을 수 있음." : "열지 못함: " + e.message);
    });
  };
  window.cosmosRefresh = function () {
    token(true).then(function (tk) { if (!tk) return; return put("etag", "").then(function () { return refresh(tk); }).then(function () { location.reload(); }); })
      .catch(function (e) { banner("갱신 실패: " + e.message); });
  };
  window.cosmosLogout = function () {
    if (!confirm("로그아웃하고 이 폰에 저장된 색인을 지울까요?")) return;
    put("text", "").then(function () { return put("etag", ""); }).then(function () { pca.logoutRedirect({account: account()}); });
  };

  function main() {
    if ("serviceWorker" in navigator && !LOCAL) navigator.serviceWorker.register("sw.js").catch(function () {});
    if (LOCAL && /[?&]local\b/.test(location.search)) { fetch("data.js").then(function (r) { return r.text(); }).then(start); return; }
    if (!pca) { show("로그인 모듈을 불러오지 못함", "인터넷 연결을 확인한 뒤 다시 열어 주세요."); return; }
    show("Cosmos", "불러오는 중…");
    pca.initialize().then(function () { return pca.handleRedirectPromise(); }).then(function (res) {
      if (res && res.account) pca.setActiveAccount(res.account);
      return get("text");
    }, function (e) { show("로그인 오류", esc(e.message), "다시 로그인", function () { token(true); }); throw e; }).then(function (cached) {
      if (cached) start(cached);
      if (!account()) {
        if (!cached) show("Cosmos", "내 OneDrive의 Cosmos 색인을 읽어 은하를 그림.<br>처음 한 번 Microsoft 계정으로 로그인하고, 파일 읽기 권한(읽기 전용)을 허용하면 됨.", "Microsoft 계정으로 로그인", function () { token(true); });
        return;
      }
      if (!cached) show("Cosmos", "OneDrive에서 색인을 받는 중… (처음 한 번 1 MB 정도)");
      return token(false).then(function (tk) {
        if (!tk) { if (cached) banner("로그인이 만료됨 · 눌러서 다시 로그인", function () { token(true); }); else show("Cosmos", "다시 로그인이 필요함.", "로그인", function () { token(true); }); return; }
        return refresh(tk).then(function (text) {
          if (!text) return;
          if (started) banner("새 색인이 있음 · 눌러서 다시 열기", function () { location.reload(); }); else start(text);
        });
      }).catch(function (e) {
        if (started) banner("최신 색인 확인 실패 (저장된 색인으로 표시 중)");
        else show("색인을 받지 못함", esc(e.message) + (e.status === 404 ? "<br>OneDrive에서 <code>" + esc(CFG.dataPath) + "</code> 를 찾지 못함." : "") + (e.detail ? "<br><small>" + esc(e.detail) + "</small>" : ""), "다시 시도", function () { location.reload(); });
      });
    }).catch(function () {});
  }
  main();
})();
