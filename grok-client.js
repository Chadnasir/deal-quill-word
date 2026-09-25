/**
 * Deal Quill — client for local Grok MCP bridge.
 * No API keys in the add-in. Auth via shared local token header only.
 */
(function (global) {
  "use strict";

  var DEFAULT_BRIDGE = "http://127.0.0.1:8787";
  var TOKEN_KEY = "dealQuill.bridgeToken";
  var URL_KEY = "dealQuill.bridgeUrl";

  function getBridgeUrl() {
    try {
      var el = document.getElementById("bridge-url");
      if (el && el.value) return el.value.replace(/\/$/, "");
    } catch (e) { /* ignore */ }
    try {
      var stored = localStorage.getItem(URL_KEY);
      if (stored) return stored.replace(/\/$/, "");
    } catch (e2) { /* ignore */ }
    return DEFAULT_BRIDGE;
  }

  function setBridgeUrl(url) {
    try { localStorage.setItem(URL_KEY, url); } catch (e) { /* ignore */ }
  }

  function getToken() {
    try {
      return localStorage.getItem(TOKEN_KEY) || "";
    } catch (e) {
      return "";
    }
  }

  function setToken(token) {
    try {
      if (token) localStorage.setItem(TOKEN_KEY, token);
      else localStorage.removeItem(TOKEN_KEY);
    } catch (e) { /* ignore */ }
  }

  /** Prefer OfficeRuntime.storage when available; fall back to localStorage. */
  function loadTokenAsync(cb) {
    try {
      if (typeof OfficeRuntime !== "undefined" && OfficeRuntime.storage) {
        OfficeRuntime.storage.getItem(TOKEN_KEY).then(function (v) {
          if (v) {
            try { localStorage.setItem(TOKEN_KEY, v); } catch (e) { /* ignore */ }
            cb(v);
          } else {
            cb(getToken());
          }
        }).catch(function () { cb(getToken()); });
        return;
      }
    } catch (e) { /* ignore */ }
    cb(getToken());
  }

  function saveTokenAsync(token, cb) {
    setToken(token);
    try {
      if (typeof OfficeRuntime !== "undefined" && OfficeRuntime.storage) {
        var p = token
          ? OfficeRuntime.storage.setItem(TOKEN_KEY, token)
          : OfficeRuntime.storage.removeItem(TOKEN_KEY);
        p.then(function () { if (cb) cb(); }).catch(function () { if (cb) cb(); });
        return;
      }
    } catch (e) { /* ignore */ }
    if (cb) cb();
  }

  function complete(opts) {
    var url = (opts.bridgeUrl || getBridgeUrl()) + "/v1/complete";
    var token = opts.token != null ? opts.token : getToken();
    var body = {
      prompt: String(opts.prompt || "").slice(0, 8000),
      context: String(opts.context || "").slice(0, 48000),
      parties: opts.parties || [],
      mode: opts.mode || "suggest"
    };

    return fetch(url, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "X-Deal-Quill-Token": token
      },
      body: JSON.stringify(body)
    }).then(function (res) {
      return res.text().then(function (raw) {
        var data = null;
        try { data = raw ? JSON.parse(raw) : null; } catch (e) { data = { text: raw }; }
        if (!res.ok) {
          var msg = (data && (data.error || data.message || data.text)) || ("HTTP " + res.status);
          if (res.status === 401) {
            msg = "Bridge rejected token (401). Set the same DEAL_QUILL_BRIDGE_TOKEN in the task pane and bridge/.env.local.";
          }
          var err = new Error(msg);
          err.status = res.status;
          err.data = data;
          throw err;
        }
        return data || { text: "" };
      });
    });
  }

  function health(bridgeUrl, token) {
    var base = (bridgeUrl || getBridgeUrl()).replace(/\/$/, "");
    return fetch(base + "/health", {
      method: "GET",
      headers: token ? { "X-Deal-Quill-Token": token } : {}
    }).then(function (res) {
      return res.json().catch(function () { return { ok: res.ok }; });
    });
  }

  global.DealQuillGrok = {
    DEFAULT_BRIDGE: DEFAULT_BRIDGE,
    getBridgeUrl: getBridgeUrl,
    setBridgeUrl: setBridgeUrl,
    getToken: getToken,
    setToken: setToken,
    loadTokenAsync: loadTokenAsync,
    saveTokenAsync: saveTokenAsync,
    complete: complete,
    health: health
  };
})(typeof window !== "undefined" ? window : this);
