/**
 * Deal Quill — local Grok MCP bridge (127.0.0.1 only).
 * Holds API keys server-side. Word add-in never stores xAI/Grok tokens.
 *
 * GET  /health
 * POST /v1/complete     { prompt, context, parties, mode } → { text }
 * POST /v1/mcp-style    same → MCP-shaped result
 *
 * Auth: X-Deal-Quill-Token must match DEAL_QUILL_BRIDGE_TOKEN.
 */
"use strict";

const http = require("http");
const { URL } = require("url");
const fs = require("fs");
const path = require("path");

const HOST = "127.0.0.1";
const PORT = Number(process.env.PORT || 8787);
const MAX_BODY = 512 * 1024;
const RATE_LIMIT = 30;
const RATE_WINDOW_MS = 60 * 1000;
const XAI_URL = process.env.XAI_API_URL || "https://api.x.ai/v1/chat/completions";
const XAI_MODEL = process.env.XAI_MODEL || "grok-2-latest";

function loadDotEnv() {
  [".env.local", ".env"].forEach(function (name) {
    const file = path.join(__dirname, name);
    try {
      if (!fs.existsSync(file)) return;
      fs.readFileSync(file, "utf8").split(/\r?\n/).forEach(function (line) {
        const t = line.trim();
        if (!t || t.startsWith("#")) return;
        const eq = t.indexOf("=");
        if (eq < 1) return;
        const k = t.slice(0, eq).trim();
        let v = t.slice(eq + 1).trim();
        if ((v.startsWith('"') && v.endsWith('"')) || (v.startsWith("'") && v.endsWith("'"))) {
          v = v.slice(1, -1);
        }
        if (!process.env[k]) process.env[k] = v;
      });
    } catch (e) { /* ignore */ }
  });
}
loadDotEnv();

const BRIDGE_TOKEN = process.env.DEAL_QUILL_BRIDGE_TOKEN || "";
const XAI_KEY = process.env.XAI_API_KEY || process.env.GROK_API_KEY || "";
const GROK_MCP_URL = process.env.GROK_MCP_URL || "";
const rateMap = new Map();

function isOriginAllowed(origin) {
  if (origin == null || origin === "" || origin === "null") return true;
  let u;
  try { u = new URL(origin); } catch (e) { return false; }
  const host = u.hostname.toLowerCase();
  if ((u.protocol === "https:" || u.protocol === "http:") && (host === "localhost" || host === "127.0.0.1")) return true;
  if (host === "raw.githack.com") return true;
  if (host.endsWith(".officeapps.live.com") || host === "officeapps.live.com") return true;
  if (host.endsWith(".office.com") || host === "office.com") return true;
  return false;
}

function corsHeaders(req) {
  const origin = req.headers.origin;
  const h = {
    "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type, X-Deal-Quill-Token",
    "Access-Control-Max-Age": "600",
    Vary: "Origin"
  };
  if (origin && isOriginAllowed(origin)) h["Access-Control-Allow-Origin"] = origin;
  else if (!origin) h["Access-Control-Allow-Origin"] = "null";
  return h;
}

function send(res, status, body, extra) {
  const payload = typeof body === "string" ? body : JSON.stringify(body);
  res.writeHead(status, Object.assign({
    "Content-Type": "application/json; charset=utf-8",
    "Content-Length": Buffer.byteLength(payload)
  }, extra || {}));
  res.end(payload);
}

function rateLimit() {
  const now = Date.now();
  let bucket = rateMap.get("local");
  if (!bucket || now - bucket.start >= RATE_WINDOW_MS) {
    bucket = { start: now, count: 0 };
    rateMap.set("local", bucket);
  }
  bucket.count += 1;
  return bucket.count <= RATE_LIMIT;
}

function checkOrigin(req) {
  const origin = req.headers.origin;
  if (origin != null && origin !== "" && !isOriginAllowed(origin)) return false;
  const referer = req.headers.referer || req.headers.referrer;
  if (referer) {
    try {
      if (!isOriginAllowed(new URL(referer).origin)) return false;
    } catch (e) { return false; }
  }
  return true;
}

function requireToken(req) {
  if (!BRIDGE_TOKEN) {
    return { ok: false, reason: "Bridge misconfigured: set DEAL_QUILL_BRIDGE_TOKEN in bridge/.env.local" };
  }
  const got = req.headers["x-deal-quill-token"];
  if (!got || got !== BRIDGE_TOKEN) return { ok: false, reason: "Unauthorized" };
  return { ok: true };
}

function readBody(req) {
  return new Promise(function (resolve, reject) {
    let size = 0;
    const chunks = [];
    req.on("data", function (chunk) {
      size += chunk.length;
      if (size > MAX_BODY) {
        reject(Object.assign(new Error("Body too large"), { status: 413 }));
        req.destroy();
        return;
      }
      chunks.push(chunk);
    });
    req.on("end", function () { resolve(Buffer.concat(chunks).toString("utf8")); });
    req.on("error", reject);
  });
}

function sanitizeInput(obj) {
  const prompt = String((obj && obj.prompt) || "").slice(0, 8000);
  const context = String((obj && obj.context) || "").slice(0, 48000);
  const mode = String((obj && obj.mode) || "suggest").slice(0, 32);
  let parties = [];
  if (obj && Array.isArray(obj.parties)) {
    parties = obj.parties.slice(0, 40).map(function (p) {
      return {
        id: String((p && p.id) || "").slice(0, 64),
        rawName: String((p && p.rawName) || "").slice(0, 120),
        role: String((p && (p.role || p.roleGuess)) || "").slice(0, 32),
        personLabel: String((p && p.personLabel) || "").slice(0, 120)
      };
    });
  }
  return { prompt, context, mode, parties };
}

function buildUserMessage(s) {
  const partyLines = (s.parties || []).map(function (p) {
    return "- " + (p.role || "other") + ": " + (p.personLabel || p.rawName || "");
  }).join("\n");
  return [
    "Mode: " + s.mode,
    partyLines ? "Parties:\n" + partyLines : "Parties: (none labeled)",
    "",
    "User prompt:",
    s.prompt || "(none)",
    "",
    "Document / selection context:",
    s.context || "(empty)"
  ].join("\n");
}

function stubSuggestion(s) {
  const snippet = (s.context || "").replace(/\s+/g, " ").trim().slice(0, 180);
  return [
    "[bridge online — set XAI_API_KEY or GROK_API_KEY, or GROK_MCP_URL, then restart]",
    "",
    "Demo suggestion for mode=" + s.mode + ":",
    s.prompt
      ? "Regarding your prompt (“" + s.prompt.slice(0, 120) + "”): consider clarifying undefined terms and aligning defined party names."
      : "Provide a clearer replacement that matches the labeled parties.",
    snippet ? "Context excerpt: “" + snippet + (snippet.length >= 180 ? "…" : "") + "”" : ""
  ].filter(Boolean).join("\n");
}

async function completeText(s) {
  if (XAI_KEY) {
    const res = await fetch(XAI_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json", Authorization: "Bearer " + XAI_KEY },
      body: JSON.stringify({
        model: XAI_MODEL,
        messages: [
          { role: "system", content: "You assist with editing commercial deal documents. Follow the user prompt. Do not invent legal advice. Return plain text suitable to insert or redline in Word." },
          { role: "user", content: buildUserMessage(s) }
        ],
        temperature: 0.3
      })
    });
    const data = await res.json().catch(function () { return null; });
    if (!res.ok) throw new Error((data && data.error && data.error.message) || ("xAI HTTP " + res.status));
    return String((data && data.choices && data.choices[0] && data.choices[0].message && data.choices[0].message.content) || "");
  }
  if (GROK_MCP_URL) {
    const res = await fetch(GROK_MCP_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        jsonrpc: "2.0",
        id: "deal-quill-" + Date.now(),
        method: "tools/call",
        params: { name: "complete", arguments: { prompt: s.prompt, context: s.context, parties: s.parties, mode: s.mode } }
      })
    });
    const data = await res.json().catch(function () { return null; });
    if (!res.ok) throw new Error("GROK_MCP_URL HTTP " + res.status);
    if (data && data.result) {
      if (typeof data.result.text === "string") return data.result.text;
      if (Array.isArray(data.result.content)) {
        return data.result.content.map(function (c) { return c && c.text ? c.text : ""; }).filter(Boolean).join("\n");
      }
    }
    if (data && data.text) return String(data.text);
    return JSON.stringify(data);
  }
  return stubSuggestion(s);
}

const server = http.createServer(async function (req, res) {
  const headers = corsHeaders(req);
  if (req.method === "OPTIONS") { res.writeHead(204, headers); res.end(); return; }
  if (!checkOrigin(req)) { send(res, 403, { error: "Origin not allowed" }, headers); return; }
  if (!rateLimit()) { send(res, 429, { error: "Rate limit exceeded" }, headers); return; }

  let pathname = "/";
  try { pathname = new URL(req.url || "/", "http://127.0.0.1").pathname; } catch (e) {
    send(res, 400, { error: "Bad request" }, headers); return;
  }

  if (req.method === "GET" && pathname === "/health") {
    send(res, 200, {
      ok: true, status: "ok", bind: HOST + ":" + PORT, authRequired: true,
      ai: XAI_KEY ? "xai" : GROK_MCP_URL ? "grok-mcp" : "stub",
      message: BRIDGE_TOKEN ? "Bridge online — send X-Deal-Quill-Token on /v1/complete" : "Set DEAL_QUILL_BRIDGE_TOKEN before use"
    }, headers);
    return;
  }

  if (req.method === "POST" && (pathname === "/v1/complete" || pathname === "/v1/mcp-style")) {
    const auth = requireToken(req);
    if (!auth.ok) { send(res, 401, { error: auth.reason }, headers); return; }
    let raw;
    try { raw = await readBody(req); } catch (e) {
      send(res, e.status || 400, { error: e.message || "Bad body" }, headers); return;
    }
    let parsed = {};
    try { parsed = raw ? JSON.parse(raw) : {}; } catch (e) {
      send(res, 400, { error: "Invalid JSON" }, headers); return;
    }
    const s = sanitizeInput(parsed);
    console.log("[deal-quill-bridge] complete mode=%s promptLen=%d contextLen=%d parties=%d", s.mode, s.prompt.length, s.context.length, s.parties.length);
    try {
      const text = await completeText(s);
      if (pathname === "/v1/mcp-style") {
        send(res, 200, { jsonrpc: "2.0", id: parsed.id || null, result: { content: [{ type: "text", text: text }], text: text } }, headers);
      } else {
        send(res, 200, { text: text }, headers);
      }
    } catch (e) {
      console.log("[deal-quill-bridge] error: %s", e && e.message ? e.message : "unknown");
      send(res, 502, { error: e && e.message ? e.message : "Upstream error" }, headers);
    }
    return;
  }

  send(res, 404, { error: "Not found" }, headers);
});

if (!BRIDGE_TOKEN) {
  console.warn("[deal-quill-bridge] WARNING: DEAL_QUILL_BRIDGE_TOKEN is not set. /v1/complete will return 401 until you set bridge/.env.local");
}

server.listen(PORT, HOST, function () {
  console.log("[deal-quill-bridge] listening on http://%s:%s (loopback only). AI=%s", HOST, PORT, XAI_KEY ? "xai" : GROK_MCP_URL ? "grok-mcp" : "stub");
  console.log("[deal-quill-bridge] Word add-in never stores API keys. Start from Grok Bot or: npm start");
});
