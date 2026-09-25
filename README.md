# Deal Quill

Bare-bones **Microsoft Word** Office Add-in (task pane) for commercial real estate deal documents.  
Built for **Chad Nasir / RealEstateCA**.

**Demo — not legal advice.** Prompts drive AI behavior; there are no hardcoded contract playbooks or clause libraries.

## Live task pane URL

`https://raw.githack.com/Chadnasir/deal-quill-word/main/taskpane.html`

Manifest SourceLocation points at that URL (local override: `https://localhost:3000/taskpane.html`).

## What it does

1. **Task pane** beside the Word document (Office Add-in).
2. **Strike out** — apply strikethrough to selection (or whole body if none).
3. **Redline** — strike selection and insert a replacement after it.
4. **Give suggestions / Ask AI** — send your prompt + selection/doc context + labeled parties to a **local Grok bridge**; Apply insert or Apply redline.
5. **Parties panel** — on open (and on **Refresh parties**), heuristic detection of names/entities from document text; you label role (buyer, seller, landlord, tenant, lender, borrower, other) and display name.

## First-run UX

The task pane includes a three-step checklist (set the bridge token, select a clause, then strike or ask AI). On first run (no bridge token in Office storage / localStorage), the Grok bridge panel auto-opens and step 1 shows an inline token field so you do not hunt inside collapsed details. After the token is saved, the panel collapses and de-emphasizes again. The checklist disappears permanently after those steps are completed. Empty states guide the next action when there is no document text, the local bridge is offline, or no parties are detected.

Clarity is intentionally not recorded inside Word: Office task panes often block analytics and document text is sensitive. **Clarity primarily on OM portal; Deal Quill uses empty-state onboarding only.**

## Install in Word (sideload)

### Windows

1. Clone or download this repo; keep `manifest.xml` handy.
2. Start the local bridge (see below) so AI calls work.
3. Word → **Insert** → **Add-ins** → **Get Add-ins** → **Upload My Add-in** → upload `manifest.xml`.
4. Or use a **shared folder** catalog / sideloading network share that contains the manifest.
5. Open **Home** ribbon → **Deal Quill** to show the task pane.

### Mac

1. Same manifest upload path: **Insert** → **Add-ins** → **My Add-ins** → **Upload**, or place the manifest in the Word wef sideload folder for your Office version.
2. Trust the add-in when prompted.
3. Task pane loads from raw.githack (or localhost if you serve files locally).

### Local HTML host (optional)

```bash
cd deal-quill
npx --yes serve -l 3000
# then temporarily point manifest SourceLocation to https://localhost:3000/taskpane.html
```

## Local Grok MCP bridge (no API keys in the add-in)

The Word task pane calls **only** `http://127.0.0.1:8787` (configurable in the pane).  
Keys never live in `taskpane.js` / client storage as xAI tokens.

```bash
cd bridge
cp .env.example .env.local
# Edit .env.local:
#   DEAL_QUILL_BRIDGE_TOKEN=<openssl rand -base64 32>
#   Optional: XAI_API_KEY=...  or GROK_API_KEY=...  or GROK_MCP_URL=...
npm start
```

Paste the **same** `DEAL_QUILL_BRIDGE_TOKEN` into the task pane **Bridge token** field (stored in OfficeRuntime.storage / localStorage only).

### Endpoints

| Method | Path | Auth | Body → Response |
|--------|------|------|-----------------|
| GET | `/health` | none | `{ ok, status, ai, bind, authRequired }` — never returns keys |
| POST | `/v1/complete` | `X-Deal-Quill-Token` | `{ prompt, context, parties, mode }` → `{ text }` |
| POST | `/v1/mcp-style` | `X-Deal-Quill-Token` | same → MCP-shaped `{ jsonrpc, result: { content: [{ type, text }], text } }` |

If `XAI_API_KEY` or `GROK_API_KEY` is set, the bridge calls xAI chat completions.  
Else if `GROK_MCP_URL` is set, it proxies an MCP-style `tools/call`.  
Else it returns a **usable stub** suggestion so the UI works offline, with a clear “set XAI_API_KEY or start Grok MCP proxy” note.

**Grok Bot / MCP:** run `node bridge/server.js` (or `npm start`). The Word add-in never stores provider API keys.

### MCP-shaped request/response notes

Request (conceptual):

```json
{
  "jsonrpc": "2.0",
  "id": "deal-quill-1",
  "method": "tools/call",
  "params": {
    "name": "complete",
    "arguments": {
      "prompt": "Tighten indemnity for the tenant",
      "context": "…selection or doc…",
      "parties": [{ "role": "tenant", "personLabel": "Beta Tenant Inc." }],
      "mode": "suggest"
    }
  }
}
```

Response (from `/v1/mcp-style`):

```json
{
  "jsonrpc": "2.0",
  "id": null,
  "result": {
    "content": [{ "type": "text", "text": "…" }],
    "text": "…"
  }
}
```

## Party detection

`party-detect.js` runs heuristics on **plain text only**:

- `between A and B`
- `Party A: …` / labeled lines (`Buyer:`, `Seller:`, `Landlord:`, `Tenant:`, `Lender:`, …)
- Entity-like tokens (LLC, Inc., L.P., Corp., …)

Returns `{ id, rawName, roleGuess, personLabel }`.  
UI: editable role dropdown + name; persisted in `localStorage` / Office document settings.  
**No** CRE contract-type catalogs or clause libraries.

## SECURITY

Threat model: the Word task pane is **untrusted UI**; document text is sensitive; **only the localhost bridge** may hold provider keys.

Locked down in this repo:

1. **Loopback bind** — bridge listens on `127.0.0.1` only (never `0.0.0.0`).
2. **Shared local secret** — every `/v1/complete` and `/v1/mcp-style` request requires header `X-Deal-Quill-Token` matching `DEAL_QUILL_BRIDGE_TOKEN`. Missing/wrong → **401**.
3. **Origin / Referer allowlist** — `null` (Word desktop), `https://localhost:*`, `http(s)://127.0.0.1`, `https://*.officeapps.live.com`, `https://*.office.com`, `https://raw.githack.com`. All else denied (**403**).
4. **No `Access-Control-Allow-Origin: *`** — echoes only an allowlisted `Origin` (or `null`).
5. **Body size limit** — 512KB; prompt/context truncated; **never logs** document body or tokens (only lengths).
6. **Rate limit** — 30 requests / minute (in-memory).
7. **Keys** — `XAI_API_KEY` / `GROK_API_KEY` only in bridge process env (`.env.local` gitignored). Never returned to the client; never in taskpane JS source.
8. **Task pane** — AI output goes into a `<textarea>` via `.value` / textContent; no unsanitized `innerHTML`; no `eval`; only Microsoft Office.js CDN as remote script.

Generate a bridge token before first use:

```bash
python3 -c 'import secrets; print(secrets.token_urlsafe(32))'
# put into bridge/.env.local AND the task pane Bridge token field
```

## Layout

```
manifest.xml
taskpane.html
taskpane.css
taskpane.js          # thin loader (or use a+b)
taskpane-a.js
taskpane-b.js
party-detect.js
grok-client.js
bridge/server.js
bridge/package.json
bridge/.env.example
assets/icon.svg
README.md
.gitignore
```

## Disclaimer

Demo software for document editing UX. **Not legal advice.** Review every AI suggestion before applying.

## Provider

Chad Nasir / RealEstateCA · [GitHub repo](https://github.com/Chadnasir/deal-quill-word)
