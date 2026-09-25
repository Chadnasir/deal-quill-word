/**
 * Deal Quill — Word task pane (Office.js).
 * Selection / whole-doc strike, redline, AI suggest+apply.
 * Escape all AI text before DOM; never eval; never store API keys.
 */
"use strict";
var parties = [];
  var DOC_KEY_PREFIX = "dealQuill.parties.";
  var CHECKLIST_KEY = "dealQuill.firstRun";

  function readChecklist() {
    try {
      var raw = localStorage.getItem(CHECKLIST_KEY);
      var parsed = raw ? JSON.parse(raw) : {};
      return parsed && typeof parsed === "object" ? parsed : {};
    } catch (e) { return {}; }
  }

  function saveChecklist(state) {
    var payload = JSON.stringify(state);
    try { localStorage.setItem(CHECKLIST_KEY, payload); } catch (e) { /* ignore */ }
    try {
      if (typeof OfficeRuntime !== "undefined" && OfficeRuntime.storage) {
        OfficeRuntime.storage.setItem(CHECKLIST_KEY, payload);
      }
    } catch (e2) { /* ignore */ }
  }

  function renderChecklist() {
    var panel = document.getElementById("first-run-checklist");
    if (!panel) return;
    var state = readChecklist();
    var steps = ["bridge", "selection", "action"];
    var complete = steps.every(function (step) { return !!state[step]; });
    if (complete) state.dismissed = true;
    if (state.dismissed) {
      panel.hidden = true;
      if (complete) saveChecklist(state);
      return;
    }
    panel.hidden = false;
    steps.forEach(function (step) {
      var item = panel.querySelector('[data-checklist-step="' + step + '"]');
      if (!item) return;
      var done = !!state[step];
      item.classList.toggle("is-complete", done);
      item.setAttribute("aria-label", done ? "Complete: " + step : "Not complete: " + step);
      var box = item.querySelector(".checklist-box");
      if (box) box.textContent = done ? "✓" : String(steps.indexOf(step) + 1);
    });
    var bridgeDone = !!state.bridge;
    var inlineWrap = document.getElementById("checklist-inline-token");
    if (inlineWrap) inlineWrap.hidden = bridgeDone || !!(window.DealQuillGrok && DealQuillGrok.getToken()) || !!getBridgeTokenInput();
  }


  function dismissChecklist() {
    var state = readChecklist();
    state.dismissed = true;
    saveChecklist(state);
    renderChecklist();
  }

  function markChecklistStep(step) {
    var state = readChecklist();
    if (state[step]) return;
    state[step] = true;
    saveChecklist(state);
    renderChecklist();
  }

  function updateDocumentEmptyState(text) {
    var el = document.getElementById("document-empty-state");
    if (el) el.hidden = !!String(text || "").trim();
  }

  function escapeHtml(s) {
    return String(s == null ? "" : s)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  function setTextContent(el, text) {
    if (!el) return;
    el.textContent = text == null ? "" : String(text);
  }

  function docStorageKey() {
    return DOC_KEY_PREFIX + "default";
  }

  function loadPartiesFromStorage() {
    try {
      var raw = localStorage.getItem(docStorageKey());
      if (!raw) return null;
      var parsed = JSON.parse(raw);
      return Array.isArray(parsed) ? parsed : null;
    } catch (e) {
      return null;
    }
  }

  function savePartiesToStorage() {
    try {
      localStorage.setItem(docStorageKey(), JSON.stringify(parties));
    } catch (e) { /* ignore */ }
    try {
      if (typeof Office !== "undefined" && Office.context && Office.context.document && Office.context.document.settings) {
        Office.context.document.settings.set("dealQuill.parties", JSON.stringify(parties));
        Office.context.document.settings.saveAsync(function () { /* ignore */ });
      }
    } catch (e2) { /* ignore */ }
  }

  function getBridgeTokenInput() {
    var el = document.getElementById("bridge-token");
    var inline = document.getElementById("bridge-token-inline");
    var main = el ? el.value.trim() : "";
    var side = inline ? inline.value.trim() : "";
    return main || side;
  }

  function syncBridgeTokenFields(value) {
    var v = value == null ? "" : String(value);
    var el = document.getElementById("bridge-token");
    var inline = document.getElementById("bridge-token-inline");
    if (el && el.value !== v) el.value = v;
    if (inline && inline.value !== v) inline.value = v;
  }

  /** First-run: open bridge panel + show checklist inline token when no token yet. */
  function syncFirstRunBridgeUi(hasToken) {
    var tokenPresent = !!hasToken;
    if (!tokenPresent) {
      tokenPresent = !!(getBridgeTokenInput() || (window.DealQuillGrok && DealQuillGrok.getToken()));
    }
    var inlineWrap = document.getElementById("checklist-inline-token");
    if (inlineWrap) inlineWrap.hidden = tokenPresent;
    var panel = document.getElementById("bridge-panel");
    if (!panel) return;
    if (!tokenPresent) {
      panel.open = true;
      panel.classList.add("is-first-run");
    } else {
      panel.classList.remove("is-first-run");
      panel.open = false;
    }
  }

  function setStatus(ok, title) {
    var pill = document.getElementById("bridge-status");
    if (!pill) return;
    pill.classList.remove("ok", "err");
    var label = document.getElementById("bridge-status-label") || pill.querySelector(".status-pill-label");
    if (ok === true) {
      pill.classList.add("ok");
      if (label) label.textContent = "Connected";
    } else if (ok === false) {
      pill.classList.add("err");
      if (label) label.textContent = "Offline";
    } else if (label) {
      label.textContent = "Bridge";
    }
    if (title) pill.title = title;
    var empty = document.getElementById("bridge-empty-state");
    if (empty) empty.hidden = ok !== false;
    var bridgePanel = document.getElementById("bridge-panel");
    if (bridgePanel && ok === false) bridgePanel.open = true;
  }

  function showToast(message, kind) {
    var msg = message == null ? "" : String(message);
    var toast = document.getElementById("dq-toast");
    if (toast) {
      toast.textContent = msg;
      toast.classList.remove("is-success", "is-visible");
      if (kind === "success") toast.classList.add("is-success");
      toast.hidden = false;
      void toast.offsetWidth;
      toast.classList.add("is-visible");
      clearTimeout(showToast._t);
      showToast._t = setTimeout(function () {
        toast.classList.remove("is-visible");
        setTimeout(function () { toast.hidden = true; }, 160);
      }, 2200);
      return;
    }
    var flash = document.getElementById("apply-flash");
    if (!flash) return;
    flash.textContent = msg || "Applied";
    flash.hidden = false;
    clearTimeout(showToast._t);
    showToast._t = setTimeout(function () { flash.hidden = true; }, 1800);
  }

  function flashOk(btn) {
    if (!btn) return;
    btn.classList.remove("flash-ok");
    void btn.offsetWidth;
    btn.classList.add("flash-ok");
    setTimeout(function () { btn.classList.remove("flash-ok"); }, 600);
  }

  function setAiLoading(on, which) {
    var ids = which === "ask" ? ["btn-ask"] : which === "suggest" ? ["btn-suggest"] : ["btn-suggest", "btn-ask"];
    ids.forEach(function (id) {
      var b = document.getElementById(id);
      if (b) {
        if (on) b.classList.add("is-loading");
        else b.classList.remove("is-loading");
        b.disabled = !!on;
        b.setAttribute("aria-busy", on ? "true" : "false");
      }
    });
    var load = document.getElementById("ai-loading");
    if (load) load.hidden = !on;
    var panel = document.getElementById("ai-panel");
    if (panel) panel.classList.toggle("is-ai-loading", !!on);
  }

  function readDocumentText(callback) {
    Word.run(function (context) {
      var body = context.document.body;
      context.load(body, "text");
      return context.sync().then(function () {
        callback(null, body.text || "");
      });
    }).catch(function (err) {
      callback(err, "");
    });
  }

  function getSelectionOrBodyText(callback) {
    Word.run(function (context) {
      var sel = context.document.getSelection();
      context.load(sel, "text");
      return context.sync().then(function () {
        var t = (sel.text || "").trim();
        if (t) {
          markChecklistStep("selection");
          callback(null, t, true);
          return null;
        }
        var body = context.document.body;
        context.load(body, "text");
        return context.sync().then(function () {
          callback(null, body.text || "", false);
        });
      });
    }).catch(function (err) {
      callback(err, "", false);
    });
  }

  function applyStrike() {
    Word.run(function (context) {
      var sel = context.document.getSelection();
      context.load(sel, "text");
      return context.sync().then(function () {
        var range = sel;
        if (!(sel.text || "").trim()) {
          range = context.document.body;
        }
        range.font.strikeThrough = true;
        return context.sync();
      });
    }).catch(function (err) {
      console.warn("Strike failed", err);
      alert("Strike failed: " + (err && err.message ? err.message : err));
    });
  }

  function applyRedline(replacement) {
    var text = String(replacement || "").trim();
    if (!text) {
      alert("Enter replacement text for redline.");
      return;
    }
    Word.run(function (context) {
      var sel = context.document.getSelection();
      context.load(sel, "text");
      return context.sync().then(function () {
        if (!(sel.text || "").trim()) {
          throw new Error("Select text to redline, or use Strike on the whole document.");
        }
        sel.font.strikeThrough = true;
        sel.insertText(" " + text, Word.InsertLocation.after);
        return context.sync();
      });
    }).catch(function (err) {
      console.warn("Redline failed", err);
      alert("Redline failed: " + (err && err.message ? err.message : err));
    });
  }

  function applyInsert(text) {
    var t = String(text || "");
    if (!t) return;
    Word.run(function (context) {
      var sel = context.document.getSelection();
      context.load(sel, "text");
      return context.sync().then(function () {
        if ((sel.text || "").trim()) {
          sel.insertText(t, Word.InsertLocation.replace);
        } else {
          context.document.body.insertText("\n" + t, Word.InsertLocation.end);
        }
        return context.sync();
      });
    }).catch(function (err) {
      alert("Apply failed: " + (err && err.message ? err.message : err));
    });
  }

  function renderParties() {
    var list = document.getElementById("parties-list");
    if (!list) return;
    while (list.firstChild) list.removeChild(list.firstChild);

    if (!parties.length) {
      var empty = document.createElement("div");
      empty.className = "dq-empty-state compact";
      empty.setAttribute("role", "status");
      var emptyTitle = document.createElement("strong");
      emptyTitle.textContent = "No parties detected yet";
      var emptyCopy = document.createElement("p");
      emptyCopy.textContent = "Open a contract and click Refresh, or add a party manually.";
      empty.appendChild(emptyTitle);
      empty.appendChild(emptyCopy);
      list.appendChild(empty);
    }

    var roles = (window.DealQuillPartyDetect && DealQuillPartyDetect.ROLE_LABELS) ||
      ["buyer", "seller", "landlord", "tenant", "lender", "borrower", "other"];

    parties.forEach(function (p, idx) {
      var card = document.createElement("div");
      card.className = "party-card";
      card.setAttribute("data-id", p.id);

      var raw = document.createElement("div");
      raw.className = "raw";
      raw.textContent = "Detected: " + (p.rawName || "");
      card.appendChild(raw);

      var row = document.createElement("div");
      row.className = "party-row";

      var roleSel = document.createElement("select");
      roleSel.setAttribute("aria-label", "Role");
      roles.forEach(function (r) {
        var opt = document.createElement("option");
        opt.value = r;
        opt.textContent = r.charAt(0).toUpperCase() + r.slice(1);
        if (r === (p.roleGuess || "other")) opt.selected = true;
        roleSel.appendChild(opt);
      });
      roleSel.addEventListener("change", function () {
        parties[idx].roleGuess = roleSel.value;
        savePartiesToStorage();
      });

      var nameInput = document.createElement("input");
      nameInput.type = "text";
      nameInput.value = p.personLabel || p.rawName || "";
      nameInput.setAttribute("aria-label", "Display name");
      nameInput.addEventListener("change", function () {
        parties[idx].personLabel = nameInput.value.trim();
        savePartiesToStorage();
      });

      row.appendChild(roleSel);
      row.appendChild(nameInput);
      card.appendChild(row);

      var actions = document.createElement("div");
      actions.className = "party-actions";
      var rm = document.createElement("button");
      rm.type = "button";
      rm.className = "btn btn-ghost btn-sm";
      rm.textContent = "Remove";
      rm.addEventListener("click", function () {
        parties.splice(idx, 1);
        savePartiesToStorage();
        renderParties();
      });
      actions.appendChild(rm);
      card.appendChild(actions);

      list.appendChild(card);
    });

    var hint = document.getElementById("parties-hint");
    if (hint) {
      setTextContent(
        hint,
        parties.length
          ? parties.length + " part" + (parties.length === 1 ? "y" : "ies") + " — label roles and names."
          : "Detected from document text. Label roles and names, then save."
      );
    }
  }
