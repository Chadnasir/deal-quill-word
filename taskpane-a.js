/**
 * Deal Quill — Word task pane (Office.js).
 * Selection / whole-doc strike, redline, AI suggest+apply.
 * Escape all AI text before DOM; never eval; never store API keys.
 */
"use strict";
var parties = [];
  var DOC_KEY_PREFIX = "dealQuill.parties.";

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
    return el ? el.value.trim() : "";
  }

  function setStatus(ok, title) {
    var dot = document.getElementById("bridge-status");
    if (!dot) return;
    dot.classList.remove("ok", "err");
    if (ok === true) dot.classList.add("ok");
    else if (ok === false) dot.classList.add("err");
    if (title) dot.title = title;
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
