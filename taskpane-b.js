
  function refreshParties() {
    readDocumentText(function (err, text) {
      if (err) {
        console.warn(err);
        alert("Could not read document: " + (err.message || err));
        return;
      }
      var detected = window.DealQuillPartyDetect
        ? DealQuillPartyDetect.detectParties(text)
        : [];
      var byRaw = {};
      parties.forEach(function (p) {
        byRaw[(p.rawName || "").toLowerCase()] = p;
      });
      var merged = detected.map(function (d) {
        var prev = byRaw[(d.rawName || "").toLowerCase()];
        if (prev) {
          return {
            id: prev.id || d.id,
            rawName: d.rawName,
            roleGuess: prev.roleGuess && prev.roleGuess !== "other" ? prev.roleGuess : d.roleGuess,
            personLabel: prev.personLabel || d.personLabel
          };
        }
        return d;
      });
      parties = merged;
      savePartiesToStorage();
      renderParties();
    });
  }

  function addPartyManual() {
    var id = "p-manual-" + Date.now();
    parties.push({
      id: id,
      rawName: "New party",
      roleGuess: "other",
      personLabel: ""
    });
    savePartiesToStorage();
    renderParties();
  }

  function partiesContext() {
    return parties.map(function (p) {
      return {
        id: p.id,
        rawName: p.rawName,
        role: p.roleGuess,
        personLabel: p.personLabel
      };
    });
  }

  function setSuggestion(text) {
    var el = document.getElementById("ai-result");
    if (!el) return;
    // textarea — assign value only (never innerHTML)
    el.value = String(text == null ? "" : text);
    var can = !!(el.value && el.value.trim());
    var a = document.getElementById("btn-apply-insert");
    var b = document.getElementById("btn-apply-redline");
    if (a) a.disabled = !can;
    if (b) b.disabled = !can;
  }

  function runAi(mode) {
    var promptEl = document.getElementById("ai-prompt");
    var prompt = promptEl ? promptEl.value.trim() : "";
    if (!prompt && mode !== "suggest") {
      alert("Enter a prompt in “Ask AI to…”");
      return;
    }
    if (!prompt) {
      prompt = "Give concise editing suggestions for the selected (or full) text. Do not invent facts. Return plain suggested replacement or notes.";
    }

    var token = getBridgeTokenInput() || (window.DealQuillGrok && DealQuillGrok.getToken()) || "";
    if (!token) {
      alert("Set the bridge token (same as DEAL_QUILL_BRIDGE_TOKEN in bridge/.env.local). Keys never go in the add-in.");
      return;
    }
    if (window.DealQuillGrok) DealQuillGrok.saveTokenAsync(token);

    getSelectionOrBodyText(function (err, contextText) {
      if (err) {
        alert("Could not read selection: " + (err.message || err));
        return;
      }
      setAiLoading(true, mode === "ask" ? "ask" : "suggest");
      setSuggestion("");
      DealQuillGrok.complete({
        prompt: prompt,
        context: contextText,
        parties: partiesContext(),
        mode: mode || "suggest",
        token: token,
        bridgeUrl: DealQuillGrok.getBridgeUrl()
      }).then(function (data) {
        var text = (data && (data.text || data.suggestion || data.content)) || "";
        setAiLoading(false);
        setSuggestion(text);
        setStatus(true, "Bridge OK");
        showToast(text ? "Suggestion ready" : "Empty response", "success");
      }).catch(function (e) {
        setAiLoading(false);
        setSuggestion("");
        setStatus(false, e.message || "Bridge error");
        alert(e.message || String(e));
      });
    });
  }

  function pingBridge() {
    if (!window.DealQuillGrok) return;
    var token = getBridgeTokenInput() || DealQuillGrok.getToken() || "";
    DealQuillGrok.health(DealQuillGrok.getBridgeUrl(), token)
      .then(function (h) {
        setStatus(!!(h && (h.ok || h.status === "ok")), h && h.message ? h.message : "Bridge reachable");
      })
      .catch(function () {
        setStatus(false, "Bridge offline — run: cd bridge && npm start");
      });
  }

  function bindUi() {
    var strike = document.getElementById("btn-strike");
    var redline = document.getElementById("btn-redline");
    var suggest = document.getElementById("btn-suggest");
    var ask = document.getElementById("btn-ask");
    var applyIns = document.getElementById("btn-apply-insert");
    var applyRed = document.getElementById("btn-apply-redline");
    var refresh = document.getElementById("btn-refresh-parties");
    var addP = document.getElementById("btn-add-party");
    var urlEl = document.getElementById("bridge-url");
    var tokEl = document.getElementById("bridge-token");

    if (strike) {
      strike.addEventListener("click", function () {
        applyStrike();
        flashOk(strike);
        showToast("Strike applied", "success");
      });
    }
    if (redline) {
      redline.addEventListener("click", function () {
        var t = document.getElementById("redline-text");
        var val = t ? t.value : "";
        applyRedline(val);
        if (String(val || "").trim()) {
          flashOk(redline);
          showToast("Redline applied", "success");
        }
      });
    }
    if (suggest) suggest.addEventListener("click", function () { runAi("suggest"); });
    if (ask) ask.addEventListener("click", function () { runAi("ask"); });
    if (applyIns) {
      applyIns.addEventListener("click", function () {
        var el = document.getElementById("ai-result");
        applyInsert(el ? el.value : "");
        flashOk(applyIns);
        showToast("Inserted into document", "success");
      });
    }
    if (applyRed) {
      applyRed.addEventListener("click", function () {
        var el = document.getElementById("ai-result");
        applyRedline(el ? el.value : "");
        flashOk(applyRed);
        showToast("Redline applied", "success");
      });
    }
    if (refresh) refresh.addEventListener("click", refreshParties);
    if (addP) addP.addEventListener("click", addPartyManual);

    if (urlEl) {
      urlEl.addEventListener("change", function () {
        if (window.DealQuillGrok) DealQuillGrok.setBridgeUrl(urlEl.value.trim());
        pingBridge();
      });
    }
    if (tokEl) {
      tokEl.addEventListener("change", function () {
        if (window.DealQuillGrok) DealQuillGrok.saveTokenAsync(tokEl.value.trim(), pingBridge);
      });
    }
  }

  function initPartiesOnReady() {
    var stored = loadPartiesFromStorage();
    if (stored && stored.length) {
      parties = stored;
      renderParties();
    }
    refreshParties();
  }


  Office.onReady(function (info) {
    bindUi();

    if (window.DealQuillGrok) {
      DealQuillGrok.loadTokenAsync(function (tok) {
        var el = document.getElementById("bridge-token");
        if (el && tok) el.value = tok;
        var urlEl = document.getElementById("bridge-url");
        if (urlEl) urlEl.value = DealQuillGrok.getBridgeUrl();
        pingBridge();
      });
    }

    if (info.host === Office.HostType.Word) {
      initPartiesOnReady();
    } else {
      // Browser preview / sideload debug without Word
      parties = window.DealQuillPartyDetect
        ? DealQuillPartyDetect.detectParties(
            "This Agreement is between Acme Holdings LLC and Beta Tenant Inc.\nBuyer: Jordan Lee\nSeller: North Park Realty Corp."
          )
        : [];
      renderParties();
      setStatus(null, "Open in Word for document integration");
    }
  });
