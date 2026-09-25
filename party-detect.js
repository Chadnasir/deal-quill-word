/**
 * Deal Quill — heuristic party extraction from plain document text.
 * Names/entities only. No contract-type taxonomy or CRE clause libraries.
 */
(function (global) {
  "use strict";

  var ROLE_LABELS = [
    "buyer", "seller", "landlord", "tenant", "lender", "borrower", "other"
  ];

  var ROLE_PATTERNS = [
    { re: /\b(buyer|purchaser)\b/i, role: "buyer" },
    { re: /\b(seller|vendor)\b/i, role: "seller" },
    { re: /\b(landlord|lessor)\b/i, role: "landlord" },
    { re: /\b(tenant|lessee)\b/i, role: "tenant" },
    { re: /\b(lender|mortgagee)\b/i, role: "lender" },
    { re: /\b(borrower|mortgagor)\b/i, role: "borrower" }
  ];

  var ENTITY_SUFFIX =
    /\b([A-Z][A-Za-z0-9&.'\-]+(?:\s+[A-Z][A-Za-z0-9&.'\-]+){0,6}\s*,?\s*(?:LLC|L\.?L\.?C\.?|Inc\.?|Incorporated|Corp\.?|Corporation|L\.?P\.?|LLP|Ltd\.?|Limited|Co\.|Company|Trust|Partnership))\b/g;

  var LABELED_LINE =
    /^(Buyer|Seller|Landlord|Tenant|Lender|Borrower|Purchaser|Vendor|Lessor|Lessee|Party\s*[A-Z0-9]+)\s*[:\-–]\s*(.+)$/gim;

  var BETWEEN =
    /\bbetween\s+(.+?)\s+and\s+(.+?)(?:\s*[,;(]|\s+(?:dated|effective|as\s+of)\b|$)/gi;

  var PARTY_AB =
    /\bParty\s+([A-Z])\s*[:\-–]?\s*([A-Z][A-Za-z0-9&.,'\- ]{2,80})/g;

  function slugId(name, i) {
    var base = String(name || "party")
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, "-")
      .replace(/^-|-$/g, "")
      .slice(0, 40);
    return "p-" + (base || "party") + "-" + i;
  }

  function guessRole(text) {
    var t = String(text || "");
    for (var i = 0; i < ROLE_PATTERNS.length; i++) {
      if (ROLE_PATTERNS[i].re.test(t)) return ROLE_PATTERNS[i].role;
    }
    return "other";
  }

  function cleanName(s) {
    return String(s || "")
      .replace(/\s+/g, " ")
      .replace(/^[\s,;:.\-–"']+|[\s,;:.\-–"']+$/g, "")
      .trim()
      .slice(0, 120);
  }

  function addParty(map, rawName, roleGuess, context) {
    var name = cleanName(rawName);
    if (!name || name.length < 2) return;
    if (/^(the|a|an|and|or|of|to|this|that|agreement|contract)$/i.test(name)) return;
    var key = name.toLowerCase();
    if (map[key]) {
      if (map[key].roleGuess === "other" && roleGuess !== "other") {
        map[key].roleGuess = roleGuess;
      }
      return;
    }
    var role = roleGuess !== "other" ? roleGuess : guessRole(context || name);
    map[key] = {
      id: slugId(name, Object.keys(map).length + 1),
      rawName: name,
      roleGuess: role,
      personLabel: name
    };
  }

  function detectParties(text) {
    var src = String(text || "").slice(0, 200000);
    var map = {};
    var m;

    LABELED_LINE.lastIndex = 0;
    while ((m = LABELED_LINE.exec(src)) !== null) {
      addParty(map, m[2], guessRole(m[1]), m[0]);
    }

    BETWEEN.lastIndex = 0;
    while ((m = BETWEEN.exec(src)) !== null) {
      addParty(map, m[1], "other", m[0]);
      addParty(map, m[2], "other", m[0]);
    }

    PARTY_AB.lastIndex = 0;
    while ((m = PARTY_AB.exec(src)) !== null) {
      addParty(map, m[2], "other", "Party " + m[1]);
    }

    ENTITY_SUFFIX.lastIndex = 0;
    var entityHits = 0;
    while ((m = ENTITY_SUFFIX.exec(src)) !== null && entityHits < 24) {
      addParty(map, m[1], "other", m[0]);
      entityHits++;
    }

    return Object.keys(map).map(function (k) {
      return map[k];
    });
  }

  global.DealQuillPartyDetect = {
    detectParties: detectParties,
    ROLE_LABELS: ROLE_LABELS
  };
})(typeof window !== "undefined" ? window : this);
