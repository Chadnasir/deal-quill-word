/** Deal Quill — loads taskpane-a.js + taskpane-b.js if this file is included alone. */
(function () {
  "use strict";
  function load(src, cb) {
    var s = document.createElement("script");
    s.src = src;
    s.onload = cb;
    s.onerror = function () { console.error("Deal Quill failed to load", src); };
    document.head.appendChild(s);
  }
  var base = (document.currentScript && document.currentScript.src) || "";
  var dir = base.replace(/[^\/]+$/, "");
  load(dir + "taskpane-a.js", function () {
    load(dir + "taskpane-b.js");
  });
})();
