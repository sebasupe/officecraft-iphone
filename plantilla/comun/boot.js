// Pre-arranque síncrono para iPhone (corre antes de inicializar WebAssembly/eframe):
//   1. Limita devicePixelRatio (2× por defecto en iPhone 13 Pro de 3×) para 2.25× menos píxeles GPU
//      y cero tirones a 120 Hz, igual que PdfCraft para iPhone.
//   2. Preconfigura los ajustes del motor en localStorage (desactiva permanentemente el aviso
//      "Get the Missing Fonts" y la mini-barra de escritorio en WordCraft, y oculta barras en GridCraft).
//   3. Sincroniza <meta name="theme-color"> dinámicamente con la barra de estado del iPhone.
(function () {
  "use strict";

  var app = window.OC_APP || { id: "app", name: "OfficeCraft" };
  var prefsKey = "officecraft-" + app.id + "-prefs";

  // 1. Límite de DPR (Fluida = 2×, Rápida = 1.5×, Máxima = nativo 3×)
  try {
    var raw = JSON.parse(localStorage.getItem(prefsKey) || "{}");
    var calidad = raw.calidad || "fluida";
    var realDpr = window.devicePixelRatio || 1;
    window.__ocNativeDpr = realDpr;
    var cap = calidad === "maxima" ? realDpr : calidad === "rapida" ? 1.5 : 2;
    var target = Math.min(realDpr, cap);
    window.__ocActiveDpr = target;
    if (realDpr > target + 0.01) {
      Object.defineProperty(window, "devicePixelRatio", {
        get: function () { return target; },
        configurable: true,
      });
      if (typeof ResizeObserverEntry !== "undefined" && typeof ResizeObserverSize !== "undefined") {
        var dpDesc = Object.getOwnPropertyDescriptor(ResizeObserverEntry.prototype, "devicePixelContentBoxSize");
        var inDesc = Object.getOwnPropertyDescriptor(ResizeObserverSize.prototype, "inlineSize");
        var blDesc = Object.getOwnPropertyDescriptor(ResizeObserverSize.prototype, "blockSize");
        if (dpDesc && dpDesc.get && inDesc && inDesc.get && blDesc && blDesc.get) {
          var dpGet = dpDesc.get, inGet = inDesc.get, blGet = blDesc.get;
          var scaled = new WeakMap();
          Object.defineProperty(ResizeObserverEntry.prototype, "devicePixelContentBoxSize", {
            configurable: true,
            get: function () {
              var arr = dpGet.call(this);
              if (arr && arr.length) {
                var css = this.contentBoxSize;
                var k = target / realDpr;
                for (var i = 0; i < arr.length; i++) {
                  if (css && css.length) {
                    var c = css[Math.min(i, css.length - 1)];
                    scaled.set(arr[i], [
                      Math.max(1, Math.round(inGet.call(c) * target)),
                      Math.max(1, Math.round(blGet.call(c) * target)),
                    ]);
                  } else {
                    scaled.set(arr[i], [
                      Math.max(1, Math.round(inGet.call(arr[i]) * k)),
                      Math.max(1, Math.round(blGet.call(arr[i]) * k)),
                    ]);
                  }
                }
              }
              return arr;
            },
          });
          Object.defineProperty(ResizeObserverSize.prototype, "inlineSize", {
            configurable: true,
            get: function () { var s = scaled.get(this); return s ? s[0] : inGet.call(this); },
          });
          Object.defineProperty(ResizeObserverSize.prototype, "blockSize", {
            configurable: true,
            get: function () { var s = scaled.get(this); return s ? s[1] : blGet.call(this); },
          });
        }
      }
    }
  } catch (e) {}

  // 2. Ajustes nativos de cada motor (bloquea diálogos de escritorio invasivos en iPhone)
  try {
    if (app.id === "word") {
      // WordCraft 0.19.x lee JSON en "wordcraft.ui"; versiones futuras (0.20+) leen TOML en "wordcraft/config.toml".
      // Parcheamos ambos para suprimir "Get the Missing Fonts" y la mini-barra de PC hoy y tras futuras actualizaciones.
      var wcJsonKey = "wordcraft.ui";
      var wcObj = {};
      try {
        var existingJson = localStorage.getItem(wcJsonKey);
        if (existingJson) wcObj = JSON.parse(existingJson) || {};
      } catch (eJson) {
        wcObj = {};
      }
      wcObj.fontDownloadNotice = false;
      wcObj.startScreen = false;
      wcObj.miniToolbar = false;
      wcObj.ribbonCollapsed = true;
      wcObj.quickAccessToolbar = false;
      wcObj.keytipsEnabled = false;
      localStorage.setItem(wcJsonKey, JSON.stringify(wcObj));

      var wcTomlKey = "wordcraft/config.toml";
      var cur = localStorage.getItem(wcTomlKey);
      var overrides = {
        fontDownloadNotice: "false",
        startScreen: "false",
        miniToolbar: "false",
        ribbonCollapsed: "true",
        quickAccessToolbar: "false",
        keytipsEnabled: "false",
      };
      if (cur) {
        var parts = cur.split(/\r?\n/);
        var firstSec = parts.findIndex(function (l) { return /^\s*\[/.test(l); });
        var head = firstSec === -1 ? parts.slice() : parts.slice(0, firstSec);
        var tail = firstSec === -1 ? [] : parts.slice(firstSec);
        if (!head.some(function (l) { return /^\s*schema_version\s*=/.test(l); })) {
          head.unshift("schema_version = 1");
        }
        Object.keys(overrides).forEach(function (k) {
          var re = new RegExp("^\\s*" + k + "\\s*=");
          var idx = head.findIndex(function (l) { return re.test(l); });
          var line = k + " = " + overrides[k];
          if (idx >= 0) head[idx] = line;
          else head.push(line);
        });
        localStorage.setItem(wcTomlKey, head.concat(tail).join("\n"));
      }
    } else if (app.id === "grid") {
      if (!/[?&]chrome=/.test(location.search)) {
        var sep = location.search ? "&" : "?";
        history.replaceState(history.state, "", location.pathname + location.search + sep + "chrome=none" + location.hash);
      }
      var gcKey = "gridcraft.ui";
      var gRaw = localStorage.getItem(gcKey);
      var gObj = {};
      if (gRaw) {
        try { gObj = JSON.parse(gRaw) || {}; } catch (e2) { gObj = {}; }
      }
      gObj.ribbon_collapsed = true;
      gObj.ribbonCollapsed = true;
      gObj.formula_bar = false;
      gObj.formulaBar = false;
      localStorage.setItem(gcKey, JSON.stringify(gObj));
    }
  } catch (e) {}

  // 3. Sincronización dinámica de la barra de estado del iPhone (<meta name="theme-color">)
  var BRAND_LIGHT = { word: "#185abd", deck: "#c43e1c", grid: "#107c41" };
  var BRAND_DARK = { word: "#1b4ea0", deck: "#b53817", grid: "#0f6b38" };

  window.ocSyncThemeColor = function (inEditor, customHex) {
    try {
      var html = document.documentElement;
      var themeAttr = html.getAttribute("data-theme");
      var isDark =
        themeAttr === "dark" ||
        (themeAttr !== "light" && window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").matches);
      var color = customHex;
      if (!color) {
        if (inEditor) {
          color = (isDark ? BRAND_DARK : BRAND_LIGHT)[app.id] || "#185abd";
        } else {
          color = isDark ? "#000000" : "#f3f4f6";
        }
      }
      var metas = document.querySelectorAll('meta[name="theme-color"]');
      if (!metas.length) {
        var m = document.createElement("meta");
        m.name = "theme-color";
        m.content = color;
        document.head.appendChild(m);
      } else {
        metas.forEach(function (m) {
          m.removeAttribute("media");
          m.setAttribute("content", color);
        });
      }
    } catch (e) {}
  };
})();
