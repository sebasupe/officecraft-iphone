// GridCraft para iPhone: interfaz móvil fiel a Microsoft Excel para iOS sobre el motor
// oficial de GridCraft (data-host="page" + window.gridcraft) con gestos a 120 fps sin tirones.
(function () {
  "use strict";

  const $ = (id) => document.getElementById(id);
  const { icon, h, sheet, list, seg, swRow, colorGrid, sizeText, describe, makeSearchBox, sleep } = window.ocUI;
  const store = window.ocStore;
  const toast = (t, ms) => window.ocToast(t, ms);
  const canvas = $("gridcraft_canvas");
  const body = document.body;

  // ---------- Preferencias ----------
  const PREFS_KEY = "officecraft-grid-prefs";
  const prefs = Object.assign(
    {
      tema: "auto",
      calidad: "fluida",
      vistaArchivos: "lista",
      orden: "reciente",
      pantallaActiva: false,
    },
    readPrefs()
  );
  function readPrefs() {
    try { return JSON.parse(localStorage.getItem(PREFS_KEY) || "{}"); } catch (e) { return {}; }
  }
  function savePrefs() {
    try { localStorage.setItem(PREFS_KEY, JSON.stringify(prefs)); } catch (e) {}
  }
  function syncThemeBar() {
    if (window.ocSyncThemeColor) window.ocSyncThemeColor(body.classList.contains("oc-in-editor"));
  }
  function applyTheme() {
    if (prefs.tema === "auto") document.documentElement.removeAttribute("data-theme");
    else document.documentElement.setAttribute("data-theme", prefs.tema === "oscuro" ? "dark" : "light");
    syncThemeBar();
  }
  applyTheme();

  // ---------- Puente GridCraft API ----------
  const gc = {
    ok: () => !!(window.gridcraft && typeof window.gridcraft.request === "function"),
    ui: () => window.gridcraft.uiInspect(),
    doc: () => window.gridcraft.docInspect(),
    run: (command, params) => window.gridcraft.execute(command, params || {}),
    setUi: (params) => window.gridcraft.request("ui.set", params || {}),
    cellGet: (cell) => window.gridcraft.request("cell.get", cell ? { cell } : {}),
    stats: () => window.gridcraft.request("selection.stats", {}),
    key: (key, opts) => window.gridcraft.request("ui.key", Object.assign({ key }, opts || {})),
  };

  // ---------- Estado ----------
  const S = {
    ready: false,
    appVersion: (window.OC_APP && window.OC_APP.version) || "",
    tab: "inicio",
    homeSearchOpen: false,
    query: "",
    filesFolder: null,
    files: [],
    folders: [],
    doc: null,              // { meta, name, dirty, dirtySince, activeCell, range, sheets, activeSheet, zoom }
    opening: false,
    desktop: false,
    modal: false,
    drawerOpen: false,
    drawerTab: "inicio",    // "inicio" | "insertar" | "formulas" | "datos" | "vista"
    fxFocused: false,
    fxOriginal: "",
    ratio: 1,
    zoomPct: 100,
    crop: { l: 0, t: 0, w: 0, h: 0 },
    polling: false,
    saveWaiters: [],
    lastUi: null,
  };

  // ---------- Eventos del modo anfitrión (data-host="page") ----------
  canvas.addEventListener("gridcraft:ready", (e) => {
    if (e.detail && e.detail.appVersion) S.appVersion = e.detail.appVersion;
  });
  canvas.addEventListener("gridcraft:dirty", (e) => {
    if (!S.doc) return;
    setDirty(!!(e.detail && e.detail.dirty));
  });
  canvas.addEventListener("gridcraft:failed", (e) => {
    window.ocShowError("GridCraft no pudo iniciar: " + ((e.detail && e.detail.error) || "sin WebGL2/WebGPU"));
  });
  canvas.addEventListener("gridcraft:save", (e) => {
    onEngineSave(e.detail || {}).catch((err) => toast("Error al guardar: " + err.message, 4500));
  });

  function setDirty(dirty) {
    if (!S.doc || S.doc.dirty === dirty) return;
    S.doc.dirty = dirty;
    S.doc.dirtySince = dirty ? Date.now() : 0;
    renderTop();
  }

  async function onEngineSave(d) {
    const rawName = d.name || (S.doc ? S.doc.meta.nombre : "Libro.xlsx");
    const isExport = !!d.export || /\.pdf$/i.test(rawName);
    const bytes = d.bytes instanceof ArrayBuffer ? d.bytes : (d.bytes && d.bytes.buffer) || new ArrayBuffer(0);
    const file = new File([bytes], rawName, { type: window.ocMimeFor(rawName) });

    const waiters = S.saveWaiters.splice(0, S.saveWaiters.length);
    if (!isExport && S.doc) {
      S.doc.meta.tamano = file.size;
      S.doc.meta.modificado = Date.now();
      const saveFile = new File([bytes], S.doc.meta.nombre, { type: window.ocMimeFor(S.doc.meta.nombre) });
      await store.save(S.doc.meta, saveFile);
      S.doc.dirty = false;
      S.doc.dirtySince = 0;
      renderTop();
      waiters.forEach((w) => w.resolve(saveFile));
      return;
    }
    waiters.forEach((w) => w.resolve(file));
    if (!waiters.length) await shareFile(file);
  }

  function saveActive() {
    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => {
        S.saveWaiters = S.saveWaiters.filter((w) => w !== waiter);
        reject(new Error("GridCraft no respondió al guardar"));
      }, 20000);
      const waiter = {
        resolve: (f) => { clearTimeout(timer); resolve(f); },
        reject: (e) => { clearTimeout(timer); reject(e); },
      };
      S.saveWaiters.push(waiter);
      gc.run("file.save").catch((e) => {
        clearTimeout(timer);
        S.saveWaiters = S.saveWaiters.filter((w) => w !== waiter);
        reject(e);
      });
    });
  }

  async function currentFile() {
    if (!S.doc) return null;
    if (S.doc.dirty) return saveActive();
    const b = await store.blob(S.doc.meta.id);
    if (b) return new File([b], S.doc.meta.nombre, { type: window.ocMimeFor(S.doc.meta.nombre) });
    return saveActive();
  }

  async function shareFile(file) {
    if (!file) return;
    const ok = await window.ocShareFile(file);
    if (!ok) {
      sheet({
        title: "Archivo listo",
        text: file.name,
        build: (close) => [
          h("button", { class: "oc-primary", type: "button", text: "Compartir o guardar en Archivos", onclick: () => { close(); window.ocShareFile(file); } }),
        ],
      });
    }
  }
  window.ocOfferShare = shareFile;

  // ---------- Arranque ----------
  let gestures = null;
  window.addEventListener("TrunkApplicationStarted", () => {
    boot().catch((e) => {
      console.error(e);
      window.ocShowError("Error al iniciar GridCraft: " + (e && e.message ? e.message : e));
    });
  });

  async function boot() {
    for (let i = 0; i < 400 && !gc.ok(); i++) await sleep(50);
    if (!gc.ok()) throw new Error("No se inicializó window.gridcraft");
    for (let i = 0; i < 80; i++) {
      try { S.lastUi = await gc.ui(); break; } catch (e) { await sleep(150); }
    }
    S.ready = true;

    const quiet = (p) => p.catch(() => {});
    quiet(gc.run("app.language.set", { language: "es" }));
    quiet(gc.setUi({ ribbonCollapsed: true, formulaBar: false, animations: false }));
    quiet(gc.run("view.statusBar", { on: false }));
    quiet(gc.run("view.showSheetTabs", { on: false }));
    quiet(gc.run("view.theme", { mode: prefs.tema === "oscuro" ? "dark" : prefs.tema === "claro" ? "light" : "system" }));

    gestures = window.ocGestures.attach(canvas, {
      mode: () => (S.desktop ? "raton-zoom" : S.modal ? "raton" : "ver"),
      ratio: () => S.ratio,
      axisLock: "grid",
      onPinchCommit: async ({ scale }) => {
        if (!S.doc || S.desktop || S.modal) return;
        const next = Math.max(40, Math.min(350, Math.round((S.zoomPct || 100) * scale)));
        S.zoomPct = next;
        await gc.run("view.zoom", { percent: next }).catch(() => {});
        updateZoomPill();
      },
      onTap: () => {
        if (S.drawerOpen) closeDrawer();
        // Sincroniza inmediatamente la barra de fórmulas nativa tras tocar una celda
        setTimeout(() => syncActiveCell(), 60);
      },
      onDoubleTap: () => {
        if (!S.doc || S.desktop || S.modal) return;
        // En Safari iOS, focus() DEBE ejecutarse síncronamente dentro del toque para abrir el teclado
        const inp = $("oc-fx-input");
        if (inp) {
          inp.focus();
          try { inp.select(); } catch (e) {}
        }
        syncActiveCell();
      },
      onLongPress: (x, y) => showPressRing(x, y),
    });

    window.addEventListener("oc:viewport", () => {
      if (S.ready && body.classList.contains("oc-in-editor")) {
        gc.ui().then(applyCanvasCrop).catch(() => {});
      }
    });

    initStaticControls();
    await reloadLibrary();
    renderShell();
    window.ocHideSplash();

    const qs = new URLSearchParams(location.search);
    const openId = qs.get("open");
    if (openId) {
      const m = S.files.find((x) => x.id === openId);
      if (m) openStoredBook(m);
    } else if (qs.get("new") === "1") {
      createFromTemplate("blank");
    }

    setInterval(tick, 420);
    document.addEventListener("visibilitychange", () => {
      window.ocSyncWakeLock(prefs.pantallaActiva && body.classList.contains("oc-in-editor"));
      if (document.visibilityState === "hidden" && S.doc && S.doc.dirty && !S.modal) {
        saveActive().catch(() => {});
      }
    });
  }

  function showPressRing(clientX, clientY) {
    const ring = $("oc-press-ring");
    const box = $("oc-doc");
    if (!ring || !box) return;
    const r = box.getBoundingClientRect();
    ring.style.left = (clientX - r.left) + "px";
    ring.style.top = (clientY - r.top) + "px";
    ring.hidden = false;
    ring.style.animation = "none";
    void ring.offsetWidth;
    ring.style.animation = "";
    clearTimeout(showPressRing.t);
    showPressRing.t = setTimeout(() => { ring.hidden = true; }, 400);
  }

  // ---------- Recorte preciso conservando encabezados A-B-C y filas 1-2-3 ----------
  function applyCanvasCrop(u) {
    const docEl = $("oc-doc");
    if (!docEl) return;
    if (S.desktop || S.modal || !u || !u.window || !u.grid) {
      if (S.crop.l || S.crop.t || S.crop.w || S.crop.h) {
        S.crop = { l: 0, t: 0, w: 0, h: 0 };
        docEl.style.setProperty("--oc-crop-left", "0px");
        docEl.style.setProperty("--oc-crop-top", "0px");
        docEl.style.setProperty("--oc-canvas-w", "100%");
        docEl.style.setProperty("--oc-canvas-h", "100%");
      }
      return;
    }
    const [ww, wh] = u.window;
    const [, gy, , gh] = u.grid;
    if (!(ww > 100 && wh > 100 && gh > 50)) return;
    const ratio = canvas.clientWidth / ww || 1;
    S.ratio = ratio;

    // `u.grid` es cells_rect: los encabezados de columna (A, B, C...) miden ~24-26 pt justo encima de gy.
    const colHeaderH = 26;
    const chromeT = Math.max(0, Math.min(wh * 0.45, gy - colHeaderH));
    const chromeB = Math.max(0, Math.min(wh * 0.25, wh - (gy + gh) - 1));

    const vw = docEl.clientWidth;
    const vh = docEl.clientHeight;
    if (!(vw > 50 && vh > 50)) return;

    const top = -(chromeT * ratio);
    const height = Math.max(vh, vh + (chromeT + chromeB) * ratio);

    if (Math.abs(top - S.crop.t) > 1.5 || Math.abs(height - S.crop.h) > 1.5 || S.crop.w !== vw) {
      S.crop = { l: 0, t: top, w: vw, h: height };
      docEl.style.setProperty("--oc-crop-left", "0px");
      docEl.style.setProperty("--oc-crop-top", top.toFixed(1) + "px");
      docEl.style.setProperty("--oc-canvas-w", "100%");
      docEl.style.setProperty("--oc-canvas-h", height.toFixed(1) + "px");
    }
  }

  // ---------- Sondeo sin tirones ----------
  async function tick() {
    if (!S.ready || S.polling || document.visibilityState !== "visible") return;
    if (gestures && gestures.busy()) return;
    S.polling = true;
    try {
      const u = await gc.ui();
      S.lastUi = u;
      if (u.window && u.window[0]) S.ratio = canvas.clientWidth / u.window[0];
      const modal = !!(u.dialog || u.backstage);
      if (modal !== S.modal) {
        S.modal = modal;
        body.classList.toggle("oc-modal", modal);
        if (modal) closeDrawer();
      }
      if (!body.classList.contains("oc-in-editor")) return;
      applyCanvasCrop(u);
      if (S.opening || !S.doc) return;

      await syncActiveCell();

      if (S.doc.dirty && S.doc.dirtySince && Date.now() - S.doc.dirtySince > 25000 && !S.modal && !S.saveWaiters.length) {
        S.doc.dirtySince = Date.now();
        saveActive().catch(() => {});
      }
    } catch (e) {
    } finally {
      S.polling = false;
    }
  }

  async function syncActiveCell() {
    if (!S.doc) return;
    try {
      const d = await gc.doc();
      if (!d) return;
      const rawSheets = Array.isArray(d.sheets) ? d.sheets : [];
      const sheets = rawSheets.length
        ? rawSheets.map((sh) => (typeof sh === "string" ? sh : sh.name || "Hoja"))
        : S.doc.sheets;
      const activeIdx = typeof d.activeSheet === "number" ? d.activeSheet : 0;
      const activeSheet =
        typeof d.activeSheet === "string" ? d.activeSheet : (sheets[activeIdx] || sheets[0] || "Sheet1");

      const activeSheetObj = rawSheets[activeIdx];
      if (activeSheetObj && typeof activeSheetObj.zoom === "number" && activeSheetObj.zoom !== S.zoomPct) {
        S.zoomPct = activeSheetObj.zoom;
        updateZoomPill();
      }

      // En GridCraft document.inspect: `activeCell` es "B7" y `selection` es "B7" o "A1:D7"
      const selStr = typeof d.selection === "string" ? d.selection : (d.selection && d.selection.range) || "";
      const active =
        d.activeCell ||
        (d.selection && typeof d.selection === "object" && d.selection.active) ||
        (selStr ? selStr.split(":")[0] : "A1");
      const range = selStr || active;

      const sheetChanged =
        activeSheet !== S.doc.activeSheet ||
        sheets.length !== S.doc.sheets.length ||
        sheets.some((n, idx) => n !== S.doc.sheets[idx]);

      S.doc.sheets = sheets;
      S.doc.activeSheet = activeSheet;
      if (sheetChanged) renderSheetsBar();

      const cellChanged = active !== S.doc.activeCell || range !== S.doc.range;
      S.doc.activeCell = active;
      S.doc.range = range;
      $("oc-cell-addr").textContent = range !== active ? range : active;

      if (!S.fxFocused) {
        const c = await gc.cellGet(active).catch(() => null);
        if (c) {
          const val = c.formula || (c.input != null && c.input !== "" ? String(c.input) : (c.text != null ? String(c.text) : (c.value != null ? String(c.value) : "")));
          if (cellChanged || $("oc-fx-input").value !== val) {
            $("oc-fx-input").value = val;
            S.fxOriginal = val;
          }
        }
      }

      // Barra de estadísticas rápidas cuando se selecciona un rango (Suma, Promedio, Recuento)
      const statsEl = $("oc-stats-bar");
      if (range && range.includes(":")) {
        const st = await gc.stats().catch(() => null);
        if (st && (st.count > 1 || st.numericCount > 1)) {
          const parts = [];
          if (st.sum != null) parts.push("Suma: " + formatStat(st.sum));
          if (st.average != null) parts.push("Prom.: " + formatStat(st.average));
          if (st.count != null) parts.push("Recuento: " + st.count);
          if (st.max != null) parts.push("Máx.: " + formatStat(st.max));
          statsEl.textContent = parts.join("   ·   ");
          statsEl.hidden = !parts.length;
        } else {
          statsEl.hidden = true;
        }
      } else {
        statsEl.hidden = true;
      }
    } catch (e) {}
  }

  function formatStat(n) {
    const num = Number(n);
    if (!Number.isFinite(num)) return String(n);
    return num.toLocaleString("es", { maximumFractionDigits: 2 });
  }

  function updateZoomPill() {
    const pill = $("oc-zoom-pill");
    if (!pill) return;
    if (!S.doc || S.desktop || S.modal) {
      pill.hidden = true;
      return;
    }
    pill.hidden = false;
    pill.textContent = (S.zoomPct || 100) + "%";
  }

  // ---------- Barra de fórmulas nativa (Excel iPhone) ----------
  function initFormulaBar() {
    const fxBtn = $("oc-fx-btn");
    const fxInput = $("oc-fx-input");
    const fxOk = $("oc-fx-ok");
    const fxCancel = $("oc-fx-cancel");
    fxBtn.insertAdjacentHTML("afterbegin", icon("fx"));
    fxOk.insertAdjacentHTML("afterbegin", icon("check"));
    fxCancel.insertAdjacentHTML("afterbegin", icon("close"));

    $("oc-cell-addr").addEventListener("click", goToCellSheet);
    fxBtn.addEventListener("click", functionCatalogSheet);

    fxInput.addEventListener("focus", () => {
      S.fxFocused = true;
      S.fxOriginal = fxInput.value;
      fxOk.hidden = false;
      fxCancel.hidden = false;
    });
    fxInput.addEventListener("blur", () => {
      setTimeout(() => {
        if (document.activeElement !== fxInput) {
          S.fxFocused = false;
          fxOk.hidden = true;
          fxCancel.hidden = true;
        }
      }, 140);
    });
    const commit = async (moveDown) => {
      if (!S.doc) return;
      const cell = S.doc.activeCell || "A1";
      const val = fxInput.value;
      S.fxFocused = false;
      fxOk.hidden = true;
      fxCancel.hidden = true;
      fxInput.blur();
      try {
        await gc.run("cell.set", { cell, input: val });
        if (moveDown) await gc.key("Enter").catch(() => {});
        await sleep(60);
        await syncActiveCell();
      } catch (e) {
        toast("Error en fórmula: " + (e.message || e), 4000);
      }
    };
    fxOk.addEventListener("mousedown", (e) => e.preventDefault());
    fxCancel.addEventListener("mousedown", (e) => e.preventDefault());
    fxOk.addEventListener("click", () => commit(false));
    fxCancel.addEventListener("click", () => {
      fxInput.value = S.fxOriginal;
      S.fxFocused = false;
      fxOk.hidden = true;
      fxCancel.hidden = true;
      fxInput.blur();
    });
    fxInput.addEventListener("keydown", (e) => {
      if (e.key === "Enter") {
        e.preventDefault();
        commit(true);
      } else if (e.key === "Escape") {
        e.preventDefault();
        fxInput.value = S.fxOriginal;
        fxInput.blur();
      }
    });
  }

  function goToCellSheet() {
    const inp = h("input", { class: "oc-field", type: "text", value: (S.doc && S.doc.activeCell) || "A1", placeholder: "Ej. B12 o A1:D20", autocapitalize: "characters" });
    sheet({
      title: "Ir a celda o rango",
      build: (close) => {
        const go = async () => {
          const ref = inp.value.trim().toUpperCase();
          if (!ref) return;
          close();
          await gc.run("selection.set", { range: ref }).catch(() => toast("Referencia no válida"));
          await syncActiveCell();
        };
        inp.addEventListener("keydown", (e) => { if (e.key === "Enter") go(); });
        return [inp, h("button", { class: "oc-primary", type: "button", text: "Ir", onclick: go })];
      },
    });
    inp.focus();
    inp.select();
  }

  function functionCatalogSheet() {
    const fns = [
      ["=SUM(", "SUM — Sumar un rango de celdas"],
      ["=AVERAGE(", "AVERAGE — Promedio aritmético"],
      ["=COUNT(", "COUNT — Contar números"],
      ["=MAX(", "MAX — Valor máximo"],
      ["=MIN(", "MIN — Valor mínimo"],
      ["=IF(", "IF — Condición lógica (si, entonces, sino)"],
      ["=SUMIF(", "SUMIF — Sumar según criterio"],
      ["=COUNTIF(", "COUNTIF — Contar según criterio"],
      ["=VLOOKUP(", "VLOOKUP — Buscar en tabla vertical"],
      ["=XLOOKUP(", "XLOOKUP — Búsqueda moderna"],
      ["=ROUND(", "ROUND — Redondear decimales"],
      ["=TODAY()", "TODAY — Fecha de hoy"],
    ];
    sheet({
      title: "Insertar función (fx)",
      build: (close) => [
        list(
          fns.map(([tpl, label]) => ({
            i: "fx",
            t: label,
            run: () => {
              close();
              const inp = $("oc-fx-input");
              inp.value = tpl;
              inp.focus();
              inp.setSelectionRange(tpl.length, tpl.length);
            },
          }))
        ),
      ],
    });
  }

  // ---------- Pestañas de hojas (Hoja1, Hoja2, +) ----------
  function renderSheetsBar() {
    const bar = $("oc-sheets-bar");
    if (!S.doc) { bar.replaceChildren(); return; }
    const nodes = S.doc.sheets.map((name) => {
      const isActive = name === S.doc.activeSheet;
      return h("button", {
        class: "oc-sheet-tab",
        type: "button",
        "aria-selected": String(isActive),
        text: name,
        onclick: () => (isActive ? sheetContextMenu(name) : activateSheet(name)),
      });
    });
    nodes.push(
      h("button", {
        class: "oc-sheet-add",
        type: "button",
        "aria-label": "Nueva hoja",
        icon: "plus",
        onclick: async () => {
          await execGrid("home.insertSheet");
          await sleep(100);
          await syncActiveCell();
        },
      })
    );
    bar.replaceChildren(...nodes);
  }

  async function activateSheet(name) {
    await execGrid("sheet.activate", { name });
    await sleep(80);
    await syncActiveCell();
  }

  function sheetContextMenu(name) {
    sheet({
      title: "Hoja «" + name + "»",
      build: (close) => [
        list([
          {
            i: "rename",
            t: "Renombrar hoja",
            run: () => {
              close();
              const inp = h("input", { class: "oc-field", type: "text", value: name });
              sheet({
                title: "Renombrar hoja",
                build: (c2) => [
                  inp,
                  h("button", {
                    class: "oc-primary",
                    type: "button",
                    text: "Guardar",
                    onclick: async () => {
                      const nn = inp.value.trim();
                      if (!nn) return;
                      c2();
                      await execGrid("sheet.rename", { name: nn, sheet: name });
                      await syncActiveCell();
                    },
                  }),
                ],
              });
              inp.focus();
              inp.select();
            },
          },
          { i: "plus", t: "Insertar nueva hoja", run: async () => { close(); await execGrid("home.insertSheet"); await syncActiveCell(); } },
        ]),
        S.doc.sheets.length > 1
          ? list([{ i: "trash", t: "Eliminar esta hoja", danger: true, run: async () => { close(); await execGrid("home.deleteSheet"); await syncActiveCell(); } }])
          : null,
      ].filter(Boolean),
    });
  }

  // ---------- Plantillas Excel para iPhone ----------
  const TEMPLATES = [
    { id: "blank", name: "Libro en blanco", style: "blank" },
    { id: "budget", name: "Presupuesto", style: "grid" },
    { id: "expenses", name: "Gastos", style: "grid" },
    { id: "tasks", name: "Lista de tareas", style: "grid" },
    { id: "schedule", name: "Horario", style: "grid" },
  ];

  // ---------- Shell Home ----------
  function initStaticControls() {
    const fab = $("oc-create-fab");
    fab.replaceChildren();
    fab.insertAdjacentHTML("afterbegin", icon("plus"));
    fab.addEventListener("click", openCreateSheet);
    $("oc-zoom-pill").addEventListener("click", zoomSheet);
    initFormulaBar();

    $("oc-file").addEventListener("change", async () => {
      const file = $("oc-file").files && $("oc-file").files[0];
      if (!file) return;
      $("oc-file").value = "";
      const meta = {
        id: store.newId(),
        app: "grid",
        nombre: file.name || "Libro.xlsx",
        tamano: file.size,
        abierto: Date.now(),
        modificado: file.lastModified || Date.now(),
      };
      try { await store.save(meta, file); } catch (e) {}
      openDocument(meta, file);
    });

    renderTabBar();
  }

  function renderTabBar() {
    const tabs = [
      { id: "inicio", t: "Recientes", i: "home" },
      { id: "destacados", t: "Destacados", i: "star" },
      { id: "archivos", t: "Abrir", i: "folder" },
    ];
    $("oc-tabbar").replaceChildren(
      ...tabs.map((tb) =>
        h("button", { class: "oc-tab-btn", type: "button", "aria-current": S.tab === tb.id ? "page" : null, onclick: () => switchTab(tb.id) }, h("span", { icon: tb.i }), h("span", { text: tb.t }))
      )
    );
  }

  async function switchTab(tab) {
    S.tab = tab;
    if (tab !== "archivos") S.filesFolder = null;
    renderTabBar();
    await reloadLibrary();
    renderShell();
  }

  async function reloadLibrary() {
    try { S.files = await store.list("grid"); } catch (e) { S.files = []; }
    try { S.folders = await store.listFolders("grid"); } catch (e) { S.folders = []; }
  }

  function sorted(arr) {
    const copy = arr.slice();
    if (prefs.orden === "nombre") copy.sort((a, b) => (a.nombre || "").localeCompare(b.nombre || "", "es"));
    else if (prefs.orden === "tamano") copy.sort((a, b) => (b.tamano || 0) - (a.tamano || 0));
    else copy.sort((a, b) => (b.abierto || 0) - (a.abierto || 0));
    return copy;
  }

  function renderShell() {
    if (S.tab === "inicio") renderHomeTab();
    else if (S.tab === "destacados") renderStarredTab();
    else renderFilesTab();
  }

  function renderHomeTab() {
    $("oc-app-head").replaceChildren(
      ...[
        h(
          "div",
          { class: "oc-head-row" },
          h("img", { class: "oc-brand-logo", src: "./icons/icon-192.png", alt: "" }),
          h("h1", { class: "oc-head-title", text: "GridCraft" }),
          h("button", {
            class: "oc-head-btn" + (S.homeSearchOpen ? " oc-active" : ""),
            type: "button",
            "aria-label": "Buscar",
            icon: "search",
            onclick: () => { S.homeSearchOpen = !S.homeSearchOpen; if (!S.homeSearchOpen) S.query = ""; renderHomeTab(); },
          }),
          h("button", { class: "oc-head-btn", type: "button", "aria-label": "Suite OfficeCraft", icon: "apps", onclick: () => { location.href = "../"; } }),
          h("button", { class: "oc-head-btn", type: "button", "aria-label": "Ajustes", icon: "settings", onclick: settingsSheet })
        ),
        S.homeSearchOpen ? makeSearchBox("Buscar hojas de cálculo…", S.query, (q) => { S.query = q; renderHomeBody(); }, true) : null,
      ].filter(Boolean)
    );
    renderHomeBody();
  }

  function renderHomeBody() {
    const el = $("oc-tab-body");
    const q = S.query.trim().toLowerCase();
    const nodes = [];
    if (!q) {
      nodes.push(
        h("div", { class: "oc-section-head" }, h("h2", { text: "Nuevo libro" }), h("button", { class: "oc-link-btn", type: "button", text: "Importar", onclick: pickFile })),
        h(
          "div",
          { class: "oc-tpl-strip" },
          TEMPLATES.map((tpl) => {
            const prev =
              tpl.style === "blank"
                ? h("div", { class: "oc-tpl-preview" }, h("span", { icon: "plus" }))
                : h("div", { class: "oc-tpl-preview" }, h("span", { icon: "table" }));
            return h("button", { class: "oc-tpl-card", type: "button", onclick: () => createFromTemplate(tpl.id) }, prev, h("b", { text: tpl.name }));
          })
        )
      );
    }

    let items = S.files;
    if (q) items = items.filter((m) => (m.nombre || "").toLowerCase().includes(q));
    items = sorted(items);

    nodes.push(
      h("div", { class: "oc-section-head" }, h("h2", { text: q ? "Resultados" : "Recientes" }), h("button", { class: "oc-sort-btn", type: "button", icon: "sort", onclick: sortSheet }, h("span", { text: prefs.orden === "nombre" ? "Nombre" : prefs.orden === "tamano" ? "Tamaño" : "Reciente" })))
    );

    if (items.length) nodes.push(renderFileList(items));
    else {
      nodes.push(
        h(
          "div",
          { class: "oc-empty" },
          h("div", { icon: "table" }),
          h("b", { text: q ? "Sin coincidencias" : "Tus libros Excel aparecerán aquí" }),
          h("div", { text: q ? "No encontramos «" + S.query + "»." : "Crea un libro nuevo o abre un archivo .xlsx / .csv desde Archivos / iCloud." }),
          !q ? h("button", { class: "oc-empty-btn", type: "button", text: "Nuevo libro en blanco", onclick: () => createFromTemplate("blank") }) : null
        )
      );
    }
    el.replaceChildren(...nodes);
  }

  function renderStarredTab() {
    $("oc-app-head").replaceChildren(h("div", { class: "oc-head-row" }, h("h1", { class: "oc-head-title", text: "Destacados" }), h("button", { class: "oc-head-btn", type: "button", icon: "settings", onclick: settingsSheet })));
    const items = sorted(S.files.filter((m) => m.destacado));
    $("oc-tab-body").replaceChildren(
      items.length ? renderFileList(items) : h("div", { class: "oc-empty" }, h("div", { icon: "star" }), h("b", { text: "Sin libros destacados" }))
    );
  }

  function renderFilesTab() {
    const folder = S.filesFolder ? S.folders.find((f) => f.id === S.filesFolder) : null;
    $("oc-app-head").replaceChildren(
      h(
        "div",
        { class: "oc-head-row" },
        S.filesFolder ? h("button", { class: "oc-head-back", type: "button", icon: "back", onclick: () => { S.filesFolder = null; renderFilesTab(); } }, h("span", { text: "Abrir" })) : null,
        h("h1", { class: "oc-head-title", text: folder ? folder.nombre : "Abrir" }),
        !S.filesFolder ? h("button", { class: "oc-head-btn", type: "button", icon: "folderPlus", onclick: () => newFolderSheet(null) }) : null
      ),
      makeSearchBox("Buscar hojas…", S.query, (q) => { S.query = q; renderFilesBody(); }, false)
    );
    renderFilesBody();
  }

  function renderFilesBody() {
    const el = $("oc-tab-body");
    const q = S.query.trim().toLowerCase();
    const nodes = [];
    if (!S.filesFolder && !q) {
      const total = S.files.reduce((a, m) => a + (m.tamano || 0), 0);
      nodes.push(
        h(
          "div",
          { class: "oc-loc-card" },
          locRow("open", "#107c41", "Archivos / iCloud Drive", "Abrir .xlsx, .xlsm, .xls, .ods o .csv", pickFile),
          locRow("iphone", "#185abd", "En este iPhone", S.files.length + " libros · " + sizeText(total), () => {})
        )
      );
      if (S.folders.length) {
        nodes.push(
          h("div", { class: "oc-section-head" }, h("h2", { text: "Carpetas" }), h("button", { class: "oc-link-btn", type: "button", text: "+ Nueva", onclick: () => newFolderSheet(null) })),
          h("div", { class: "oc-loc-card" }, S.folders.map((f) => locRow("folder", "#5a6375", f.nombre, S.files.filter((m) => m.carpeta === f.id).length + " archivos", () => { S.filesFolder = f.id; renderFilesTab(); })))
        );
      }
      nodes.push(h("div", { class: "oc-section-head" }, h("h2", { text: "Todos los libros" }), h("small", { text: String(S.files.length) })));
    }
    let items = S.files;
    if (S.filesFolder) items = items.filter((m) => m.carpeta === S.filesFolder);
    if (q) items = items.filter((m) => (m.nombre || "").toLowerCase().includes(q));
    items = sorted(items);
    nodes.push(items.length ? renderFileList(items) : h("div", { class: "oc-empty" }, h("div", { icon: "folder" }), h("b", { text: "Sin archivos" })));
    el.replaceChildren(...nodes);
  }

  function locRow(ic, color, title, sub, onclick) {
    const b = h("span", { class: "oc-loc-ico", icon: ic });
    b.style.background = color;
    return h("button", { class: "oc-loc-row", type: "button", onclick }, b, h("span", { class: "oc-loc-info" }, h("b", { text: title }), h("small", { text: sub })), h("span", { class: "oc-loc-chev", icon: "chevRight" }));
  }

  function renderFileList(items) {
    return h(
      "ul",
      { class: "oc-recents" },
      items.map((meta) => {
        const thumb = h("span", { class: "oc-thumb" });
        thumb.insertAdjacentHTML("afterbegin", icon("table"));
        thumb.append(h("span", { class: "oc-thumb-tag", text: "XLSX" }));
        return h(
          "li",
          { class: "oc-recent" },
          h("button", { class: "oc-recent-open", type: "button", onclick: () => openStoredBook(meta) }, thumb, h("span", { class: "oc-recent-text" }, h("b", { text: meta.nombre }), h("small", { text: describe(meta, meta.hojas ? meta.hojas + " hojas" : "") }))),
          h("button", { class: "oc-star-btn" + (meta.destacado ? " oc-starred" : ""), type: "button", [meta.destacado ? "iconFilled" : "icon"]: "star", onclick: () => toggleStar(meta) }),
          h("button", { class: "oc-recent-more", type: "button", icon: "more", onclick: () => fileMenu(meta) })
        );
      })
    );
  }

  async function toggleStar(meta) {
    meta.destacado = !meta.destacado;
    await store.putMeta(meta).catch(() => {});
    await reloadLibrary();
    if (!body.classList.contains("oc-in-editor")) renderShell();
  }

  function fileMenu(meta) {
    sheet({
      title: meta.nombre,
      text: describe(meta),
      build: (close) => [
        list([
          { i: "open", t: "Abrir", run: () => { close(); openStoredBook(meta); } },
          { i: "star", t: meta.destacado ? "Quitar de destacados" : "Destacar", run: () => { close(); toggleStar(meta); } },
          { i: "share", t: "Compartir / Guardar en Archivos", run: async () => { close(); const b = await store.blob(meta.id); if (b) shareFile(new File([b], meta.nombre, { type: window.ocMimeFor(meta.nombre) })); } },
          { i: "rename", t: "Renombrar", run: () => { close(); renameSheet(meta); } },
          { i: "copy", t: "Duplicar", run: () => { close(); duplicateBook(meta); } },
        ]),
        list([{ i: "trash", t: "Eliminar", danger: true, run: () => { close(); confirmDelete(meta); } }]),
      ],
    });
  }

  function renameSheet(meta) {
    const base = (meta.nombre || "Libro.xlsx").replace(/\.(xlsx|xlsm|xls|ods|csv)$/i, "");
    const inp = h("input", { class: "oc-field", type: "text", value: base });
    sheet({
      title: "Renombrar libro",
      build: (close) => {
        const go = async () => {
          const clean = inp.value.trim();
          if (!clean) return;
          meta.nombre = clean.replace(/\.[a-z0-9]+$/i, "") + ".xlsx";
          await store.putMeta(meta).catch(() => {});
          if (S.doc && S.doc.meta.id === meta.id) { S.doc.meta.nombre = meta.nombre; renderTop(); }
          await reloadLibrary();
          if (!body.classList.contains("oc-in-editor")) renderShell();
          close();
        };
        inp.addEventListener("keydown", (e) => { if (e.key === "Enter") go(); });
        return [inp, h("button", { class: "oc-primary", type: "button", text: "Guardar nombre", onclick: go })];
      },
    });
    inp.focus();
    inp.select();
  }

  async function duplicateBook(meta) {
    const b = await store.blob(meta.id);
    if (!b) return;
    const copyName = meta.nombre.replace(/(\.[^.]+)$/, " (copia)$1");
    await store.save(Object.assign({}, meta, { id: store.newId(), nombre: copyName, abierto: Date.now(), modificado: Date.now(), destacado: false }), b);
    await reloadLibrary();
    renderShell();
  }

  function newFolderSheet(moveMeta) {
    const inp = h("input", { class: "oc-field", type: "text", placeholder: "Nombre de la carpeta" });
    sheet({
      title: "Nueva carpeta",
      build: (close) => [
        inp,
        h("button", { class: "oc-primary", type: "button", text: "Crear", onclick: async () => {
          const n = inp.value.trim(); if (!n) return;
          const f = { id: store.newId(), app: "grid", nombre: n, creada: Date.now() };
          await store.saveFolder(f);
          if (moveMeta) { moveMeta.carpeta = f.id; await store.putMeta(moveMeta); }
          await reloadLibrary();
          renderShell();
          close();
        } }),
      ],
    });
    inp.focus();
  }

  function confirmDelete(meta) {
    sheet({
      title: "¿Eliminar libro?",
      text: "Se borrará «" + meta.nombre + "» de GridCraft en este iPhone.",
      build: (close) => [list([{ i: "trash", t: "Eliminar", danger: true, run: async () => { close(); await store.remove(meta.id); await reloadLibrary(); renderShell(); } }])],
    });
  }

  function sortSheet() {
    sheet({
      title: "Ordenar libros",
      build: (close) => [list([["reciente", "Reciente"], ["nombre", "Nombre"], ["tamano", "Tamaño"]].map(([v, t]) => ({ i: prefs.orden === v ? "check" : "sort", t, run: () => { prefs.orden = v; savePrefs(); close(); renderShell(); } })))],
    });
  }

  function openCreateSheet() {
    sheet({
      title: "Nuevo libro",
      build: (close) => [
        list([
          { i: "blank", t: "Libro en blanco", small: ".xlsx", run: () => { close(); createFromTemplate("blank"); } },
          { i: "open", t: "Abrir desde Archivos / iCloud", small: ".xlsx, .csv, .ods", run: () => { close(); pickFile(); } },
          { i: "table", t: "Presupuesto mensual", small: "Con fórmulas", run: () => { close(); createFromTemplate("budget"); } },
          { i: "table", t: "Registro de gastos", small: "Con fórmulas", run: () => { close(); createFromTemplate("expenses"); } },
          { i: "table", t: "Lista de tareas", small: "Plantilla", run: () => { close(); createFromTemplate("tasks"); } },
          { i: "table", t: "Horario semanal", small: "Plantilla", run: () => { close(); createFromTemplate("schedule"); } },
        ]),
      ],
    });
  }

  function pickFile() {
    const i = $("oc-file");
    i.value = "";
    i.click();
  }

  // ---------- Abrir / Crear libro ----------
  async function openStoredBook(meta) {
    const blob = await store.blob(meta.id);
    if (!blob) { toast("Ese archivo ya no existe."); await store.remove(meta.id); await reloadLibrary(); renderShell(); return; }
    meta.abierto = Date.now();
    store.putMeta(meta).catch(() => {});
    await openDocument(meta, blob);
  }

  async function createFromTemplate(tplId) {
    if (!S.ready) { toast("GridCraft todavía está iniciando"); return; }
    S.opening = true;
    showEditor();
    busy(true, "Creando libro…");
    try {
      await gc.run("file.new", {});
      await sleep(100);
      const stamp = new Date().toLocaleDateString("es", { day: "2-digit", month: "2-digit" }).replace("/", "-");
      const titles = {
        blank: "Libro " + stamp + ".xlsx",
        budget: "Presupuesto " + stamp + ".xlsx",
        expenses: "Gastos " + stamp + ".xlsx",
        tasks: "Tareas " + stamp + ".xlsx",
        schedule: "Horario " + stamp + ".xlsx",
      };

      if (tplId === "budget") {
        await gc.run("range.setValues", {
          range: "A1:D7",
          values: [
            ["Categoría", "Presupuesto", "Real", "Diferencia"],
            ["Vivienda", 900, 900, "=B2-C2"],
            ["Alimentación", 450, 410, "=B3-C3"],
            ["Transporte", 180, 165, "=B4-C4"],
            ["Servicios", 150, 142, "=B5-C5"],
            ["Ahorro", 320, 350, "=B6-C6"],
            ["Total", "=SUM(B2:B6)", "=SUM(C2:C6)", "=SUM(D2:D6)"],
          ],
        }).catch(() => {});
        await gc.run("selection.set", { range: "A1:D1" }).catch(() => {});
        await gc.run("home.bold").catch(() => {});
        await gc.run("home.fillColor", { color: "#107C41" }).catch(() => {});
        await gc.run("home.fontColor", { color: "#FFFFFF" }).catch(() => {});
        await gc.run("selection.set", { range: "A7:D7" }).catch(() => {});
        await gc.run("home.bold").catch(() => {});
        await gc.run("selection.set", { range: "A1:D7" }).catch(() => {});
        await gc.run("home.autofitColumnWidth").catch(() => {});
        await gc.run("selection.set", { range: "A1" }).catch(() => {});
      } else if (tplId === "expenses") {
        await gc.run("range.setValues", {
          range: "A1:D6",
          values: [
            ["Fecha", "Concepto", "Categoría", "Monto"],
            [stamp, "Supermercado", "Comida", 64.5],
            [stamp, "Café", "Ocio", 4.75],
            [stamp, "Gasolina", "Transporte", 45.0],
            ["", "", "Total", "=SUM(D2:D4)"],
            ["", "", "Promedio", "=AVERAGE(D2:D4)"],
          ],
        }).catch(() => {});
        await gc.run("selection.set", { range: "A1:D1" }).catch(() => {});
        await gc.run("home.bold").catch(() => {});
        await gc.run("selection.set", { range: "A1:D6" }).catch(() => {});
        await gc.run("home.autofitColumnWidth").catch(() => {});
        await gc.run("selection.set", { range: "A1" }).catch(() => {});
      } else if (tplId === "tasks") {
        await gc.run("range.setValues", {
          range: "A1:D5",
          values: [
            ["Tarea", "Prioridad", "Estado", "Fecha límite"],
            ["Revisar informe", "Alta", "En curso", stamp],
            ["Preparar presentación", "Media", "Pendiente", stamp],
            ["Enviar presupuesto", "Alta", "Listo", stamp],
            ["", "Completadas", '=COUNTIF(C2:C4,"Listo")', ""],
          ],
        }).catch(() => {});
        await gc.run("selection.set", { range: "A1:D1" }).catch(() => {});
        await gc.run("home.bold").catch(() => {});
        await gc.run("selection.set", { range: "A1:D5" }).catch(() => {});
        await gc.run("home.autofitColumnWidth").catch(() => {});
        await gc.run("selection.set", { range: "A1" }).catch(() => {});
      } else if (tplId === "schedule") {
        await gc.run("range.setValues", {
          range: "A1:F6",
          values: [
            ["Hora", "Lunes", "Martes", "Miércoles", "Jueves", "Viernes"],
            ["08:00", "", "", "", "", ""],
            ["10:00", "", "", "", "", ""],
            ["12:00", "", "", "", "", ""],
            ["15:00", "", "", "", "", ""],
            ["17:00", "", "", "", "", ""],
          ],
        }).catch(() => {});
        await gc.run("selection.set", { range: "A1:F1" }).catch(() => {});
        await gc.run("home.bold").catch(() => {});
        await gc.run("selection.set", { range: "A1:F6" }).catch(() => {});
        await gc.run("home.autofitColumnWidth").catch(() => {});
        await gc.run("selection.set", { range: "A1" }).catch(() => {});
      }

      const meta = {
        id: store.newId(),
        app: "grid",
        nombre: titles[tplId] || "Libro.xlsx",
        tamano: 0,
        abierto: Date.now(),
        modificado: Date.now(),
        hojas: 1,
      };
      S.doc = { meta, name: meta.nombre, dirty: true, dirtySince: Date.now(), activeCell: "A1", range: "A1", sheets: ["Hoja1"], activeSheet: "Hoja1" };
      await prepareEditorView();
      await saveActive().catch(() => {});
      busy(false);
      renderTop();
      renderBottom();
      await syncActiveCell();
    } catch (e) {
      busy(false);
      toast("No se pudo crear el libro: " + e.message, 4500);
      S.doc = null;
      showHome();
    } finally {
      S.opening = false;
    }
  }

  async function openDocument(meta, blob) {
    if (!S.ready) { toast("GridCraft todavía está iniciando"); return; }
    S.opening = true;
    showEditor();
    renderTop(meta.nombre);
    busy(true, "Abriendo " + meta.nombre + "…");
    try {
      const bytes = await blob.arrayBuffer();
      canvas.dispatchEvent(new CustomEvent("gridcraft:open", { detail: { name: meta.nombre, bytes } }));
      await sleep(320);
      S.doc = { meta, name: meta.nombre, dirty: false, dirtySince: 0, activeCell: "A1", range: "A1", sheets: ["Hoja1"], activeSheet: "Hoja1" };
      await prepareEditorView();
      busy(false);
      renderTop();
      renderBottom();
      await syncActiveCell();
      if (S.doc.sheets.length) {
        meta.hojas = S.doc.sheets.length;
        store.putMeta(meta).catch(() => {});
      }
    } catch (e) {
      busy(false);
      toast("No se pudo abrir: " + e.message, 4500);
      S.doc = null;
      showHome();
    } finally {
      S.opening = false;
    }
  }

  async function prepareEditorView() {
    await gc.setUi({ ribbonCollapsed: true, formulaBar: false, animations: false }).catch(() => {});
    await gc.run("view.statusBar", { on: false }).catch(() => {});
    await gc.run("view.showSheetTabs", { on: false }).catch(() => {});
    await gc.run("view.zoom", { percent: 100 }).catch(() => {});
    S.zoomPct = 100;
    await sleep(110);
    applyCanvasCrop(await gc.ui().catch(() => null));
    await sleep(60);
    applyCanvasCrop(await gc.ui().catch(() => null));
    updateZoomPill();
    window.ocSyncWakeLock(prefs.pantallaActiva);
  }

  async function goHome() {
    if (S.doc && S.doc.dirty) {
      busy(true, "Guardando libro…");
      await saveActive().catch(() => {});
      busy(false);
    }
    closeDrawer();
    S.doc = null;
    showHome();
  }

  function showEditor() {
    body.classList.add("oc-in-editor");
    syncThemeBar();
    renderBottom();
  }

  async function showHome() {
    body.classList.remove("oc-in-editor", "oc-desktop", "oc-modal");
    S.desktop = false;
    S.modal = false;
    syncThemeBar();
    window.ocSyncWakeLock(false);
    if (window.ocUpdateReady) { location.reload(); return; }
    await reloadLibrary();
    renderShell();
  }

  function busy(on, text) {
    $("oc-busy").hidden = !on;
    if (text) $("oc-busy-text").textContent = text;
  }

  // ---------- Barra superior (Excel iPhone) ----------
  function renderTop(fallbackName) {
    const title = S.doc ? S.doc.meta.nombre : fallbackName || "GridCraft";
    const sub = S.doc ? (S.doc.dirty ? "Editando…" : "Guardado en iPhone") : "";
    $("oc-top").replaceChildren(
      h("button", { class: "oc-icon-btn", type: "button", "aria-label": "Volver", icon: "back", onclick: goHome }),
      h("button", { class: "oc-title", type: "button", onclick: () => { if (S.doc) renameSheet(S.doc.meta); } }, h("b", { text: title }), h("small", { text: sub })),
      h("button", { class: "oc-icon-btn", type: "button", "aria-label": "Deshacer", icon: "undo", onclick: () => execGrid("edit.undo") }),
      h("button", { class: "oc-icon-btn", type: "button", "aria-label": "Rehacer", icon: "redo", onclick: () => execGrid("edit.redo") }),
      h("button", { class: "oc-icon-btn", type: "button", "aria-label": "Buscar", icon: "search", onclick: findSheet }),
      h("button", { class: "oc-icon-btn", type: "button", "aria-label": "Compartir", icon: "share", onclick: shareMenu }),
      h("button", { class: "oc-icon-btn", type: "button", "aria-label": "Más", icon: "more", onclick: moreMenu })
    );
  }

  function findSheet() {
    const inp = h("input", { class: "oc-field", type: "search", placeholder: "Buscar texto o número en la hoja…" });
    sheet({
      title: "Buscar en el libro",
      build: (close) => {
        const go = async () => {
          const q = inp.value.trim();
          if (!q) return;
          close();
          await gc.run("edit.find", { query: q }).catch(() => gc.run("ui.dialog", { name: "find" }).catch(() => toast("Sin coincidencias")));
          await syncActiveCell();
        };
        inp.addEventListener("keydown", (e) => { if (e.key === "Enter") go(); });
        return [inp, h("button", { class: "oc-primary", type: "button", text: "Buscar", onclick: go })];
      },
    });
    inp.focus();
  }

  // ---------- Barra rápida y Cinta móvil Excel iPhone ----------
  function renderBottom() {
    const items = [
      { icon: "keyboard", label: "Editar celda", fn: () => { const inp = $("oc-fx-input"); if (inp) { inp.focus(); try { inp.select(); } catch (e) {} } } },
      { icon: "bold", label: "Negrita", fn: () => execGrid("home.bold") },
      { icon: "italic", label: "Cursiva", fn: () => execGrid("home.italic") },
      { icon: "underline", label: "Subrayado", fn: () => execGrid("home.underline") },
      { icon: "fillColor", label: "Relleno", fn: fillColorPicker },
      { icon: "textColor", label: "Color texto", fn: fontColorPicker },
      { icon: "borders", label: "Bordes", fn: bordersSheet },
      { icon: "merge", label: "Combinar", fn: () => execGrid("home.mergeCenter") },
      { icon: "dollar", label: "Moneda", fn: () => execGrid("home.accounting") },
      { icon: "percent", label: "Porcentaje", fn: () => execGrid("home.percent") },
      { icon: "sigma", label: "Autosuma", fn: () => execGrid("formulas.autoSum") },
    ];
    $("oc-bottom").replaceChildren(
      h(
        "div",
        { class: "oc-quickbar" },
        h("div", { class: "oc-quick-scroll" }, items.map((it) => h("button", { class: "oc-qbtn", type: "button", "aria-label": it.label, icon: it.icon, onclick: it.fn }))),
        h("button", {
          class: "oc-drawer-toggle",
          type: "button",
          "aria-label": S.drawerOpen ? "Cerrar cinta" : "Abrir cinta de Excel",
          icon: S.drawerOpen ? "chevDown" : "chevUp",
          onclick: toggleDrawer,
        })
      )
    );
  }

  function toggleDrawer() {
    S.drawerOpen = !S.drawerOpen;
    renderDrawer();
    renderBottom();
    setTimeout(() => { gc.ui().then(applyCanvasCrop).catch(() => {}); }, 80);
  }
  function closeDrawer() {
    if (!S.drawerOpen) return;
    S.drawerOpen = false;
    $("oc-ribbon-drawer").hidden = true;
    renderBottom();
    setTimeout(() => { gc.ui().then(applyCanvasCrop).catch(() => {}); }, 80);
  }

  function renderDrawer() {
    const drawer = $("oc-ribbon-drawer");
    drawer.hidden = !S.drawerOpen;
    if (!S.drawerOpen) return;
    const tabs = [
      ["inicio", "Inicio"],
      ["insertar", "Insertar"],
      ["formulas", "Fórmulas"],
      ["datos", "Datos"],
      ["vista", "Vista"],
    ];
    drawer.replaceChildren(
      h(
        "div",
        { class: "oc-ribbon-tabs", role: "tablist" },
        tabs.map(([id, label]) => h("button", { class: "oc-rtab", type: "button", role: "tab", "aria-selected": String(S.drawerTab === id), text: label, onclick: () => { S.drawerTab = id; renderDrawer(); } })),
        h("button", { class: "oc-ribbon-close", type: "button", "aria-label": "Cerrar", icon: "close", onclick: closeDrawer })
      ),
      h("div", { class: "oc-ribbon-body" }, buildDrawerBody(S.drawerTab))
    );
  }

  function buildDrawerBody(tab) {
    if (tab === "inicio") {
      return [
        h("div", { class: "oc-rgroup" },
          h("div", { class: "oc-rgroup-title", text: "Fuente y colores" }),
          h("div", { class: "oc-rrow" },
            h("button", { class: "oc-rbtn", type: "button", icon: "bold", onclick: () => execGrid("home.bold") }),
            h("button", { class: "oc-rbtn", type: "button", icon: "italic", onclick: () => execGrid("home.italic") }),
            h("button", { class: "oc-rbtn", type: "button", icon: "underline", onclick: () => execGrid("home.underline") }),
            h("button", { class: "oc-rbtn", type: "button", icon: "strike", onclick: () => execGrid("home.strikethrough") }),
            h("button", { class: "oc-rbtn", type: "button", text: "A-", onclick: () => execGrid("home.decreaseFontSize") }),
            h("button", { class: "oc-rbtn", type: "button", text: "A+", onclick: () => execGrid("home.increaseFontSize") }),
            h("button", { class: "oc-rbtn", type: "button", icon: "fillColor", text: "Relleno", onclick: fillColorPicker }),
            h("button", { class: "oc-rbtn", type: "button", icon: "textColor", text: "Texto", onclick: fontColorPicker }),
            h("button", { class: "oc-rbtn", type: "button", icon: "borders", text: "Bordes", onclick: bordersSheet })
          )
        ),
        h("div", { class: "oc-rgroup" },
          h("div", { class: "oc-rgroup-title", text: "Alineación y celdas" }),
          h("div", { class: "oc-rrow" },
            h("button", { class: "oc-rbtn", type: "button", icon: "alignLeft", onclick: () => execGrid("home.alignLeft") }),
            h("button", { class: "oc-rbtn", type: "button", icon: "alignCenter", onclick: () => execGrid("home.alignCenter") }),
            h("button", { class: "oc-rbtn", type: "button", icon: "alignRight", onclick: () => execGrid("home.alignRight") }),
            h("button", { class: "oc-rbtn", type: "button", icon: "wrap", text: "Ajustar texto", onclick: () => execGrid("home.wrapText") }),
            h("button", { class: "oc-rbtn", type: "button", icon: "merge", text: "Combinar", onclick: () => execGrid("home.mergeCenter") }),
            h("button", { class: "oc-rbtn", type: "button", text: "Autoajustar col.", onclick: () => execGrid("home.autofitColumnWidth") })
          )
        ),
        h("div", { class: "oc-rgroup" },
          h("div", { class: "oc-rgroup-title", text: "Formato de número" }),
          h("div", { class: "oc-rrow" },
            h("button", { class: "oc-rbtn", type: "button", text: "General", onclick: () => execGrid("home.numberFormat", { format: "General" }) }),
            h("button", { class: "oc-rbtn", type: "button", text: "$ Moneda", onclick: () => execGrid("home.accounting") }),
            h("button", { class: "oc-rbtn", type: "button", text: "%", onclick: () => execGrid("home.percent") }),
            h("button", { class: "oc-rbtn", type: "button", text: "0,000", onclick: () => execGrid("home.comma") }),
            h("button", { class: "oc-rbtn", type: "button", text: "+.0", onclick: () => execGrid("home.increaseDecimal") }),
            h("button", { class: "oc-rbtn", type: "button", text: "-.0", onclick: () => execGrid("home.decreaseDecimal") }),
            h("button", { class: "oc-rbtn", type: "button", text: "Fecha", onclick: () => execGrid("home.numberFormat", { format: "dd/mm/yyyy" }) })
          )
        ),
      ];
    }
    if (tab === "insertar") {
      return [
        h("div", { class: "oc-rgrid-3" },
          rtile("chart", "Gráfico", insertChartSheet),
          rtile("table", "Tabla", () => execGrid("insert.table", { header: true })),
          rtile("plus", "Nueva hoja", async () => { await execGrid("home.insertSheet"); await syncActiveCell(); }),
          rtile("comment", "Nota", addNoteSheet),
          rtile("link", "Vínculo", insertLinkSheet),
          rtile("shape", "Forma", () => execGrid("insert.shape", { preset: "rect" }))
        ),
      ];
    }
    if (tab === "formulas") {
      return [
        h("div", { class: "oc-rgrid-2" },
          rtile("sigma", "Autosuma (SUM)", () => execGrid("formulas.autoSum")),
          rtile("fx", "Catálogo fx…", () => { closeDrawer(); functionCatalogSheet(); }),
          rtile("eye", "Mostrar fórmulas", () => execGrid("formulas.showFormulas")),
          rtile("refresh", "Calcular ahora", () => execGrid("formulas.calculateNow"))
        ),
      ];
    }
    if (tab === "datos") {
      return [
        h("div", { class: "oc-rgrid-2" },
          rtile("sortAsc", "Ordenar A → Z", () => execGrid("data.sortAscending")),
          rtile("sortDesc", "Ordenar Z → A", () => execGrid("data.sortDescending")),
          rtile("filter", "Activar filtro", () => execGrid("data.filter")),
          rtile("clearFmt", "Quitar filtros", () => execGrid("data.clearFilter")),
          rtile("trash", "Quitar duplicados", () => execGrid("data.removeDuplicates"))
        ),
      ];
    }
    return [
      h("div", { class: "oc-rgrid-2" },
        rtile("freeze", "Inmovilizar fila 1", () => execGrid("view.freezeTopRow")),
        rtile("freeze", "Inmovilizar col. A", () => execGrid("view.freezeFirstColumn")),
        rtile("freeze", "Inmovilizar paneles", () => execGrid("view.freezePanes")),
        rtile("close", "Movilizar paneles", () => execGrid("view.unfreezePanes")),
        rtile("table", "Líneas cuadrícula", () => execGrid("view.gridlines")),
        rtile("zoom", "Zoom 100%", () => { S.zoomPct = 100; execGrid("view.zoom100"); updateZoomPill(); }),
        rtile("desktop", "Modo PC completo", () => { closeDrawer(); setDesktop(true); })
      ),
    ];
  }

  function rtile(ic, label, onclick) {
    return h("button", { class: "oc-rtile", type: "button", icon: ic, onclick }, h("span", { text: label }));
  }

  async function execGrid(cmd, params) {
    try {
      const res = await gc.run(cmd, params || {});
      setTimeout(() => syncActiveCell(), 60);
      return res;
    } catch (e) {
      toast("No disponible: " + (e.message || e), 3000);
      return null;
    }
  }

  // ---------- Selectores Excel ----------
  function fillColorPicker() {
    sheet({
      title: "Color de relleno",
      build: (close) => [
        colorGrid([
          ["none", "Sin relleno", ""],
          ["#FFF3BF", "Amarillo suave", "#FFF3BF"],
          ["#D3F9D8", "Verde suave", "#D3F9D8"],
          ["#D0EBFF", "Azul suave", "#D0EBFF"],
          ["#FFE3E3", "Rojo suave", "#FFE3E3"],
          ["#E9ECEF", "Gris suave", "#E9ECEF"],
          ["#107C41", "Verde Excel", "#107C41"],
          ["#185ABD", "Azul", "#185ABD"],
          ["#F59F00", "Ámbar", "#F59F00"],
          ["#C43E1C", "Naranja", "#C43E1C"],
          ["#343A40", "Carbón", "#343A40"],
          ["#FFFFFF", "Blanco", "#FFFFFF"],
        ], (v) => { close(); execGrid("home.fillColor", { color: v === "none" ? "" : v }); }),
      ],
    });
  }

  function fontColorPicker() {
    sheet({
      title: "Color de fuente",
      build: (close) => [
        colorGrid([
          ["auto", "Automático", ""],
          ["#000000", "Negro", "#000000"],
          ["#FFFFFF", "Blanco", "#FFFFFF"],
          ["#107C41", "Verde Excel", "#107C41"],
          ["#D92D20", "Rojo", "#D92D20"],
          ["#185ABD", "Azul", "#185ABD"],
          ["#E8590C", "Naranja", "#E8590C"],
          ["#595959", "Gris", "#595959"],
        ], (v) => { close(); execGrid("home.fontColor", { color: v === "auto" ? "" : v }); }),
      ],
    });
  }

  function bordersSheet() {
    const opts = [
      ["all", "Todos los bordes"],
      ["outside", "Borde exterior"],
      ["thick", "Borde exterior grueso"],
      ["bottom", "Borde inferior"],
      ["topBottom", "Superior e inferior"],
      ["none", "Sin borde"],
    ];
    sheet({
      title: "Bordes de celda",
      build: (close) => [list(opts.map(([preset, t]) => ({ i: "borders", t, run: () => { close(); execGrid("home.borders", { preset }); } })))],
    });
  }

  function insertChartSheet() {
    const types = [
      ["column", "Columnas"],
      ["bar", "Barras"],
      ["line", "Líneas"],
      ["pie", "Circular"],
      ["area", "Área"],
      ["scatter", "Dispersión"],
    ];
    sheet({
      title: "Insertar gráfico",
      text: "Usa el rango seleccionado (" + (S.doc ? S.doc.range : "A1") + ")",
      build: (close) => [list(types.map(([type, t]) => ({ i: "chart", t, run: () => { close(); execGrid("insert.chart", { type, range: S.doc ? S.doc.range : undefined }); } })))],
    });
  }

  function addNoteSheet() {
    const inp = h("input", { class: "oc-field", type: "text", placeholder: "Nota para " + (S.doc ? S.doc.activeCell : "A1") });
    sheet({
      title: "Insertar nota",
      build: (close) => [
        inp,
        h("button", { class: "oc-primary", type: "button", text: "Guardar nota", onclick: () => { const t = inp.value.trim(); if (!t) return; close(); execGrid("review.newNote", { text: t }); } }),
      ],
    });
    inp.focus();
  }

  function insertLinkSheet() {
    const inp = h("input", { class: "oc-field", type: "url", placeholder: "https://ejemplo.com" });
    sheet({
      title: "Insertar vínculo",
      build: (close) => [
        inp,
        h("button", { class: "oc-primary", type: "button", text: "Insertar", onclick: () => { const url = inp.value.trim(); if (!url) return; close(); execGrid("insert.link", { url }); } }),
      ],
    });
    inp.focus();
  }

  function zoomSheet() {
    sheet({
      title: "Zoom de hoja",
      build: (close) => [
        h("div", { class: "oc-btn-row" },
          [75, 100, 125, 150].map((pct) =>
            h("button", { class: "oc-secondary", type: "button", text: pct + "%", onclick: () => { close(); S.zoomPct = pct; execGrid("view.zoom", { percent: pct }); updateZoomPill(); } })
          )
        ),
      ],
    });
  }

  function shareMenu() {
    if (!S.doc) return;
    sheet({
      title: "Compartir o exportar",
      text: S.doc.meta.nombre,
      build: (close) => [
        list([
          { i: "share", t: "Compartir como Excel (.xlsx)", small: "Guardar en Archivos", run: async () => { close(); busy(true, "Preparando .xlsx…"); try { const f = await currentFile(); busy(false); if (f) shareFile(f); } catch (e) { busy(false); toast(e.message); } } },
          { i: "pdf", t: "Exportar hoja a PDF", small: ".pdf", run: () => { close(); execGrid("file.exportPdf"); toast("Generando PDF…"); } },
          { i: "print", t: "Imprimir (AirPrint)", run: () => { close(); execGrid("file.print"); } },
        ]),
      ],
    });
  }

  function moreMenu() {
    if (!S.doc) return;
    sheet({
      title: S.doc.meta.nombre,
      build: (close) => [
        list([
          { i: "check", t: "Guardar ahora", run: async () => { close(); try { await saveActive(); toast("Guardado en este iPhone"); } catch (e) { toast(e.message); } } },
          { i: "star", t: S.doc.meta.destacado ? "Quitar de destacados" : "Destacar", run: () => { close(); toggleStar(S.doc.meta); } },
          { i: "rename", t: "Renombrar", run: () => { close(); renameSheet(S.doc.meta); } },
          { i: "fx", t: "Ir a celda…", run: () => { close(); goToCellSheet(); } },
          { i: "pdf", t: "Exportar a PDF", run: () => { close(); execGrid("file.exportPdf"); } },
          { i: "desktop", t: "Interfaz completa (PC)", small: "Cinta completa", run: () => { close(); setDesktop(true); } },
        ]),
      ],
    });
  }

  function settingsSheet() {
    sheet({
      title: "Ajustes de GridCraft",
      build: (close) => [
        h("div", { class: "oc-label", text: "Apariencia" }),
        seg([["auto", "Sistema"], ["claro", "Claro"], ["oscuro", "Oscuro"]], () => prefs.tema, (v) => {
          prefs.tema = v;
          savePrefs();
          applyTheme();
          if (S.ready) gc.run("view.theme", { mode: v === "oscuro" ? "dark" : v === "claro" ? "light" : "system" }).catch(() => {});
        }),
        h("div", { class: "oc-label", text: "Fluidez y nitidez (iPhone)" }),
        seg([["rapida", "Rápida"], ["fluida", "Fluida (2×)"], ["maxima", "Máxima"]], () => prefs.calidad, (v) => {
          prefs.calidad = v;
          savePrefs();
          toast("Se aplicará la nueva fluidez al volver a abrir la app");
        }),
        h("div", { class: "oc-list" }, swRow("Mantener pantalla encendida", () => prefs.pantallaActiva, (v) => { prefs.pantallaActiva = v; savePrefs(); })),
        list([
          { i: "refresh", t: "Buscar actualizaciones", run: () => { close(); window.ocCheckUpdate(); } },
          { i: "apps", t: "Ir al portal OfficeCraft", small: "Word · PowerPoint · Excel", run: () => { location.href = "../"; } },
        ]),
        h("p", { text: "GridCraft para iPhone · Motor de código abierto (github.com/storytold/gridcraft) " + (S.appVersion ? "v" + S.appVersion : "") + ". Tus libros no salen del teléfono." }),
      ],
    });
  }

  let backBtn = null;
  async function setDesktop(on) {
    S.desktop = on;
    body.classList.toggle("oc-desktop", on);
    if (on) {
      closeDrawer();
      await gc.setUi({ ribbonCollapsed: false, formulaBar: true }).catch(() => {});
      await gc.run("view.statusBar", { on: true }).catch(() => {});
      await gc.run("view.showSheetTabs", { on: true }).catch(() => {});
      await gc.run("view.interfaceScale", { percent: 80 }).catch(() => {});
      applyCanvasCrop(null);
      backBtn = backBtn || h("button", { class: "oc-back-mobile", type: "button", text: "Volver a vista iPhone", onclick: () => setDesktop(false) });
      body.append(backBtn);
    } else {
      if (backBtn) backBtn.remove();
      await gc.setUi({ ribbonCollapsed: true, formulaBar: false }).catch(() => {});
      await gc.run("view.statusBar", { on: false }).catch(() => {});
      await gc.run("view.showSheetTabs", { on: false }).catch(() => {});
      await gc.run("view.interfaceScale", { percent: 100 }).catch(() => {});
      await sleep(120);
      applyCanvasCrop(await gc.ui().catch(() => null));
    }
  }

  window.ocOnUpdate = () => {
    if (!body.classList.contains("oc-in-editor")) location.reload();
    else toast("Hay una versión nueva de GridCraft: se aplicará al volver a Inicio.", 5500);
  };

  window.gridcraftMobile = { openCreateSheet, createFromTemplate, toggleDrawer, setDesktop, goHome, state: () => S };
})();
