// DeckCraft para iPhone: interfaz móvil fiel a Microsoft PowerPoint para iOS sobre el motor
// oficial de DeckCraft (data-host="page" + window.deckcraft) con gestos a 120 fps sin tirones.
(function () {
  "use strict";

  const $ = (id) => document.getElementById(id);
  const { icon, h, sheet, list, seg, swRow, colorGrid, sizeText, describe, makeSearchBox, sleep } = window.ocUI;
  const store = window.ocStore;
  const toast = (t, ms) => window.ocToast(t, ms);
  const canvas = $("deckcraft_canvas");
  const body = document.body;

  // ---------- Preferencias ----------
  const PREFS_KEY = "officecraft-deck-prefs";
  const prefs = Object.assign(
    {
      tema: "auto",
      autor: "",
      vistaArchivos: "lista",
      orden: "reciente",
      pantallaActiva: true,
    },
    readPrefs()
  );
  function readPrefs() {
    try { return JSON.parse(localStorage.getItem(PREFS_KEY) || "{}"); } catch (e) { return {}; }
  }
  function savePrefs() {
    try { localStorage.setItem(PREFS_KEY, JSON.stringify(prefs)); } catch (e) {}
  }
  function applyTheme() {
    if (prefs.tema === "auto") document.documentElement.removeAttribute("data-theme");
    else document.documentElement.setAttribute("data-theme", prefs.tema === "oscuro" ? "dark" : "light");
  }
  applyTheme();

  // ---------- Puente DeckCraft API ----------
  const dc = {
    ok: () => !!(window.deckcraft && typeof window.deckcraft.request === "function"),
    ui: () => window.deckcraft.uiInspect(),
    doc: () => window.deckcraft.docInspect(),
    run: (command, params) => window.deckcraft.execute(command, params || {}),
    set: (params) => window.deckcraft.request("ui.set", params || {}),
    tool: (tool, preset) => window.deckcraft.request("ui.tool.select", preset ? { tool, preset } : { tool }),
    key: (key) => window.deckcraft.request("ui.key", { key }),
    renderSlide: (slide, scale) => window.deckcraft.request("ui.render", { slide, scale: scale || 0.24 }),
    showStart: (from) => window.deckcraft.request("show.start", from != null ? { from } : {}),
    showEnd: () => window.deckcraft.request("show.end", {}),
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
    doc: null,              // { meta, name, dirty, dirtySince, slides, slideIdx, zoomFit }
    opening: false,
    desktop: false,
    modal: false,
    showing: false,
    notesOpen: false,
    drawerOpen: false,
    drawerTab: "inicio",    // "inicio" | "insertar" | "dibujar" | "diseno" | "transiciones" | "presentacion"
    activeTool: "select",
    ratio: 1,
    canvasScale: 1,
    fitScale: 1,
    crop: { l: 0, t: 0, w: 0, h: 0 },
    thumbs: new Map(),      // slideIdx -> dataURL
    thumbGen: 0,
    polling: false,
    saveWaiters: [],
    lastUi: null,
  };

  // ---------- Eventos del modo anfitrión (data-host="page") ----------
  canvas.addEventListener("deckcraft:ready", (e) => {
    if (e.detail && e.detail.appVersion) S.appVersion = e.detail.appVersion;
  });
  canvas.addEventListener("deckcraft:dirty", (e) => {
    if (!S.doc) return;
    setDirty(!!(e.detail && e.detail.dirty));
  });
  canvas.addEventListener("deckcraft:failed", (e) => {
    window.ocShowError("DeckCraft no pudo iniciar: " + ((e.detail && e.detail.error) || "sin WebGL2/WebGPU"));
  });
  canvas.addEventListener("deckcraft:save", (e) => {
    onEngineSave(e.detail || {}).catch((err) => toast("Error al guardar: " + err.message, 4500));
  });

  function setDirty(dirty) {
    if (!S.doc || S.doc.dirty === dirty) return;
    S.doc.dirty = dirty;
    S.doc.dirtySince = dirty ? Date.now() : 0;
    renderTop();
  }

  async function onEngineSave(d) {
    const rawName = d.name || (S.doc ? S.doc.meta.nombre : "Presentacion.pptx");
    const isDeck = /\.(pptx|deckcraft|potx|ppsx)$/i.test(rawName);
    const bytes = d.bytes instanceof ArrayBuffer ? d.bytes : (d.bytes && d.bytes.buffer) || new ArrayBuffer(0);
    const file = new File([bytes], rawName, { type: window.ocMimeFor(rawName) });

    const waiters = S.saveWaiters.splice(0, S.saveWaiters.length);
    if (isDeck && S.doc) {
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

  function saveActive(asPptx) {
    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => {
        S.saveWaiters = S.saveWaiters.filter((w) => w !== waiter);
        reject(new Error("DeckCraft no respondió al guardar"));
      }, 20000);
      const waiter = {
        resolve: (f) => { clearTimeout(timer); resolve(f); },
        reject: (e) => { clearTimeout(timer); reject(e); },
      };
      S.saveWaiters.push(waiter);
      const targetName = S.doc ? S.doc.meta.nombre : "Presentacion.pptx";
      const cmd = asPptx ? "file.saveAs" : "file.save";
      const params = asPptx ? { path: targetName.replace(/\.[^.]+$/, "") + ".pptx" } : {};
      dc.run(cmd, params).catch((e) => {
        // Fallback: si file.save necesita nombre la primera vez, intenta saveAs .pptx
        if (!asPptx) {
          dc.run("file.saveAs", { path: targetName.replace(/\.[^.]+$/, "") + ".pptx" }).catch((e2) => {
            clearTimeout(timer);
            S.saveWaiters = S.saveWaiters.filter((w) => w !== waiter);
            reject(e2);
          });
        } else {
          clearTimeout(timer);
          S.saveWaiters = S.saveWaiters.filter((w) => w !== waiter);
          reject(e);
        }
      });
    });
  }

  async function currentFile() {
    if (!S.doc) return null;
    if (S.doc.dirty) return saveActive(false);
    const b = await store.blob(S.doc.meta.id);
    if (b) return new File([b], S.doc.meta.nombre, { type: window.ocMimeFor(S.doc.meta.nombre) });
    return saveActive(true);
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
      window.ocShowError("Error al iniciar DeckCraft: " + (e && e.message ? e.message : e));
    });
  });

  async function boot() {
    for (let i = 0; i < 400 && !dc.ok(); i++) await sleep(50);
    if (!dc.ok()) throw new Error("No se inicializó window.deckcraft");
    for (let i = 0; i < 80; i++) {
      try { S.lastUi = await dc.ui(); break; } catch (e) { await sleep(150); }
    }
    S.ready = true;

    const isDark = prefs.tema === "oscuro" || (prefs.tema === "auto" && matchMedia("(prefers-color-scheme: dark)").matches);
    await dc.set({ ribbonCollapsed: true, notes: false, pane: null, view: "normal", dark: isDark }).catch(() => {});

    gestures = window.ocGestures.attach(canvas, {
      mode: () => {
        if (S.showing) return "ver";
        if (S.desktop) return "raton-zoom";
        if (S.modal) return "raton";
        if (S.activeTool !== "select") return "herramienta";
        return "ver";
      },
      ratio: () => S.ratio,
      axisLock: "free",
      swipePages: () => {
        if (S.showing) return true;
        if (!S.doc || S.desktop || S.modal || S.activeTool !== "select") return false;
        return S.canvasScale <= S.fitScale * 1.08;
      },
      onSwipePage: (dir) => {
        if (S.showing) {
          dc.key(dir > 0 ? "Right" : "Left").catch(() => {});
          return;
        }
        stepSlide(dir);
      },
      onPinchCommit: async ({ scale }) => {
        if (!S.doc || S.showing || S.desktop || S.modal) return;
        const cur = S.canvasScale || S.fitScale || 0.5;
        const next = Math.max(0.15, Math.min(4.0, cur * scale));
        if (next <= S.fitScale * 1.03) {
          await fitSlide();
        } else {
          await dc.set({ zoom: Number(next.toFixed(3)) }).catch(() => {});
          S.canvasScale = next;
        }
        updateSlidePill();
      },
      onDoubleTap: () => {
        if (S.showing) return;
        toggleDoubleTapZoom();
      },
      onTap: () => {
        if (S.showing) {
          dc.key("Right").catch(() => {});
          return;
        }
        if (S.drawerOpen) closeDrawer();
      },
      onLongPress: (x, y) => { if (!S.showing) showPressRing(x, y); },
    });

    initStaticControls();
    await reloadLibrary();
    renderShell();
    window.ocHideSplash();

    const qs = new URLSearchParams(location.search);
    const openId = qs.get("open");
    if (openId) {
      const m = S.files.find((x) => x.id === openId);
      if (m) openStoredDeck(m);
    } else if (qs.get("new") === "1") {
      createFromTemplate("blank");
    }

    setInterval(tick, 450);
    document.addEventListener("visibilitychange", () => {
      window.ocSyncWakeLock(prefs.pantallaActiva && body.classList.contains("oc-in-editor"));
      if (document.visibilityState === "hidden" && S.doc && S.doc.dirty && !S.modal && !S.showing) {
        saveActive(false).catch(() => {});
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

  async function fitSlide() {
    await dc.run("view.fit").catch(() => dc.set({ zoom: null }).catch(() => {}));
    await sleep(90);
    const u = await dc.ui().catch(() => null);
    if (u && u.canvasScale) {
      S.canvasScale = u.canvasScale;
      S.fitScale = u.canvasScale;
    }
    updateSlidePill();
  }

  async function toggleDoubleTapZoom() {
    if (!S.doc || S.desktop || S.modal) return;
    if (S.canvasScale > S.fitScale * 1.2) {
      await fitSlide();
    } else {
      const target = Math.min(3.0, Number((S.fitScale * 1.9).toFixed(3)));
      await dc.set({ zoom: target }).catch(() => {});
      S.canvasScale = target;
      updateSlidePill();
    }
  }

  // ---------- Recorte exacto al lienzo de la diapositiva (0 % barras / panel izquierdo PC) ----------
  function applyCanvasCrop(u) {
    const docEl = $("oc-doc");
    if (!docEl) return;
    if (S.desktop || S.modal || S.showing || !u || !u.window || !u.canvasRect) {
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
    const [cx, cy, cw, ch] = u.canvasRect;
    if (!(ww > 100 && wh > 100 && cw > 50 && ch > 50)) return;
    const ratio = canvas.clientWidth / ww || 1;
    S.ratio = ratio;

    const chromeL = Math.max(0, Math.min(ww * 0.55, cx));
    const chromeT = Math.max(0, Math.min(wh * 0.55, cy));
    const chromeR = Math.max(0, Math.min(ww * 0.45, ww - (cx + cw)));
    const chromeB = Math.max(0, Math.min(wh * 0.35, wh - (cy + ch)));

    const vw = docEl.clientWidth;
    const vh = docEl.clientHeight;
    if (!(vw > 50 && vh > 50)) return;

    const left = -(chromeL * ratio);
    const top = -(chromeT * ratio);
    const width = vw + (chromeL + chromeR) * ratio;
    const height = vh + (chromeT + chromeB) * ratio;

    if (
      Math.abs(left - S.crop.l) > 1.5 ||
      Math.abs(top - S.crop.t) > 1.5 ||
      Math.abs(width - S.crop.w) > 1.5 ||
      Math.abs(height - S.crop.h) > 1.5
    ) {
      S.crop = { l: left, t: top, w: width, h: height };
      docEl.style.setProperty("--oc-crop-left", left.toFixed(1) + "px");
      docEl.style.setProperty("--oc-crop-top", top.toFixed(1) + "px");
      docEl.style.setProperty("--oc-canvas-w", width.toFixed(1) + "px");
      docEl.style.setProperty("--oc-canvas-h", height.toFixed(1) + "px");
    }
  }

  // ---------- Sondeo sin tirones ----------
  async function tick() {
    if (!S.ready || S.polling || document.visibilityState !== "visible") return;
    if (gestures && gestures.busy()) return;
    S.polling = true;
    try {
      const u = await dc.ui();
      S.lastUi = u;
      if (u.window && u.window[0]) S.ratio = canvas.clientWidth / u.window[0];
      const showing = !!u.showing;
      if (showing !== S.showing) {
        S.showing = showing;
        body.classList.toggle("oc-show", showing);
        $("oc-show-hud").hidden = !showing;
        applyCanvasCrop(showing ? null : u);
      }
      if (showing && u.showing) {
        const idx = (u.showing.slide != null ? u.showing.slide : 0) + 1;
        const tot = (S.doc && S.doc.slides) || 1;
        $("oc-show-count").textContent = idx + " / " + tot;
      }

      const modal = !!u.dialog;
      if (modal !== S.modal) {
        S.modal = modal;
        body.classList.toggle("oc-modal", modal);
        if (modal) closeDrawer();
      }
      if (!body.classList.contains("oc-in-editor")) return;
      applyCanvasCrop(u);
      if (S.opening || !S.doc || S.showing) return;

      if (u.canvasScale) S.canvasScale = u.canvasScale;
      const d = await dc.doc().catch(() => null);
      if (d) {
        const count = Array.isArray(d.slides) ? d.slides.length : (d.slideCount || S.doc.slides || 1);
        const cur = d.selection && typeof d.selection.slide === "number" ? d.selection.slide : S.doc.slideIdx;
        const changed = count !== S.doc.slides || cur !== S.doc.slideIdx;
        S.doc.slides = Math.max(1, count);
        S.doc.slideIdx = Math.max(0, Math.min(S.doc.slides - 1, cur));
        if (changed) {
          renderTop();
          renderFilmstrip();
          updateSlidePill();
          if (S.notesOpen) refreshNotesBox();
        }
      }

      if (S.doc.dirty && S.doc.dirtySince && Date.now() - S.doc.dirtySince > 25000 && !S.modal && !S.saveWaiters.length) {
        S.doc.dirtySince = Date.now();
        saveActive(false).catch(() => {});
      }
    } catch (e) {
    } finally {
      S.polling = false;
    }
  }

  function updateSlidePill() {
    const pill = $("oc-slide-pill");
    if (!pill) return;
    if (!S.doc || S.desktop || S.modal || S.showing) {
      pill.hidden = true;
      return;
    }
    pill.hidden = false;
    const pct = S.fitScale > 0 ? Math.round((S.canvasScale / S.fitScale) * 100) : 100;
    pill.textContent = (S.doc.slideIdx + 1) + " / " + S.doc.slides + (Math.abs(pct - 100) > 6 ? " · " + pct + "%" : "");
  }

  // ---------- Plantillas PowerPoint para iPhone ----------
  const TEMPLATES = [
    { id: "blank", name: "En blanco", bg: "#ffffff", fg: "#c43e1c" },
    { id: "pitch", name: "Pitch moderno", bg: "#1a2238", fg: "#f26c47" },
    { id: "academic", name: "Académica", bg: "#f8f5ef", fg: "#2b57db" },
    { id: "dark", name: "Oscura", bg: "#141418", fg: "#ffffff" },
    { id: "report", name: "Resumen ejecutivo", bg: "#eef4f0", fg: "#107c41" },
  ];

  // ---------- Shell Home ----------
  function initStaticControls() {
    const fab = $("oc-create-fab");
    fab.replaceChildren();
    fab.insertAdjacentHTML("afterbegin", icon("plus"));
    fab.addEventListener("click", openCreateSheet);
    $("oc-slide-pill").addEventListener("click", fitSlide);

    $("oc-show-prev").addEventListener("click", () => dc.key("Left").catch(() => {}));
    $("oc-show-next").addEventListener("click", () => dc.key("Right").catch(() => {}));
    $("oc-show-exit").addEventListener("click", () => dc.showEnd().catch(() => {}));

    $("oc-file").addEventListener("change", async () => {
      const file = $("oc-file").files && $("oc-file").files[0];
      if (!file) return;
      $("oc-file").value = "";
      const meta = {
        id: store.newId(),
        app: "deck",
        nombre: file.name || "Presentacion.pptx",
        tamano: file.size,
        abierto: Date.now(),
        modificado: file.lastModified || Date.now(),
      };
      try { await store.save(meta, file); } catch (e) {}
      openDocument(meta, file);
    });

    $("oc-img").addEventListener("change", async () => {
      const file = $("oc-img").files && $("oc-img").files[0];
      if (!file) return;
      $("oc-img").value = "";
      insertImageFromFile(file);
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
        h(
          "button",
          { class: "oc-tab-btn", type: "button", "aria-current": S.tab === tb.id ? "page" : null, onclick: () => switchTab(tb.id) },
          h("span", { icon: tb.i }),
          h("span", { text: tb.t })
        )
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
    try { S.files = await store.list("deck"); } catch (e) { S.files = []; }
    try { S.folders = await store.listFolders("deck"); } catch (e) { S.folders = []; }
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
          h("h1", { class: "oc-head-title", text: "DeckCraft" }),
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
        S.homeSearchOpen ? makeSearchBox("Buscar presentaciones…", S.query, (q) => { S.query = q; renderHomeBody(); }, true) : null,
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
        h("div", { class: "oc-section-head" }, h("h2", { text: "Nueva presentación" }), h("button", { class: "oc-link-btn", type: "button", text: "Importar", onclick: pickFile })),
        h(
          "div",
          { class: "oc-tpl-strip" },
          TEMPLATES.map((tpl) => {
            const prev = h("div", { class: "oc-tpl-preview oc-tpl-wide" });
            prev.style.background = tpl.bg;
            if (tpl.id === "blank") {
              prev.insertAdjacentHTML("afterbegin", icon("plus"));
            } else {
              const bar = h("div", { class: "oc-tpl-lines" }, h("i", { class: "oc-tpl-accent" }), h("i", { style: "width:55%;opacity:.45" }));
              bar.firstChild.style.background = tpl.fg;
              prev.append(bar);
            }
            return h("button", { class: "oc-tpl-card oc-tpl-wide", type: "button", onclick: () => createFromTemplate(tpl.id) }, prev, h("b", { text: tpl.name }));
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
          h("div", { icon: "slide" }),
          h("b", { text: q ? "Sin coincidencias" : "Tus presentaciones aparecerán aquí" }),
          h("div", { text: q ? "No encontramos «" + S.query + "»." : "Crea una presentación o abre un archivo .pptx desde Archivos / iCloud." }),
          !q ? h("button", { class: "oc-empty-btn", type: "button", text: "Nueva presentación", onclick: () => createFromTemplate("blank") }) : null
        )
      );
    }
    el.replaceChildren(...nodes);
  }

  function renderStarredTab() {
    $("oc-app-head").replaceChildren(h("div", { class: "oc-head-row" }, h("h1", { class: "oc-head-title", text: "Destacados" }), h("button", { class: "oc-head-btn", type: "button", icon: "settings", onclick: settingsSheet })));
    const items = sorted(S.files.filter((m) => m.destacado));
    $("oc-tab-body").replaceChildren(
      items.length ? renderFileList(items) : h("div", { class: "oc-empty" }, h("div", { icon: "star" }), h("b", { text: "Sin presentaciones destacadas" }))
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
        !S.filesFolder ? h("button", { class: "oc-head-btn", type: "button", icon: "folderPlus", onclick: () => newFolderSheet(null) }) : null,
        h("button", { class: "oc-head-btn", type: "button", icon: prefs.vistaArchivos === "lista" ? "gridView" : "listView", onclick: () => { prefs.vistaArchivos = prefs.vistaArchivos === "lista" ? "cuadricula" : "lista"; savePrefs(); renderFilesTab(); } })
      ),
      makeSearchBox("Buscar presentaciones…", S.query, (q) => { S.query = q; renderFilesBody(); }, false)
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
          locRow("open", "#c43e1c", "Archivos / iCloud Drive", "Abrir .pptx, .potx, .ppsx o .deckcraft", pickFile),
          locRow("iphone", "#2b8a3e", "En este iPhone", S.files.length + " presentaciones · " + sizeText(total), () => {})
        )
      );
      if (S.folders.length) {
        nodes.push(
          h("div", { class: "oc-section-head" }, h("h2", { text: "Carpetas" }), h("button", { class: "oc-link-btn", type: "button", text: "+ Nueva", onclick: () => newFolderSheet(null) })),
          h("div", { class: "oc-loc-card" }, S.folders.map((f) => locRow("folder", "#5a6375", f.nombre, S.files.filter((m) => m.carpeta === f.id).length + " archivos", () => { S.filesFolder = f.id; renderFilesTab(); })))
        );
      }
      nodes.push(h("div", { class: "oc-section-head" }, h("h2", { text: "Todas las presentaciones" }), h("small", { text: String(S.files.length) })));
    }
    let items = S.files;
    if (S.filesFolder) items = items.filter((m) => m.carpeta === S.filesFolder);
    if (q) items = items.filter((m) => (m.nombre || "").toLowerCase().includes(q));
    items = sorted(items);
    nodes.push(items.length ? (prefs.vistaArchivos === "cuadricula" ? renderFileGrid(items) : renderFileList(items)) : h("div", { class: "oc-empty" }, h("div", { icon: "folder" }), h("b", { text: "Sin archivos" })));
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
        const thumb = h("span", { class: "oc-thumb oc-thumb-wide" });
        if (meta.miniatura) thumb.append(h("img", { src: meta.miniatura, alt: "" }));
        else { thumb.insertAdjacentHTML("afterbegin", icon("slide")); thumb.append(h("span", { class: "oc-thumb-tag", text: "PPTX" })); }
        return h(
          "li",
          { class: "oc-recent" },
          h("button", { class: "oc-recent-open", type: "button", onclick: () => openStoredDeck(meta) }, thumb, h("span", { class: "oc-recent-text" }, h("b", { text: meta.nombre }), h("small", { text: describe(meta, meta.diapositivas ? meta.diapositivas + " diap." : "") }))),
          h("button", { class: "oc-star-btn" + (meta.destacado ? " oc-starred" : ""), type: "button", [meta.destacado ? "iconFilled" : "icon"]: "star", onclick: () => toggleStar(meta) }),
          h("button", { class: "oc-recent-more", type: "button", icon: "more", onclick: () => fileMenu(meta) })
        );
      })
    );
  }

  function renderFileGrid(items) {
    return h(
      "div",
      { class: "oc-file-grid" },
      items.map((meta) => {
        const prev = h("div", { class: "oc-file-card-preview" });
        if (meta.miniatura) prev.append(h("img", { src: meta.miniatura, alt: "" }));
        else prev.insertAdjacentHTML("afterbegin", icon("slide"));
        return h(
          "div",
          { class: "oc-file-card" },
          h("button", { class: "oc-file-card-tap", type: "button", onclick: () => openStoredDeck(meta) }, prev, h("div", { class: "oc-file-card-meta" }, h("b", { text: meta.nombre }), h("small", { text: describe(meta) }))),
          h("div", { class: "oc-file-card-actions" },
            h("button", { class: "oc-card-dot" + (meta.destacado ? " oc-starred" : ""), type: "button", [meta.destacado ? "iconFilled" : "icon"]: "star", onclick: () => toggleStar(meta) }),
            h("button", { class: "oc-card-dot", type: "button", icon: "more", onclick: () => fileMenu(meta) })
          )
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
          { i: "open", t: "Abrir", run: () => { close(); openStoredDeck(meta); } },
          { i: "play", t: "Abrir y presentar", run: async () => { close(); await openStoredDeck(meta); setTimeout(() => startShow(0), 600); } },
          { i: "star", t: meta.destacado ? "Quitar de destacados" : "Destacar", run: () => { close(); toggleStar(meta); } },
          { i: "share", t: "Compartir / Guardar en Archivos", run: async () => { close(); const b = await store.blob(meta.id); if (b) shareFile(new File([b], meta.nombre, { type: window.ocMimeFor(meta.nombre) })); } },
          { i: "rename", t: "Renombrar", run: () => { close(); renameSheet(meta); } },
          { i: "copy", t: "Duplicar", run: () => { close(); duplicateDeck(meta); } },
        ]),
        list([{ i: "trash", t: "Eliminar", danger: true, run: () => { close(); confirmDelete(meta); } }]),
      ],
    });
  }

  function renameSheet(meta) {
    const base = (meta.nombre || "Presentacion.pptx").replace(/\.(pptx|deckcraft|potx|ppsx)$/i, "");
    const inp = h("input", { class: "oc-field", type: "text", value: base });
    sheet({
      title: "Renombrar presentación",
      build: (close) => {
        const go = async () => {
          const clean = inp.value.trim();
          if (!clean) return;
          meta.nombre = clean.replace(/\.[a-z0-9]+$/i, "") + ".pptx";
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

  async function duplicateDeck(meta) {
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
          const f = { id: store.newId(), app: "deck", nombre: n, creada: Date.now() };
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
      title: "¿Eliminar presentación?",
      text: "Se borrará «" + meta.nombre + "» de DeckCraft en este iPhone.",
      build: (close) => [list([{ i: "trash", t: "Eliminar", danger: true, run: async () => { close(); await store.remove(meta.id); await reloadLibrary(); renderShell(); } }])],
    });
  }

  function sortSheet() {
    sheet({
      title: "Ordenar presentaciones",
      build: (close) => [list([["reciente", "Reciente"], ["nombre", "Nombre"], ["tamano", "Tamaño"]].map(([v, t]) => ({ i: prefs.orden === v ? "check" : "sort", t, run: () => { prefs.orden = v; savePrefs(); close(); renderShell(); } })))],
    });
  }

  function openCreateSheet() {
    sheet({
      title: "Nueva presentación",
      build: (close) => [
        list([
          { i: "slidePlus", t: "Presentación en blanco", small: "16:9 .pptx", run: () => { close(); createFromTemplate("blank"); } },
          { i: "open", t: "Abrir desde Archivos / iCloud", small: ".pptx, .deckcraft", run: () => { close(); pickFile(); } },
          { i: "slide", t: "Pitch moderno", small: "3 diapositivas", run: () => { close(); createFromTemplate("pitch"); } },
          { i: "slide", t: "Presentación académica", small: "Plantilla", run: () => { close(); createFromTemplate("academic"); } },
          { i: "slide", t: "Minimalista oscura", small: "Plantilla", run: () => { close(); createFromTemplate("dark"); } },
        ]),
      ],
    });
  }

  function pickFile() {
    const i = $("oc-file");
    i.value = "";
    i.click();
  }

  // ---------- Abrir / Crear presentación ----------
  async function openStoredDeck(meta) {
    const blob = await store.blob(meta.id);
    if (!blob) { toast("Ese archivo ya no existe."); await store.remove(meta.id); await reloadLibrary(); renderShell(); return; }
    meta.abierto = Date.now();
    store.putMeta(meta).catch(() => {});
    await openDocument(meta, blob);
  }

  async function createFromTemplate(tplId) {
    if (!S.ready) { toast("DeckCraft todavía está iniciando"); return; }
    S.opening = true;
    showEditor();
    busy(true, "Creando presentación…");
    try {
      await dc.run("file.new", {});
      await sleep(120);
      if (tplId === "pitch") {
        await dc.run("slide.new", { layout: "title", at: 0, title: "Propuesta de proyecto", body: "Visión, estrategia y resultados clave" }).catch(() => {});
        await dc.run("slide.delete", { index: 1 }).catch(() => {});
        await dc.run("slide.new", { layout: "titleAndContent", title: "Problema y solución", body: "• Reto principal\n• Propuesta de valor\n• Resultados clave" }).catch(() => {});
        await dc.run("slide.new", { layout: "twoContent", title: "Plan y métricas" }).catch(() => {});
        await dc.run("design.background", { color: "#1A2238", all: true }).catch(() => {});
        for (let i = 0; i < 3; i++) {
          await dc.run("slide.go", { index: i }).catch(() => {});
          await dc.run("edit.selectAll", {}).catch(() => {});
          await dc.run("format.color", { color: "#FFFFFF" }).catch(() => {});
          await dc.run("edit.deselect", {}).catch(() => {});
        }
        await dc.run("slide.go", { index: 0 }).catch(() => {});
      } else if (tplId === "dark") {
        await dc.run("slide.new", { layout: "title", at: 0, title: "Presentación", body: "Diseño minimalista oscuro" }).catch(() => {});
        await dc.run("slide.delete", { index: 1 }).catch(() => {});
        await dc.run("design.background", { color: "#141418", all: true }).catch(() => {});
        await dc.run("edit.selectAll", {}).catch(() => {});
        await dc.run("format.color", { color: "#FFFFFF" }).catch(() => {});
        await dc.run("edit.deselect", {}).catch(() => {});
      } else if (tplId === "academic") {
        await dc.run("slide.new", { layout: "title", at: 0, title: "Tema de investigación", body: "Facultad · Curso académico" }).catch(() => {});
        await dc.run("slide.delete", { index: 1 }).catch(() => {});
        await dc.run("slide.new", { layout: "titleAndContent", title: "Introducción", body: "• Contexto\n• Objetivos\n• Metodología" }).catch(() => {});
        await dc.run("design.background", { color: "#FAF8F3", all: true }).catch(() => {});
        await dc.run("slide.go", { index: 0 }).catch(() => {});
      } else if (tplId === "report") {
        await dc.run("slide.new", { layout: "title", at: 0, title: "Resumen ejecutivo", body: "Resultados e indicadores del período" }).catch(() => {});
        await dc.run("slide.delete", { index: 1 }).catch(() => {});
        await dc.run("slide.new", { layout: "titleAndContent", title: "Indicadores clave", body: "• Logros del trimestre\n• Métricas principales\n• Próximos pasos" }).catch(() => {});
        await dc.run("design.background", { color: "#F2F7F4", all: true }).catch(() => {});
        await dc.run("slide.go", { index: 0 }).catch(() => {});
      }
      const d = await dc.doc().catch(() => null);
      const count = d && Array.isArray(d.slides) ? d.slides.length : 1;
      const stamp = new Date().toLocaleDateString("es", { day: "2-digit", month: "2-digit" }).replace("/", "-");
      const meta = {
        id: store.newId(),
        app: "deck",
        nombre: "Presentación " + stamp + ".pptx",
        tamano: 0,
        abierto: Date.now(),
        modificado: Date.now(),
        diapositivas: count,
      };
      S.doc = { meta, name: meta.nombre, dirty: true, dirtySince: Date.now(), slides: count, slideIdx: 0 };
      S.thumbs.clear();
      await prepareEditorView();
      await saveActive(true).catch(() => {});
      busy(false);
      renderTop();
      renderFilmstrip();
      renderBottom();
      refreshThumbnails(true);
    } catch (e) {
      busy(false);
      toast("No se pudo crear: " + e.message, 4500);
      S.doc = null;
      showHome();
    } finally {
      S.opening = false;
    }
  }

  async function openDocument(meta, blob) {
    if (!S.ready) { toast("DeckCraft todavía está iniciando"); return; }
    S.opening = true;
    showEditor();
    renderTop(meta.nombre);
    busy(true, "Abriendo " + meta.nombre + "…");
    try {
      const bytes = await blob.arrayBuffer();
      canvas.dispatchEvent(new CustomEvent("deckcraft:open", { detail: { name: meta.nombre, bytes } }));
      await sleep(350);
      let d = null;
      for (let i = 0; i < 40; i++) {
        d = await dc.doc().catch(() => null);
        if (d && Array.isArray(d.slides) && d.slides.length) break;
        await sleep(120);
      }
      const count = d && Array.isArray(d.slides) ? d.slides.length : 1;
      meta.diapositivas = count;
      store.putMeta(meta).catch(() => {});
      S.doc = { meta, name: meta.nombre, dirty: false, dirtySince: 0, slides: count, slideIdx: 0 };
      S.thumbs.clear();
      await prepareEditorView();
      busy(false);
      renderTop();
      renderFilmstrip();
      renderBottom();
      refreshThumbnails(true);
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
    const isDark = prefs.tema === "oscuro" || (prefs.tema === "auto" && matchMedia("(prefers-color-scheme: dark)").matches);
    await dc.set({ ribbonCollapsed: true, notes: false, pane: null, view: "normal", dark: isDark }).catch(() => {});
    await dc.tool("select").catch(() => {});
    S.activeTool = "select";
    await sleep(110);
    applyCanvasCrop(await dc.ui().catch(() => null));
    await sleep(90);
    await fitSlide();
    applyCanvasCrop(await dc.ui().catch(() => null));
    window.ocSyncWakeLock(prefs.pantallaActiva);
  }

  async function goHome() {
    if (S.showing) await dc.showEnd().catch(() => {});
    if (S.doc && S.doc.dirty) {
      busy(true, "Guardando presentación…");
      await saveActive(false).catch(() => {});
      busy(false);
    }
    closeDrawer();
    S.notesOpen = false;
    $("oc-notes-drawer").hidden = true;
    S.doc = null;
    showHome();
  }

  function showEditor() {
    body.classList.add("oc-in-editor");
    renderBottom();
  }

  async function showHome() {
    body.classList.remove("oc-in-editor", "oc-desktop", "oc-modal", "oc-show");
    S.desktop = false;
    S.modal = false;
    S.showing = false;
    $("oc-show-hud").hidden = true;
    window.ocSyncWakeLock(false);
    if (window.ocUpdateReady) { location.reload(); return; }
    await reloadLibrary();
    renderShell();
  }

  function busy(on, text) {
    $("oc-busy").hidden = !on;
    if (text) $("oc-busy-text").textContent = text;
  }

  // ---------- Tira de diapositivas (exacta a PowerPoint para iPhone) ----------
  function renderFilmstrip() {
    const strip = $("oc-filmstrip");
    if (!S.doc) { strip.replaceChildren(); return; }
    const nodes = [];
    for (let i = 0; i < S.doc.slides; i++) {
      const thumbUrl = S.thumbs.get(i);
      const isCur = i === S.doc.slideIdx;
      const btn = h(
        "button",
        {
          class: "oc-slide-thumb",
          type: "button",
          "aria-current": isCur ? "true" : null,
          "aria-label": "Diapositiva " + (i + 1),
          onclick: () => (isCur ? slideContextMenu(i) : goToSlide(i)),
        },
        thumbUrl ? h("img", { src: thumbUrl, alt: "" }) : null,
        h("span", { class: "oc-slide-num", text: String(i + 1) })
      );
      nodes.push(btn);
    }
    nodes.push(
      h("button", {
        class: "oc-slide-add",
        type: "button",
        "aria-label": "Añadir diapositiva",
        icon: "plus",
        onclick: newSlideSheet,
      })
    );
    strip.replaceChildren(...nodes);
    const activeEl = strip.querySelector('[aria-current="true"]');
    if (activeEl) activeEl.scrollIntoView({ behavior: "smooth", inline: "center", block: "nearest" });
  }

  async function refreshThumbnails(saveCover) {
    if (!S.doc) return;
    const gen = ++S.thumbGen;
    const count = Math.min(S.doc.slides, 40);
    for (let i = 0; i < count; i++) {
      if (gen !== S.thumbGen || !S.doc) return;
      try {
        const res = await dc.renderSlide(i, 0.22);
        if (res && res.png) {
          const url = "data:image/png;base64," + res.png;
          S.thumbs.set(i, url);
          if (i === 0 && saveCover && S.doc.meta) {
            S.doc.meta.miniatura = url;
            store.putMeta(S.doc.meta).catch(() => {});
          }
          renderFilmstrip();
        }
      } catch (e) {}
      await sleep(30);
    }
  }

  async function goToSlide(idx) {
    if (!S.doc) return;
    const clamped = Math.max(0, Math.min(S.doc.slides - 1, idx));
    await dc.run("slide.go", { index: clamped }).catch(() => {});
    S.doc.slideIdx = clamped;
    renderTop();
    renderFilmstrip();
    updateSlidePill();
    if (S.notesOpen) refreshNotesBox();
  }

  async function stepSlide(dir) {
    if (!S.doc) return;
    const next = S.doc.slideIdx + dir;
    if (next < 0 || next >= S.doc.slides) return;
    await goToSlide(next);
  }

  function slideContextMenu(idx) {
    sheet({
      title: "Diapositiva " + (idx + 1),
      build: (close) => [
        list([
          { i: "slidePlus", t: "Nueva diapositiva después", run: () => { close(); newSlideSheet(); } },
          { i: "copy", t: "Duplicar diapositiva", run: async () => { close(); await execDeck("slide.duplicate"); await syncDeckAfterEdit(); } },
          { i: "sorter", t: "Cambiar diseño…", run: () => { close(); layoutPickerSheet(); } },
          { i: "play", t: "Presentar desde esta diapositiva", run: () => { close(); startShow(idx); } },
          { i: "eye", t: "Ocultar / mostrar diapositiva", run: () => { close(); execDeck("slide.hide"); } },
        ]),
        S.doc.slides > 1
          ? list([{ i: "trash", t: "Eliminar diapositiva", danger: true, run: async () => { close(); await execDeck("slide.delete"); await syncDeckAfterEdit(); } }])
          : null,
      ].filter(Boolean),
    });
  }

  function newSlideSheet() {
    const layouts = [
      ["titleAndContent", "Título y objetos"],
      ["title", "Diapositiva de título"],
      ["sectionHeader", "Encabezado de sección"],
      ["twoContent", "Dos objetos"],
      ["comparison", "Comparación"],
      ["titleOnly", "Solo título"],
      ["blank", "En blanco"],
    ];
    sheet({
      title: "Nueva diapositiva",
      build: (close) => [
        list(
          layouts.map(([layout, t]) => ({
            i: "slidePlus",
            t,
            run: async () => {
              close();
              await execDeck("slide.new", { layout });
              await syncDeckAfterEdit();
            },
          }))
        ),
      ],
    });
  }

  function layoutPickerSheet() {
    const layouts = [
      ["title", "Título"],
      ["titleAndContent", "Título y objetos"],
      ["twoContent", "Dos objetos"],
      ["titleOnly", "Solo título"],
      ["blank", "En blanco"],
    ];
    sheet({
      title: "Cambiar diseño",
      build: (close) => [list(layouts.map(([layout, t]) => ({ i: "sorter", t, run: async () => { close(); await execDeck("slide.layout", { layout }); await syncDeckAfterEdit(); } })))],
    });
  }

  async function syncDeckAfterEdit() {
    await sleep(100);
    const d = await dc.doc().catch(() => null);
    if (d && S.doc) {
      S.doc.slides = Array.isArray(d.slides) ? Math.max(1, d.slides.length) : S.doc.slides;
      if (d.selection && typeof d.selection.slide === "number") S.doc.slideIdx = d.selection.slide;
      S.doc.meta.diapositivas = S.doc.slides;
      store.putMeta(S.doc.meta).catch(() => {});
    }
    renderTop();
    renderFilmstrip();
    updateSlidePill();
    refreshThumbnails(true);
  }

  // ---------- Notas del orador ----------
  async function toggleNotes() {
    S.notesOpen = !S.notesOpen;
    const box = $("oc-notes-drawer");
    box.hidden = !S.notesOpen;
    if (S.notesOpen) await refreshNotesBox();
    renderTop();
    setTimeout(() => { dc.ui().then(applyCanvasCrop).catch(() => {}); }, 80);
  }

  async function refreshNotesBox() {
    const box = $("oc-notes-drawer");
    if (!S.notesOpen || !S.doc) return;
    let currentText = "";
    try {
      const d = await dc.doc();
      const sl = d && d.slides && d.slides[S.doc.slideIdx];
      currentText = (sl && sl.notes) || "";
    } catch (e) {}
    const area = h("textarea", { class: "oc-notes-area", placeholder: "Escribe notas del orador para la diapositiva " + (S.doc.slideIdx + 1) + "…" });
    area.value = currentText;
    let tmr = null;
    area.addEventListener("input", () => {
      clearTimeout(tmr);
      tmr = setTimeout(() => {
        dc.run("slide.notes", { text: area.value }).catch(() => {});
      }, 450);
    });
    box.replaceChildren(
      h("div", { class: "oc-notes-head" }, h("span", { text: "Notas · Diapositiva " + (S.doc.slideIdx + 1) }), h("button", { class: "oc-link-btn", type: "button", text: "Listo", onclick: () => { dc.run("slide.notes", { text: area.value }).catch(() => {}); toggleNotes(); } })),
      area
    );
  }

  // ---------- Barra superior PowerPoint iPhone ----------
  function renderTop(fallbackName) {
    const title = S.doc ? S.doc.meta.nombre : fallbackName || "DeckCraft";
    const sub = S.doc ? "Diap. " + (S.doc.slideIdx + 1) + " de " + S.doc.slides + (S.doc.dirty ? " · editado" : "") : "";
    $("oc-top").replaceChildren(
      h("button", { class: "oc-icon-btn", type: "button", "aria-label": "Volver", icon: "back", onclick: goHome }),
      h("button", { class: "oc-title", type: "button", onclick: () => { if (S.doc) renameSheet(S.doc.meta); } }, h("b", { text: title }), h("small", { text: sub })),
      h("button", { class: "oc-icon-btn", type: "button", "aria-label": "Presentar", icon: "play", onclick: () => startShow(S.doc ? S.doc.slideIdx : 0) }),
      h("button", { class: "oc-icon-btn" + (S.notesOpen ? " oc-on" : ""), type: "button", "aria-label": "Notas del orador", icon: "notes", onclick: toggleNotes }),
      h("button", { class: "oc-icon-btn", type: "button", "aria-label": "Deshacer", icon: "undo", onclick: () => execDeck("edit.undo") }),
      h("button", { class: "oc-icon-btn", type: "button", "aria-label": "Compartir", icon: "share", onclick: shareMenu }),
      h("button", { class: "oc-icon-btn", type: "button", "aria-label": "Más", icon: "more", onclick: moreMenu })
    );
  }

  async function startShow(fromIdx) {
    closeDrawer();
    if (S.notesOpen) toggleNotes();
    await dc.showStart(fromIdx != null ? fromIdx : 0).catch((e) => toast(e.message));
  }

  // ---------- Barra rápida y Cinta móvil PowerPoint iPhone ----------
  function renderBottom() {
    const items = [
      { icon: "slidePlus", label: "Nueva diap.", fn: newSlideSheet },
      { icon: "textBox", label: "Texto", fn: () => execDeck("shape.insert", { preset: "rect", rect: [180, 180, 560, 140], text: "Nuevo texto" }) },
      { icon: "image", label: "Imagen", fn: () => $("oc-img").click() },
      { icon: "shape", label: "Forma", fn: insertShapeSheet },
      { icon: "table", label: "Tabla", fn: insertTableSheet },
      { icon: "bold", label: "Negrita", fn: () => execDeck("format.bold") },
      { icon: "italic", label: "Cursiva", fn: () => execDeck("format.italic") },
      { icon: "underline", label: "Subrayado", fn: () => execDeck("format.underline") },
      { icon: "textColor", label: "Color", fn: fontColorPicker },
      { icon: "redo", label: "Rehacer", fn: () => execDeck("edit.redo") },
    ];
    $("oc-bottom").replaceChildren(
      h(
        "div",
        { class: "oc-quickbar" },
        h(
          "div",
          { class: "oc-quick-scroll" },
          items.map((it) => h("button", { class: "oc-qbtn", type: "button", "aria-label": it.label, icon: it.icon, onclick: it.fn }))
        ),
        h("button", {
          class: "oc-drawer-toggle",
          type: "button",
          "aria-label": S.drawerOpen ? "Cerrar cinta" : "Abrir cinta de PowerPoint",
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
    setTimeout(() => { dc.ui().then(applyCanvasCrop).catch(() => {}); }, 80);
  }
  function closeDrawer() {
    if (!S.drawerOpen) return;
    S.drawerOpen = false;
    $("oc-ribbon-drawer").hidden = true;
    renderBottom();
    setTimeout(() => { dc.ui().then(applyCanvasCrop).catch(() => {}); }, 80);
  }

  function renderDrawer() {
    const drawer = $("oc-ribbon-drawer");
    drawer.hidden = !S.drawerOpen;
    if (!S.drawerOpen) return;
    const tabs = [
      ["inicio", "Inicio"],
      ["insertar", "Insertar"],
      ["dibujar", "Dibujar"],
      ["diseno", "Diseño"],
      ["transiciones", "Transiciones"],
      ["presentacion", "Presentación"],
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
        h(
          "div",
          { class: "oc-rgroup" },
          h("div", { class: "oc-rgroup-title", text: "Diapositivas" }),
          h("div", { class: "oc-rrow" },
            h("button", { class: "oc-rbtn", type: "button", icon: "slidePlus", text: "Nueva", onclick: newSlideSheet }),
            h("button", { class: "oc-rbtn", type: "button", icon: "sorter", text: "Diseño", onclick: layoutPickerSheet }),
            h("button", { class: "oc-rbtn", type: "button", icon: "copy", text: "Duplicar", onclick: async () => { await execDeck("slide.duplicate"); await syncDeckAfterEdit(); } })
          )
        ),
        h(
          "div",
          { class: "oc-rgroup" },
          h("div", { class: "oc-rgroup-title", text: "Fuente y párrafo" }),
          h("div", { class: "oc-rrow" },
            h("button", { class: "oc-rbtn", type: "button", icon: "bold", onclick: () => execDeck("format.bold") }),
            h("button", { class: "oc-rbtn", type: "button", icon: "italic", onclick: () => execDeck("format.italic") }),
            h("button", { class: "oc-rbtn", type: "button", icon: "underline", onclick: () => execDeck("format.underline") }),
            h("button", { class: "oc-rbtn", type: "button", icon: "strike", onclick: () => execDeck("format.strikethrough") }),
            h("button", { class: "oc-rbtn", type: "button", text: "A-", onclick: () => execDeck("format.shrink") }),
            h("button", { class: "oc-rbtn", type: "button", text: "A+", onclick: () => execDeck("format.grow") }),
            h("button", { class: "oc-rbtn", type: "button", icon: "textColor", text: "Color", onclick: fontColorPicker }),
            h("button", { class: "oc-rbtn", type: "button", icon: "alignLeft", onclick: () => execDeck("format.alignLeft") }),
            h("button", { class: "oc-rbtn", type: "button", icon: "alignCenter", onclick: () => execDeck("format.alignCenter") }),
            h("button", { class: "oc-rbtn", type: "button", icon: "alignRight", onclick: () => execDeck("format.alignRight") }),
            h("button", { class: "oc-rbtn", type: "button", icon: "bullets", onclick: () => execDeck("format.bullets") }),
            h("button", { class: "oc-rbtn", type: "button", icon: "numbering", onclick: () => execDeck("format.numbering") })
          )
        ),
      ];
    }
    if (tab === "insertar") {
      return [
        h("div", { class: "oc-rgrid-3" },
          rtile("textBox", "Cuadro texto", () => execDeck("shape.insert", { preset: "rect", rect: [180, 180, 560, 140], text: "Escribe aquí" })),
          rtile("image", "Imagen", () => { closeDrawer(); $("oc-img").click(); }),
          rtile("shape", "Formas", insertShapeSheet),
          rtile("table", "Tabla", insertTableSheet),
          rtile("chart", "Gráfico", () => execDeck("insert.chart", { type: "column", x: 140, y: 120, w: 660, h: 320, title: "Datos", categories: ["Q1", "Q2", "Q3", "Q4"], series: [{ name: "Ventas", values: [24, 38, 31, 46] }] })),
          rtile("notes", "Notas", () => { closeDrawer(); toggleNotes(); })
        ),
      ];
    }
    if (tab === "dibujar") {
      const tools = [
        ["select", "Seleccionar", "edit"],
        ["pen", "Pluma", "draw"],
        ["highlighter", "Resaltador", "highlight"],
        ["eraser", "Borrador", "eraser"],
      ];
      return [
        h("div", { class: "oc-rgrid-2" },
          tools.map(([id, label, ic]) =>
            h("button", {
              class: "oc-rbtn",
              type: "button",
              icon: ic,
              "aria-pressed": String(S.activeTool === id),
              text: label,
              onclick: async () => {
                S.activeTool = id;
                await dc.tool(id).catch(() => {});
                renderDrawer();
                toast("Herramienta: " + label);
              },
            })
          )
        ),
      ];
    }
    if (tab === "diseno") {
      return [
        h("div", { class: "oc-rgroup" },
          h("div", { class: "oc-rgroup-title", text: "Color de fondo de diapositiva" }),
          colorGrid([
            ["#FFFFFF", "Blanco", "#FFFFFF"],
            ["#F5F6FA", "Humo", "#F5F6FA"],
            ["#1A2238", "Azul noche", "#1A2238"],
            ["#141418", "Carbón", "#141418"],
            ["#FEF3E7", "Crema", "#FEF3E7"],
            ["#EAF4EE", "Menta", "#EAF4EE"],
          ], async (hex) => { await execDeck("design.background", { color: hex }); refreshThumbnails(false); })
        ),
        h("div", { class: "oc-rgroup" },
          h("div", { class: "oc-rgroup-title", text: "Tamaño de diapositiva" }),
          h("div", { class: "oc-rrow" },
            h("button", { class: "oc-rbtn", type: "button", text: "Panorámica 16:9", onclick: () => execDeck("design.slideSize", { preset: "widescreen" }) }),
            h("button", { class: "oc-rbtn", type: "button", text: "Estándar 4:3", onclick: () => execDeck("design.slideSize", { preset: "standard" }) })
          )
        ),
      ];
    }
    if (tab === "transiciones") {
      const trans = [
        ["none", "Ninguna"],
        ["fade", "Desvanecer"],
        ["push", "Empujar"],
        ["wipe", "Barrido"],
        ["morph", "Transformación"],
        ["cover", "Cubrir"],
      ];
      return [
        h("div", { class: "oc-rgrid-3" }, trans.map(([kind, label]) => rtile("transition", label, () => execDeck("transition.set", { type: kind })))),
        h("button", { class: "oc-secondary", type: "button", icon: "play", text: "Vista previa de animación", onclick: () => execDeck("animation.preview") }),
      ];
    }
    return [
      h("div", { class: "oc-rgrid-2" },
        rtile("play", "Desde el principio", () => startShow(0)),
        rtile("play", "Desde esta diap.", () => startShow(S.doc ? S.doc.slideIdx : 0)),
        rtile("zoom", "Ajustar diapositiva", fitSlide),
        rtile("pdf", "Exportar a PDF", () => exportAs("pdf")),
        rtile("desktop", "Modo PC completo", () => { closeDrawer(); setDesktop(true); })
      ),
    ];
  }

  function rtile(ic, label, onclick) {
    return h("button", { class: "oc-rtile", type: "button", icon: ic, onclick }, h("span", { text: label }));
  }

  async function execDeck(cmd, params) {
    try {
      const res = await dc.run(cmd, params || {});
      setTimeout(() => refreshCurrentThumb(), 220);
      return res;
    } catch (e) {
      toast("No disponible: " + (e.message || e), 3000);
      return null;
    }
  }

  async function refreshCurrentThumb() {
    if (!S.doc) return;
    try {
      const i = S.doc.slideIdx;
      const res = await dc.renderSlide(i, 0.22);
      if (res && res.png) {
        const url = "data:image/png;base64," + res.png;
        S.thumbs.set(i, url);
        if (i === 0 && S.doc.meta) { S.doc.meta.miniatura = url; store.putMeta(S.doc.meta).catch(() => {}); }
        renderFilmstrip();
      }
    } catch (e) {}
  }

  function fontColorPicker() {
    sheet({
      title: "Color de texto",
      build: (close) => [
        colorGrid([
          ["#000000", "Negro", "#000000"],
          ["#FFFFFF", "Blanco", "#FFFFFF"],
          ["#C43E1C", "Naranja PPT", "#C43E1C"],
          ["#185ABD", "Azul", "#185ABD"],
          ["#107C41", "Verde", "#107C41"],
          ["#F59F00", "Dorado", "#F59F00"],
          ["#D92D20", "Rojo", "#D92D20"],
          ["#6741D9", "Morado", "#6741D9"],
          ["#595959", "Gris", "#595959"],
        ], (v) => { close(); execDeck("format.color", { color: v }); }),
      ],
    });
  }

  function insertShapeSheet() {
    const shapes = [
      ["rect", "Rectángulo"],
      ["roundRect", "Rectángulo redondeado"],
      ["ellipse", "Círculo / Óvalo"],
      ["rightArrow", "Flecha derecha"],
      ["star5", "Estrella"],
      ["triangle", "Triángulo"],
    ];
    sheet({
      title: "Insertar forma",
      build: (close) => [list(shapes.map(([preset, t]) => ({ i: "shape", t, run: () => { close(); execDeck("shape.insert", { preset, rect: [300, 180, 320, 180] }); } })))],
    });
  }

  function insertTableSheet() {
    sheet({
      title: "Insertar tabla",
      build: (close) => [
        h("div", { class: "oc-btn-row" },
          h("button", { class: "oc-secondary", type: "button", text: "2 × 2", onclick: () => { close(); execDeck("insert.table", { rows: 2, cols: 2 }); } }),
          h("button", { class: "oc-secondary", type: "button", text: "3 × 3", onclick: () => { close(); execDeck("insert.table", { rows: 3, cols: 3 }); } }),
          h("button", { class: "oc-secondary", type: "button", text: "4 × 4", onclick: () => { close(); execDeck("insert.table", { rows: 4, cols: 4 }); } })
        ),
      ],
    });
  }

  async function insertImageFromFile(file) {
    try {
      const dt = new DataTransfer();
      dt.items.add(file);
      canvas.dispatchEvent(new DragEvent("drop", { bubbles: true, cancelable: true, dataTransfer: dt }));
      toast("Imagen añadida a la diapositiva");
      setTimeout(() => refreshCurrentThumb(), 500);
    } catch (e) {
      toast("No se pudo añadir la imagen");
    }
  }

  async function exportAs(fmt) {
    if (!S.doc) return;
    const base = S.doc.meta.nombre.replace(/\.[^.]+$/, "");
    toast("Generando " + fmt.toUpperCase() + "…");
    await dc.run("file.export", { path: base + "." + fmt }).catch((e) => toast(e.message));
  }

  function shareMenu() {
    if (!S.doc) return;
    sheet({
      title: "Compartir o exportar",
      text: S.doc.meta.nombre,
      build: (close) => [
        list([
          { i: "share", t: "Compartir como PowerPoint (.pptx)", small: "Guardar en Archivos", run: async () => { close(); busy(true, "Preparando .pptx…"); try { const f = await currentFile(); busy(false); if (f) shareFile(f); } catch (e) { busy(false); toast(e.message); } } },
          { i: "pdf", t: "Exportar a PDF", small: ".pdf", run: () => { close(); exportAs("pdf"); } },
          { i: "image", t: "Exportar diapositiva como PNG", small: ".png", run: () => { close(); exportAs("png"); } },
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
          { i: "check", t: "Guardar ahora", run: async () => { close(); try { await saveActive(false); toast("Guardado en este iPhone"); } catch (e) { toast(e.message); } } },
          { i: "play", t: "Presentar desde el principio", run: () => { close(); startShow(0); } },
          { i: "notes", t: S.notesOpen ? "Ocultar notas del orador" : "Ver notas del orador", run: () => { close(); toggleNotes(); } },
          { i: "rename", t: "Renombrar", run: () => { close(); renameSheet(S.doc.meta); } },
          { i: "desktop", t: "Interfaz completa (PC)", small: "Todas las herramientas", run: () => { close(); setDesktop(true); } },
        ]),
      ],
    });
  }

  function settingsSheet() {
    sheet({
      title: "Ajustes de DeckCraft",
      build: (close) => [
        h("div", { class: "oc-label", text: "Apariencia" }),
        seg([["auto", "Sistema"], ["claro", "Claro"], ["oscuro", "Oscuro"]], () => prefs.tema, (v) => {
          prefs.tema = v;
          savePrefs();
          applyTheme();
          const dark = v === "oscuro" || (v === "auto" && matchMedia("(prefers-color-scheme: dark)").matches);
          if (S.ready) dc.set({ dark }).catch(() => {});
        }),
        h("div", { class: "oc-list" }, swRow("Mantener pantalla encendida al presentar", () => prefs.pantallaActiva, (v) => { prefs.pantallaActiva = v; savePrefs(); })),
        list([
          { i: "refresh", t: "Buscar actualizaciones", run: () => { close(); window.ocCheckUpdate(); } },
          { i: "apps", t: "Ir al portal OfficeCraft", small: "Word · PowerPoint · Excel", run: () => { location.href = "../"; } },
        ]),
        h("p", { text: "DeckCraft para iPhone · Motor de código abierto (github.com/storytold/deckcraft) " + (S.appVersion ? "v" + S.appVersion : "") + ". Tus archivos no salen del teléfono." }),
      ],
    });
  }

  let backBtn = null;
  async function setDesktop(on) {
    S.desktop = on;
    body.classList.toggle("oc-desktop", on);
    if (on) {
      closeDrawer();
      await dc.set({ ribbonCollapsed: false }).catch(() => {});
      applyCanvasCrop(null);
      backBtn = backBtn || h("button", { class: "oc-back-mobile", type: "button", text: "Volver a vista iPhone", onclick: () => setDesktop(false) });
      body.append(backBtn);
    } else {
      if (backBtn) backBtn.remove();
      await dc.set({ ribbonCollapsed: true, notes: false, pane: null }).catch(() => {});
      await sleep(120);
      applyCanvasCrop(await dc.ui().catch(() => null));
      await fitSlide();
    }
  }

  window.ocOnUpdate = () => {
    if (!body.classList.contains("oc-in-editor")) location.reload();
    else toast("Hay una versión nueva de DeckCraft: se aplicará al volver a Inicio.", 5500);
  };

  window.deckcraftMobile = { openCreateSheet, createFromTemplate, startShow, toggleDrawer, setDesktop, goHome, state: () => S };
})();
