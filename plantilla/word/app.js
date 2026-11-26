// WordCraft para iPhone: interfaz móvil fiel a Microsoft Word para iOS sobre el motor
// oficial de WordCraft (data-host="page" + window.wordcraft) con gestos a 120 fps sin tirones.
(function () {
  "use strict";

  const $ = (id) => document.getElementById(id);
  const { icon, h, sheet, list, seg, swRow, colorGrid, sizeText, describe, makeSearchBox, sleep } = window.ocUI;
  const store = window.ocStore;
  const toast = (t, ms) => window.ocToast(t, ms);
  const canvas = $("wordcraft_canvas");
  const body = document.body;

  // ---------- Preferencias ----------
  const PREFS_KEY = "officecraft-word-prefs";
  const prefs = Object.assign(
    {
      tema: "auto",        // "auto" | "claro" | "oscuro"
      vista: "movil",      // "movil" (webLayout) | "impresion" (printLayout)
      paginaOscura: false,
      autor: "",
      vistaArchivos: "lista",
      orden: "reciente",
      pantallaActiva: false,
      motor: "auto",       // "auto" | "webgl"
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

  // ---------- Puente WordCraft API ----------
  const wc = {
    ok: () => !!(window.wordcraft && typeof window.wordcraft.request === "function"),
    ui: () => window.wordcraft.uiInspect(),
    doc: (params) => window.wordcraft.docInspect(params || {}),
    run: (command, params) => window.wordcraft.execute(command, params || {}),
    req: (method, params) => window.wordcraft.request(method, params || {}),
    key: (key, opts) => window.wordcraft.request("ui.key", Object.assign({ key }, opts || {})),
    renderPage: (page, scale) => window.wordcraft.request("ui.render", { page: page || 1, scale: scale || 0.45 }),
  };

  // ---------- Estado ----------
  const S = {
    ready: false,
    engineReady: false,
    appVersion: (window.OC_APP && window.OC_APP.version) || "",
    tab: "inicio",           // "inicio" | "destacados" | "archivos"
    homeSearchOpen: false,
    query: "",
    filesFolder: null,
    files: [],
    folders: [],
    doc: null,               // { id, meta, name, dirty, dirtySince, pages, page, zoom, mobileView }
    opening: false,
    desktop: false,
    modal: false,
    drawerOpen: false,
    drawerTab: "inicio",     // "inicio" | "insertar" | "disposicion" | "revisar" | "vista"
    searching: false,
    searchQuery: "",
    immersive: false,
    fmt: {},                 // Estado actual de formato (bold/italic/underline/align...)
    ratio: 1,
    zoomPct: 100,
    fitZoomPct: 100,
    crop: { l: 0, t: 0, w: 0, h: 0 },
    polling: false,
    saveWaiters: [],
    lastUi: null,
  };

  // ---------- Eventos del modo anfitrión (data-host="page") ----------
  canvas.addEventListener("wordcraft:ready", (e) => {
    S.engineReady = true;
    if (e.detail && e.detail.appVersion) S.appVersion = e.detail.appVersion;
  });
  canvas.addEventListener("wordcraft:dirty", (e) => {
    if (!S.doc) return;
    setDirty(!!(e.detail && e.detail.dirty));
  });
  canvas.addEventListener("wordcraft:failed", (e) => {
    window.ocShowError("WordCraft no pudo iniciar: " + ((e.detail && e.detail.error) || "sin WebGL2/WebGPU"));
  });
  canvas.addEventListener("wordcraft:save", (e) => {
    onEngineSave(e.detail || {}).catch((err) => toast("Error al guardar: " + err.message, 4500));
  });

  function setDirty(dirty) {
    if (!S.doc || S.doc.dirty === dirty) return;
    S.doc.dirty = dirty;
    S.doc.dirtySince = dirty ? Date.now() : 0;
    renderTop();
  }

  async function onEngineSave(d) {
    const rawName = d.name || (S.doc ? S.doc.meta.nombre : "Documento.docx");
    const isPdf = /\.pdf$/i.test(rawName);
    const isDoc = /\.(docx|odt|rtf|md|txt)$/i.test(rawName);
    const bytes = d.bytes instanceof ArrayBuffer ? d.bytes : (d.bytes && d.bytes.buffer) || new ArrayBuffer(0);
    const file = new File([bytes], rawName, { type: window.ocMimeFor(rawName) });

    const waiters = S.saveWaiters.splice(0, S.saveWaiters.length);
    if (isDoc && S.doc && !isPdf) {
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
    if (!waiters.length) {
      await shareFile(file);
    }
  }

  function saveActive() {
    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => {
        S.saveWaiters = S.saveWaiters.filter((w) => w !== waiter);
        reject(new Error("WordCraft no respondió al guardar"));
      }, 20000);
      const waiter = {
        resolve: (f) => { clearTimeout(timer); resolve(f); },
        reject: (e) => { clearTimeout(timer); reject(e); },
      };
      S.saveWaiters.push(waiter);
      wc.run("file.save").catch((e) => {
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
      window.ocShowError("Error al iniciar WordCraft: " + (e && e.message ? e.message : e));
    });
  });

  async function boot() {
    for (let i = 0; i < 400 && !wc.ok(); i++) await sleep(50);
    if (!wc.ok()) throw new Error("No se inicializó la API window.wordcraft");
    for (let i = 0; i < 80; i++) {
      try {
        S.lastUi = await wc.ui();
        break;
      } catch (e) {
        await sleep(150);
      }
    }
    S.ready = true;

    // Configuración inicial optimizada para celular (sin tirones y en español)
    const quiet = (p) => p.catch(() => {});
    quiet(wc.run("ui.language", { value: "es" }));
    quiet(wc.run("ui.startScreen", { value: false }));
    quiet(wc.run("ui.collapseRibbon", { value: true }));
    quiet(wc.run("ui.quickAccessToolbar", { value: false }));
    quiet(wc.run("view.ruler", { value: false }));
    quiet(wc.run("ui.theme", { value: prefs.tema === "oscuro" ? "dark" : prefs.tema === "claro" ? "light" : "system" }));
    if (prefs.autor) quiet(wc.run("file.setAuthor", { name: prefs.autor }));

    gestures = window.ocGestures.attach(canvas, {
      mode: () => (S.desktop ? "raton-zoom" : S.modal ? "raton" : "ver"),
      ratio: () => S.ratio,
      axisLock: "doc",
      onPinchCommit: async ({ scale }) => {
        if (!S.doc || S.desktop || S.modal) return;
        const cur = S.zoomPct || 100;
        const next = Math.max(40, Math.min(400, Math.round(cur * scale)));
        S.zoomPct = next;
        await wc.run("view.zoom", { value: next }).catch(() => {});
        updatePagePill();
      },
      onDoubleTap: () => toggleDoubleTapZoom(),
      onTap: () => {
        if (S.drawerOpen) closeDrawer();
      },
      onLongPress: (x, y) => showPressRing(x, y),
    });

    initStaticControls();
    await reloadLibrary();
    renderShell();
    window.ocHideSplash();

    // Abrir documento si se llegó desde el Portal de OfficeCraft (?open=<id> o ?new=1)
    const qs = new URLSearchParams(location.search);
    const openId = qs.get("open");
    if (openId) {
      const m = S.files.find((x) => x.id === openId);
      if (m) openStoredDoc(m);
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

  async function toggleDoubleTapZoom() {
    if (!S.doc || S.desktop || S.modal) return;
    const base = S.fitZoomPct || 100;
    if (S.zoomPct > base * 1.25) {
      await fitPageWidth();
    } else {
      const target = Math.min(260, Math.round(base * 1.8));
      S.zoomPct = target;
      await wc.run("view.zoom", { value: target }).catch(() => {});
      updatePagePill();
    }
  }

  async function fitPageWidth() {
    await wc.run("view.pageWidth").catch(() => {});
    await sleep(90);
    const u = await wc.ui().catch(() => null);
    if (u && u.scale) {
      S.zoomPct = Math.round(u.scale * 100);
      S.fitZoomPct = S.zoomPct;
      updatePagePill();
    }
  }

  // ---------- Recorte exacto al área de documento (0 % barras de escritorio) ----------
  function applyCanvasCrop(u) {
    const docEl = $("oc-doc");
    if (!docEl) return;
    if (S.desktop || S.modal || !u || !u.window || !u.canvasRect) {
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

    const chromeL = Math.max(0, Math.min(ww * 0.45, cx));
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

  // ---------- Sondeo inteligente (se pausa por completo durante scroll/zoom = 0 tirones) ----------
  async function tick() {
    if (!S.ready || S.polling || document.visibilityState !== "visible") return;
    if (gestures && gestures.busy()) return;
    S.polling = true;
    try {
      const u = await wc.ui();
      S.lastUi = u;
      if (u.window && u.window[0]) S.ratio = canvas.clientWidth / u.window[0];
      const modal = !!(u.dialogs && u.dialogs.length > 0);
      if (modal !== S.modal) {
        S.modal = modal;
        body.classList.toggle("oc-modal", modal);
        if (modal) {
          closeDrawer();
          wc.run("ui.interfaceSize", { percent: 85 }).catch(() => {});
        } else {
          wc.run("ui.interfaceSize", { value: "auto" }).catch(() => {});
        }
      }
      if (!body.classList.contains("oc-in-editor")) return;
      applyCanvasCrop(u);
      if (S.opening || !S.doc) return;

      if (u.scale) S.zoomPct = Math.round(u.scale * 100);
      if (typeof u.dirty === "boolean") setDirty(u.dirty);

      // Número de página visible
      const pages = u.pages || [];
      if (pages.length && u.canvasRect) {
        const midY = u.canvasRect[1] + u.canvasRect[3] * 0.45;
        let cur = 1;
        for (let i = 0; i < pages.length; i++) {
          const p = pages[i];
          if (p[1] <= midY) cur = i + 1;
        }
        S.doc.page = cur;
        if (pages.length > S.doc.pages) S.doc.pages = pages.length;
      }
      updatePagePill();

      // Estado de formato actual para iluminar B / I / U en la barra inferior
      if (S.drawerOpen || Date.now() - (S.fmtAt || 0) > 900) {
        S.fmtAt = Date.now();
        const st = await wc.run("format.state").catch(() => null);
        if (st) {
          S.fmt = st;
          updateQuickBarState();
        }
      }

      // Guardado automático cada 25 s si hay cambios
      if (S.doc.dirty && S.doc.dirtySince && Date.now() - S.doc.dirtySince > 25000 && !S.modal && !S.saveWaiters.length) {
        S.doc.dirtySince = Date.now();
        saveActive().catch(() => {});
      }
    } catch (e) {
      // Motor ocupado
    } finally {
      S.polling = false;
    }
  }

  function updatePagePill() {
    const pill = $("oc-page-pill");
    if (!pill) return;
    if (!S.doc || S.desktop || S.modal) {
      pill.hidden = true;
      return;
    }
    pill.hidden = false;
    const modeLabel = S.doc.mobileView ? "Móvil" : (S.zoomPct + "%");
    pill.textContent = "Pág. " + (S.doc.page || 1) + " / " + Math.max(1, S.doc.pages || 1) + " · " + modeLabel;
  }

  // ---------- Plantillas estilo Microsoft Word para iPhone ----------
  const TEMPLATES = [
    { id: "blank", name: "Documento en blanco", style: "blank" },
    { id: "resume", name: "Currículum", style: "doc" },
    { id: "report", name: "Informe", style: "doc" },
    { id: "letter", name: "Carta formal", style: "doc" },
    { id: "notes", name: "Apuntes", style: "doc" },
    { id: "sample", name: "Ejemplo WordCraft", style: "doc" },
  ];

  // ---------- Shell Home (Inicio / Destacados / Archivos) ----------
  function initStaticControls() {
    const fab = $("oc-create-fab");
    fab.replaceChildren();
    fab.insertAdjacentHTML("afterbegin", icon("plus"));
    fab.addEventListener("click", openCreateSheet);
    $("oc-page-pill").addEventListener("click", zoomAndPageSheet);

    $("oc-file").addEventListener("change", async () => {
      const file = $("oc-file").files && $("oc-file").files[0];
      if (!file) return;
      $("oc-file").value = "";
      const meta = {
        id: store.newId(),
        app: "word",
        nombre: file.name || "Documento.docx",
        tamano: file.size,
        abierto: Date.now(),
        modificado: file.lastModified || Date.now(),
      };
      try { await store.save(meta, file); } catch (e) { toast("Se abrirá sin copia local"); }
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
          {
            class: "oc-tab-btn",
            type: "button",
            "aria-current": S.tab === tb.id ? "page" : null,
            onclick: () => switchTab(tb.id),
          },
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
    try { S.files = await store.list("word"); } catch (e) { S.files = []; }
    try { S.folders = await store.listFolders("word"); } catch (e) { S.folders = []; }
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
    const head = $("oc-app-head");
    head.replaceChildren(
      ...[
        h(
          "div",
          { class: "oc-head-row" },
          h("img", { class: "oc-brand-logo", src: "./icons/icon-192.png", alt: "" }),
          h("h1", { class: "oc-head-title", text: "WordCraft" }),
          h("button", {
            class: "oc-head-btn" + (S.homeSearchOpen ? " oc-active" : ""),
            type: "button",
            "aria-label": "Buscar",
            icon: "search",
            onclick: () => {
              S.homeSearchOpen = !S.homeSearchOpen;
              if (!S.homeSearchOpen) S.query = "";
              renderHomeTab();
            },
          }),
          h("button", {
            class: "oc-head-btn",
            type: "button",
            "aria-label": "Suite OfficeCraft",
            icon: "apps",
            onclick: () => { location.href = "../"; },
          }),
          h("button", { class: "oc-head-btn", type: "button", "aria-label": "Ajustes", icon: "settings", onclick: settingsSheet })
        ),
        S.homeSearchOpen
          ? makeSearchBox("Buscar documentos…", S.query, (q) => { S.query = q; renderHomeBody(); }, true)
          : null,
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
        h("div", { class: "oc-section-head" }, h("h2", { text: "Nuevo documento" }), h("button", { class: "oc-link-btn", type: "button", text: "Importar", onclick: pickFile })),
        h(
          "div",
          { class: "oc-tpl-strip" },
          TEMPLATES.map((tpl) => {
            const preview =
              tpl.style === "blank"
                ? h("div", { class: "oc-tpl-preview" }, h("span", { icon: "plus" }))
                : h(
                    "div",
                    { class: "oc-tpl-preview" },
                    h(
                      "div",
                      { class: "oc-tpl-lines" },
                      h("i", { class: "oc-tpl-accent" }),
                      h("i", {}),
                      h("i", {}),
                      h("i", { style: "width:78%" }),
                      h("i", { style: "width:60%" })
                    )
                  );
            return h(
              "button",
              { class: "oc-tpl-card", type: "button", onclick: () => createFromTemplate(tpl.id) },
              preview,
              h("b", { text: tpl.name })
            );
          })
        )
      );
    }

    let listItems = S.files;
    if (q) listItems = listItems.filter((m) => (m.nombre || "").toLowerCase().includes(q));
    listItems = sorted(listItems);

    nodes.push(
      h(
        "div",
        { class: "oc-section-head" },
        h("h2", { text: q ? "Resultados" : "Recientes" }),
        h("button", { class: "oc-sort-btn", type: "button", icon: "sort", onclick: sortSheet }, h("span", { text: prefs.orden === "nombre" ? "Nombre" : prefs.orden === "tamano" ? "Tamaño" : "Reciente" }))
      )
    );

    if (listItems.length) {
      nodes.push(renderFileList(listItems));
    } else {
      nodes.push(
        h(
          "div",
          { class: "oc-empty" },
          h("div", { icon: "file" }),
          h("b", { text: q ? "Sin coincidencias" : "Tus documentos Word aparecerán aquí" }),
          h("div", { text: q ? "No encontramos «" + S.query + "»." : "Crea un documento en blanco o abre un .docx desde Archivos / iCloud." }),
          !q
            ? h("button", { class: "oc-empty-btn", type: "button", text: "Nuevo documento en blanco", onclick: () => createFromTemplate("blank") })
            : null
        )
      );
    }
    el.replaceChildren(...nodes);
  }

  function renderStarredTab() {
    $("oc-app-head").replaceChildren(
      h("div", { class: "oc-head-row" }, h("h1", { class: "oc-head-title", text: "Destacados" }), h("button", { class: "oc-head-btn", type: "button", "aria-label": "Ajustes", icon: "settings", onclick: settingsSheet }))
    );
    const items = sorted(S.files.filter((m) => m.destacado));
    $("oc-tab-body").replaceChildren(
      items.length
        ? renderFileList(items)
        : h(
            "div",
            { class: "oc-empty" },
            h("div", { icon: "star" }),
            h("b", { text: "Sin documentos destacados" }),
            h("div", { text: "Toca la estrella en cualquier documento para fijarlo aquí." })
          )
    );
  }

  function renderFilesTab() {
    const folder = S.filesFolder ? S.folders.find((f) => f.id === S.filesFolder) : null;
    $("oc-app-head").replaceChildren(
      h(
        "div",
        { class: "oc-head-row" },
        S.filesFolder
          ? h("button", { class: "oc-head-back", type: "button", icon: "back", onclick: () => { S.filesFolder = null; renderFilesTab(); } }, h("span", { text: "Abrir" }))
          : null,
        h("h1", { class: "oc-head-title", text: folder ? folder.nombre : "Abrir" }),
        !S.filesFolder ? h("button", { class: "oc-head-btn", type: "button", "aria-label": "Nueva carpeta", icon: "folderPlus", onclick: () => newFolderSheet(null) }) : null,
        h("button", {
          class: "oc-head-btn",
          type: "button",
          "aria-label": "Cambiar vista",
          icon: prefs.vistaArchivos === "lista" ? "gridView" : "listView",
          onclick: () => {
            prefs.vistaArchivos = prefs.vistaArchivos === "lista" ? "cuadricula" : "lista";
            savePrefs();
            renderFilesTab();
          },
        })
      ),
      makeSearchBox("Buscar en archivos…", S.query, (q) => { S.query = q; renderFilesBody(); }, false)
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
          locRow("open", "#185abd", "Archivos / iCloud Drive", "Abrir .docx, .doc, .odt, .rtf o .md", pickFile),
          locRow("iphone", "#2b8a3e", "En este iPhone", S.files.length + " documentos · " + sizeText(total), () => {
            const a = el.querySelector("#oc-all-docs");
            if (a) a.scrollIntoView({ behavior: "smooth" });
          })
        )
      );
      if (S.folders.length) {
        nodes.push(
          h("div", { class: "oc-section-head" }, h("h2", { text: "Carpetas" }), h("button", { class: "oc-link-btn", type: "button", text: "+ Nueva", onclick: () => newFolderSheet(null) })),
          h(
            "div",
            { class: "oc-loc-card" },
            S.folders.map((f) => {
              const cnt = S.files.filter((m) => m.carpeta === f.id).length;
              return locRow("folder", "#5a6375", f.nombre, cnt + (cnt === 1 ? " documento" : " documentos"), () => {
                S.filesFolder = f.id;
                renderFilesTab();
              });
            })
          )
        );
      }
      nodes.push(h("div", { class: "oc-section-head", id: "oc-all-docs" }, h("h2", { text: "Todos los documentos" }), h("small", { text: String(S.files.length) })));
    }

    let items = S.files;
    if (S.filesFolder) items = items.filter((m) => m.carpeta === S.filesFolder);
    if (q) items = items.filter((m) => (m.nombre || "").toLowerCase().includes(q));
    items = sorted(items);

    if (items.length) nodes.push(prefs.vistaArchivos === "cuadricula" ? renderFileGrid(items) : renderFileList(items));
    else nodes.push(h("div", { class: "oc-empty" }, h("div", { icon: "folder" }), h("b", { text: "Sin archivos aquí" })));
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
        if (meta.miniatura) thumb.append(h("img", { src: meta.miniatura, alt: "" }));
        else {
          thumb.insertAdjacentHTML("afterbegin", icon("file"));
          thumb.append(h("span", { class: "oc-thumb-tag", text: "DOCX" }));
        }
        return h(
          "li",
          { class: "oc-recent" },
          h(
            "button",
            { class: "oc-recent-open", type: "button", onclick: () => openStoredDoc(meta) },
            thumb,
            h(
              "span",
              { class: "oc-recent-text" },
              h("b", { text: meta.nombre }),
              h("small", { text: describe(meta, meta.paginas ? meta.paginas + " pág." : "") })
            )
          ),
          h("button", {
            class: "oc-star-btn" + (meta.destacado ? " oc-starred" : ""),
            type: "button",
            "aria-label": "Destacar",
            [meta.destacado ? "iconFilled" : "icon"]: "star",
            onclick: () => toggleStar(meta),
          }),
          h("button", { class: "oc-recent-more", type: "button", "aria-label": "Más", icon: "more", onclick: () => fileMenu(meta) })
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
        else prev.insertAdjacentHTML("afterbegin", icon("file"));
        return h(
          "div",
          { class: "oc-file-card" },
          h("button", { class: "oc-file-card-tap", type: "button", onclick: () => openStoredDoc(meta) }, prev, h("div", { class: "oc-file-card-meta" }, h("b", { text: meta.nombre }), h("small", { text: describe(meta) }))),
          h(
            "div",
            { class: "oc-file-card-actions" },
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
          { i: "open", t: "Abrir", run: () => { close(); openStoredDoc(meta); } },
          { i: "star", t: meta.destacado ? "Quitar de destacados" : "Destacar", run: () => { close(); toggleStar(meta); } },
          { i: "share", t: "Compartir / Guardar en Archivos", run: async () => { close(); const b = await store.blob(meta.id); if (b) shareFile(new File([b], meta.nombre, { type: window.ocMimeFor(meta.nombre) })); } },
          { i: "rename", t: "Renombrar", run: () => { close(); renameSheet(meta); } },
          { i: "copy", t: "Duplicar", run: () => { close(); duplicateDoc(meta); } },
          { i: "folder", t: "Mover a carpeta…", run: () => { close(); moveFolderSheet(meta); } },
        ]),
        list([{ i: "trash", t: "Eliminar de este iPhone", danger: true, run: () => { close(); confirmDelete(meta); } }]),
      ],
    });
  }

  function renameSheet(meta) {
    const base = (meta.nombre || "Documento.docx").replace(/\.(docx|odt|rtf|md|txt)$/i, "");
    const inp = h("input", { class: "oc-field", type: "text", value: base, placeholder: "Nombre del documento" });
    sheet({
      title: "Renombrar documento",
      build: (close) => {
        const save = async () => {
          const clean = inp.value.trim();
          if (!clean) return;
          const extMatch = (meta.nombre || ".docx").match(/\.[a-z0-9]+$/i);
          const ext = extMatch ? extMatch[0] : ".docx";
          meta.nombre = clean.replace(/\.[a-z0-9]+$/i, "") + ext;
          await store.putMeta(meta).catch(() => {});
          if (S.doc && S.doc.meta.id === meta.id) { S.doc.meta.nombre = meta.nombre; renderTop(); }
          await reloadLibrary();
          if (!body.classList.contains("oc-in-editor")) renderShell();
          close();
        };
        inp.addEventListener("keydown", (e) => { if (e.key === "Enter") save(); });
        return [inp, h("button", { class: "oc-primary", type: "button", text: "Guardar nombre", onclick: save })];
      },
    });
    inp.focus();
    inp.select();
  }

  async function duplicateDoc(meta) {
    const b = await store.blob(meta.id);
    if (!b) return;
    const copyName = meta.nombre.replace(/(\.[^.]+)$/, " (copia)$1");
    await store.save(Object.assign({}, meta, { id: store.newId(), nombre: copyName, abierto: Date.now(), modificado: Date.now(), destacado: false }), b);
    await reloadLibrary();
    renderShell();
    toast("Copia creada");
  }

  function newFolderSheet(moveMeta) {
    const inp = h("input", { class: "oc-field", type: "text", placeholder: "Ej. Universidad, Trabajo…" });
    sheet({
      title: "Nueva carpeta",
      build: (close) => {
        const go = async () => {
          const name = inp.value.trim();
          if (!name) return;
          const f = { id: store.newId(), app: "word", nombre: name, creada: Date.now() };
          await store.saveFolder(f);
          if (moveMeta) { moveMeta.carpeta = f.id; await store.putMeta(moveMeta); }
          await reloadLibrary();
          renderShell();
          close();
        };
        inp.addEventListener("keydown", (e) => { if (e.key === "Enter") go(); });
        return [inp, h("button", { class: "oc-primary", type: "button", text: "Crear carpeta", onclick: go })];
      },
    });
    inp.focus();
  }

  function moveFolderSheet(meta) {
    sheet({
      title: "Mover a carpeta",
      build: (close) => [
        list([
          { i: "iphone", t: "Sin carpeta", small: !meta.carpeta ? "Actual" : "", run: async () => { close(); delete meta.carpeta; await store.putMeta(meta); await reloadLibrary(); renderShell(); } },
          ...S.folders.map((f) => ({ i: "folder", t: f.nombre, small: meta.carpeta === f.id ? "Actual" : "", run: async () => { close(); meta.carpeta = f.id; await store.putMeta(meta); await reloadLibrary(); renderShell(); } })),
        ]),
        h("button", { class: "oc-secondary", type: "button", icon: "folderPlus", text: "Nueva carpeta", onclick: () => { close(); newFolderSheet(meta); } }),
      ],
    });
  }

  function confirmDelete(meta) {
    sheet({
      title: "¿Eliminar documento?",
      text: "Se borrará «" + meta.nombre + "» de WordCraft en este iPhone.",
      build: (close) => [list([{ i: "trash", t: "Eliminar", danger: true, run: async () => { close(); await store.remove(meta.id); await reloadLibrary(); renderShell(); toast("Eliminado"); } }])],
    });
  }

  function sortSheet() {
    sheet({
      title: "Ordenar documentos",
      build: (close) => [
        list([
          ["reciente", "Más reciente primero"],
          ["nombre", "Nombre (A – Z)"],
          ["tamano", "Tamaño"],
        ].map(([v, label]) => ({
          i: prefs.orden === v ? "check" : "sort",
          t: label,
          run: () => { prefs.orden = v; savePrefs(); close(); renderShell(); },
        }))),
      ],
    });
  }

  function openCreateSheet() {
    sheet({
      title: "Nuevo documento",
      build: (close) => [
        list([
          { i: "blank", t: "Documento en blanco", small: ".docx", run: () => { close(); createFromTemplate("blank"); } },
          { i: "open", t: "Abrir desde Archivos / iCloud", small: ".docx, .odt, .rtf", run: () => { close(); pickFile(); } },
          { i: "file", t: "Currículum vitae", small: "Plantilla", run: () => { close(); createFromTemplate("resume"); } },
          { i: "file", t: "Informe profesional", small: "Plantilla", run: () => { close(); createFromTemplate("report"); } },
          { i: "file", t: "Carta formal", small: "Plantilla", run: () => { close(); createFromTemplate("letter"); } },
          { i: "file", t: "Apuntes de clase", small: "Plantilla", run: () => { close(); createFromTemplate("notes"); } },
        ]),
      ],
    });
  }

  function pickFile() {
    const input = $("oc-file");
    input.value = "";
    input.click();
  }

  // ---------- Abrir / Crear documento ----------
  async function openStoredDoc(meta) {
    const blob = await store.blob(meta.id);
    if (!blob) {
      toast("Ese archivo ya no está disponible.");
      await store.remove(meta.id);
      await reloadLibrary();
      renderShell();
      return;
    }
    meta.abierto = Date.now();
    store.putMeta(meta).catch(() => {});
    openDocument(meta, blob);
  }

  async function createFromTemplate(tplId) {
    if (!S.ready) { toast("WordCraft todavía está iniciando"); return; }
    S.opening = true;
    showEditor();
    busy(true, "Creando documento…");
    try {
      await closeEngineDocs();
      const stamp = new Date().toLocaleDateString("es", { day: "2-digit", month: "2-digit" }).replace("/", "-");
      const titles = {
        blank: "Documento " + stamp + ".docx",
        resume: "Currículum " + stamp + ".docx",
        report: "Informe " + stamp + ".docx",
        letter: "Carta " + stamp + ".docx",
        notes: "Apuntes " + stamp + ".docx",
        sample: "Ejemplo WordCraft.docx",
      };
      if (tplId === "notes") {
        await wc.run("file.new", {});
        await wc.run("para.heading1");
        await wc.run("text.insert", { text: "Apuntes — " + stamp + "\n" });
        await wc.run("para.normal");
        await wc.run("text.insert", { text: "Tema principal:\n" });
        await wc.run("para.bullets");
        await wc.run("text.insert", { text: "Idea clave 1\nIdea clave 2\n" });
      } else if (tplId === "blank") {
        await wc.run("file.new", {});
      } else {
        await wc.run("file.new", { template: tplId }).catch(() => wc.run("file.new", {}));
      }
      await sleep(160);
      const u = await wc.ui();
      const meta = {
        id: store.newId(),
        app: "word",
        nombre: titles[tplId] || "Documento.docx",
        tamano: 0,
        abierto: Date.now(),
        modificado: Date.now(),
        paginas: Math.max(1, (u.pages || []).length),
      };
      S.doc = {
        id: u.document,
        meta,
        name: meta.nombre,
        dirty: true,
        dirtySince: Date.now(),
        pages: meta.paginas,
        page: 1,
        zoom: 100,
        mobileView: prefs.vista === "movil",
      };
      await prepareEditorView();
      await saveActive().catch(() => {});
      busy(false);
      renderTop();
      renderBottom();
      setTimeout(() => captureThumb(meta), 800);
    } catch (e) {
      busy(false);
      toast("No se pudo crear el documento: " + e.message, 4500);
      S.doc = null;
      showHome();
    } finally {
      S.opening = false;
    }
  }

  async function openDocument(meta, blob) {
    if (!S.ready) { toast("WordCraft todavía está iniciando"); return; }
    S.opening = true;
    showEditor();
    renderTop(meta.nombre);
    busy(true, "Abriendo " + meta.nombre + "…");
    try {
      await closeEngineDocs();
      const before = new Set(((await wc.ui()).documents || []).map((d) => d.id));
      const bytes = await blob.arrayBuffer();
      canvas.dispatchEvent(new CustomEvent("wordcraft:open", { detail: { name: meta.nombre, bytes } }));
      let docId = null;
      for (let i = 0; i < 120; i++) {
        await sleep(100);
        const u = await wc.ui();
        const fresh = (u.documents || []).find((d) => !before.has(d.id));
        if (fresh) {
          docId = fresh.id;
          break;
        }
      }
      const u = await wc.ui();
      const pages = Math.max(1, (u.pages || []).length);
      meta.paginas = pages;
      store.putMeta(meta).catch(() => {});
      S.doc = {
        id: docId != null ? docId : u.document,
        meta,
        name: meta.nombre,
        dirty: false,
        dirtySince: 0,
        pages,
        page: 1,
        zoom: 100,
        mobileView: prefs.vista === "movil",
      };
      await prepareEditorView();
      busy(false);
      renderTop();
      renderBottom();
      if (!meta.miniatura) setTimeout(() => captureThumb(meta), 900);
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
    await wc.run("ui.backstage", { value: false }).catch(() => {});
    await wc.run("ui.collapseRibbon", { value: true }).catch(() => {});
    await wc.run("view.ruler", { value: false }).catch(() => {});
    if (prefs.vista === "movil") {
      await wc.run("view.webLayout").catch(() => {});
    } else {
      await wc.run("view.printLayout").catch(() => {});
    }
    if (prefs.paginaOscura) {
      await wc.run("view.darkMode", { value: true }).catch(() => {});
    }
    await sleep(120);
    const u = await wc.ui();
    applyCanvasCrop(u);
    await sleep(80);
    if (prefs.vista !== "movil") {
      await fitPageWidth();
    } else {
      await wc.run("view.zoom", { value: 115 }).catch(() => {});
      S.zoomPct = 115;
      S.fitZoomPct = 115;
    }
    const u2 = await wc.ui().catch(() => null);
    if (u2) applyCanvasCrop(u2);
    updatePagePill();
    window.ocSyncWakeLock(prefs.pantallaActiva);
  }

  async function closeEngineDocs() {
    for (let i = 0; i < 6; i++) {
      const u = await wc.ui().catch(() => null);
      if (!u || !u.documents || !u.documents.length) break;
      await wc.run("file.close", { force: true }).catch(() => {});
      await sleep(80);
    }
  }

  async function captureThumb(meta) {
    try {
      if (!S.doc || S.doc.meta !== meta) return;
      const res = await wc.renderPage(1, 0.35);
      if (!res || !res.png) return;
      const img = new Image();
      await new Promise((resOk, rej) => { img.onload = resOk; img.onerror = rej; img.src = "data:image/png;base64," + res.png; });
      const c = document.createElement("canvas");
      c.width = 160;
      c.height = Math.min(220, Math.round((160 * img.height) / Math.max(1, img.width)));
      c.getContext("2d").drawImage(img, 0, 0, c.width, c.height);
      meta.miniatura = c.toDataURL("image/jpeg", 0.72);
      await store.putMeta(meta);
    } catch (e) {}
  }

  async function goHome() {
    if (S.doc && S.doc.dirty) {
      busy(true, "Guardando cambios…");
      await saveActive().catch(() => {});
      await captureThumb(S.doc.meta).catch(() => {});
      busy(false);
    }
    closeDrawer();
    S.doc = null;
    await closeEngineDocs();
    showHome();
  }

  function showEditor() {
    body.classList.add("oc-in-editor");
    renderBottom();
  }

  async function showHome() {
    body.classList.remove("oc-in-editor", "oc-desktop", "oc-modal", "oc-immersive");
    S.desktop = false;
    S.modal = false;
    S.searching = false;
    closeDrawer();
    window.ocSyncWakeLock(false);
    if (window.ocUpdateReady) {
      location.reload();
      return;
    }
    await reloadLibrary();
    renderShell();
  }

  function busy(on, text) {
    $("oc-busy").hidden = !on;
    if (text) $("oc-busy-text").textContent = text;
  }

  // ---------- Barra superior (Microsoft Word iPhone) ----------
  function renderTop(fallbackName) {
    const top = $("oc-top");
    if (S.searching) {
      const inp = h("input", {
        type: "search",
        placeholder: "Buscar en el documento…",
        value: S.searchQuery,
        enterkeyhint: "search",
      });
      inp.addEventListener("keydown", (e) => {
        if (e.key === "Enter") {
          e.preventDefault();
          S.searchQuery = inp.value;
          runFind(inp.value);
        }
      });
      top.replaceChildren(
        h("label", { class: "oc-search" }, h("span", { icon: "search" }), inp),
        h("button", { class: "oc-text-btn", type: "button", text: "Buscar", onclick: () => { S.searchQuery = inp.value; runFind(inp.value); } }),
        h("button", { class: "oc-text-btn", type: "button", text: "Listo", onclick: () => { S.searching = false; renderTop(); } })
      );
      setTimeout(() => inp.focus(), 20);
      return;
    }

    const title = S.doc ? S.doc.meta.nombre : fallbackName || "WordCraft";
    const sub = S.doc ? (S.doc.dirty ? "Editando…" : "Guardado en iPhone") : "";
    const isMob = !!(S.doc && S.doc.mobileView);

    top.replaceChildren(
      h("button", { class: "oc-icon-btn", type: "button", "aria-label": "Volver", icon: "back", onclick: goHome }),
      h("button", { class: "oc-title", type: "button", onclick: () => { if (S.doc) renameSheet(S.doc.meta); } }, h("b", { text: title }), h("small", { text: sub })),
      h("button", {
        class: "oc-icon-btn" + (isMob ? " oc-on" : ""),
        type: "button",
        "aria-label": isMob ? "Cambiar a Diseño de impresión" : "Cambiar a Vista móvil",
        title: isMob ? "Vista móvil activa" : "Diseño de impresión",
        icon: isMob ? "mobileView" : "printLayout",
        onclick: toggleMobileView,
      }),
      h("button", { class: "oc-icon-btn", type: "button", "aria-label": "Deshacer", icon: "undo", onclick: () => execCmd("edit.undo") }),
      h("button", { class: "oc-icon-btn", type: "button", "aria-label": "Buscar", icon: "search", onclick: () => { S.searching = true; renderTop(); } }),
      h("button", { class: "oc-icon-btn", type: "button", "aria-label": "Compartir", icon: "share", onclick: shareMenu }),
      h("button", { class: "oc-icon-btn", type: "button", "aria-label": "Más opciones", icon: "more", onclick: moreMenu })
    );
  }

  async function toggleMobileView() {
    if (!S.doc) return;
    S.doc.mobileView = !S.doc.mobileView;
    prefs.vista = S.doc.mobileView ? "movil" : "impresion";
    savePrefs();
    if (S.doc.mobileView) {
      await wc.run("view.webLayout").catch(() => {});
      await wc.run("view.zoom", { value: 115 }).catch(() => {});
      S.zoomPct = 115;
      toast("Vista móvil: texto adaptado al ancho del iPhone");
    } else {
      await wc.run("view.printLayout").catch(() => {});
      await sleep(80);
      await fitPageWidth();
      toast("Diseño de impresión (páginas completas)");
    }
    renderTop();
    updatePagePill();
  }

  async function runFind(q) {
    const query = String(q || "").trim();
    if (!query) return;
    try {
      await wc.run("select.text", { text: query });
      toast("Coincidencia seleccionada");
    } catch (e) {
      await wc.run("ui.dialog", { name: "find" }).catch(() => toast("No se encontró «" + query + "»"));
    }
  }

  // ---------- Barra inferior rápida + Cinta móvil (exacta a Word iPhone) ----------
  function renderBottom() {
    const bar = $("oc-bottom");
    const items = [
      { id: "bold", icon: "bold", label: "Negrita", cmd: "format.bold", key: "bold" },
      { id: "italic", icon: "italic", label: "Cursiva", cmd: "format.italic", key: "italic" },
      { id: "underline", icon: "underline", label: "Subrayado", cmd: "format.underline", key: "underline" },
      { id: "highlight", icon: "highlight", label: "Resaltar", fn: highlightPicker },
      { id: "color", icon: "textColor", label: "Color", fn: fontColorPicker },
      { id: "bullets", icon: "bullets", label: "Viñetas", cmd: "para.bullets" },
      { id: "numbering", icon: "numbering", label: "Numeración", cmd: "para.numbering" },
      { id: "image", icon: "image", label: "Imagen", fn: () => $("oc-img").click() },
      { id: "comment", icon: "comment", label: "Comentario", fn: addCommentSheet },
      { id: "redo", icon: "redo", label: "Rehacer", cmd: "edit.redo" },
    ];
    bar.replaceChildren(
      h(
        "div",
        { class: "oc-quickbar" },
        h(
          "div",
          { class: "oc-quick-scroll", id: "oc-quick-scroll" },
          items.map((it) =>
            h("button", {
              class: "oc-qbtn",
              type: "button",
              "data-fmt": it.key || "",
              "aria-label": it.label,
              "aria-pressed": it.key ? String(!!S.fmt[it.key]) : null,
              icon: it.icon,
              onclick: () => (it.fn ? it.fn() : execCmd(it.cmd)),
            })
          )
        ),
        h("button", {
          class: "oc-drawer-toggle",
          type: "button",
          "aria-label": S.drawerOpen ? "Cerrar menú de herramientas" : "Abrir cinta de Word",
          icon: S.drawerOpen ? "chevDown" : "chevUp",
          onclick: toggleDrawer,
        })
      )
    );
  }

  function updateQuickBarState() {
    document.querySelectorAll("#oc-quick-scroll [data-fmt]").forEach((btn) => {
      const k = btn.getAttribute("data-fmt");
      if (!k) return;
      const val = k === "underline" ? S.fmt.underline && S.fmt.underline !== "none" : !!S.fmt[k];
      btn.setAttribute("aria-pressed", String(val));
    });
  }

  function toggleDrawer() {
    S.drawerOpen = !S.drawerOpen;
    renderDrawer();
    renderBottom();
    setTimeout(() => { wc.ui().then(applyCanvasCrop).catch(() => {}); }, 80);
  }
  function closeDrawer() {
    if (!S.drawerOpen) return;
    S.drawerOpen = false;
    $("oc-ribbon-drawer").hidden = true;
    renderBottom();
    setTimeout(() => { wc.ui().then(applyCanvasCrop).catch(() => {}); }, 80);
  }

  function renderDrawer() {
    const drawer = $("oc-ribbon-drawer");
    drawer.hidden = !S.drawerOpen;
    if (!S.drawerOpen) return;
    const tabs = [
      ["inicio", "Inicio"],
      ["insertar", "Insertar"],
      ["disposicion", "Disposición"],
      ["revisar", "Revisar"],
      ["vista", "Vista"],
    ];
    drawer.replaceChildren(
      h(
        "div",
        { class: "oc-ribbon-tabs", role: "tablist" },
        tabs.map(([id, label]) =>
          h("button", {
            class: "oc-rtab",
            type: "button",
            role: "tab",
            "aria-selected": String(S.drawerTab === id),
            text: label,
            onclick: () => { S.drawerTab = id; renderDrawer(); },
          })
        ),
        h("button", { class: "oc-ribbon-close", type: "button", "aria-label": "Cerrar", icon: "close", onclick: closeDrawer })
      ),
      h("div", { class: "oc-ribbon-body" }, buildDrawerContent(S.drawerTab))
    );
  }

  function buildDrawerContent(tab) {
    if (tab === "inicio") {
      return [
        h(
          "div",
          { class: "oc-rgroup" },
          h("div", { class: "oc-rgroup-title", text: "Fuente y tamaño" }),
          h(
            "div",
            { class: "oc-rrow" },
            h("button", { class: "oc-rbtn", type: "button", text: S.fmt.font || "Calibri", onclick: fontPicker }),
            h("button", { class: "oc-rbtn", type: "button", text: (S.fmt.size ? Math.round(S.fmt.size) : 11) + " pt", onclick: fontSizePicker }),
            h("button", { class: "oc-rbtn", type: "button", text: "A-", onclick: () => execCmd("format.shrinkFont") }),
            h("button", { class: "oc-rbtn", type: "button", text: "A+", onclick: () => execCmd("format.growFont") }),
            h("button", { class: "oc-rbtn", type: "button", icon: "bold", onclick: () => execCmd("format.bold") }),
            h("button", { class: "oc-rbtn", type: "button", icon: "italic", onclick: () => execCmd("format.italic") }),
            h("button", { class: "oc-rbtn", type: "button", icon: "underline", onclick: () => execCmd("format.underline") }),
            h("button", { class: "oc-rbtn", type: "button", icon: "strike", onclick: () => execCmd("format.strikethrough") }),
            h("button", { class: "oc-rbtn", type: "button", text: "x²", onclick: () => execCmd("format.superscript") }),
            h("button", { class: "oc-rbtn", type: "button", text: "x₂", onclick: () => execCmd("format.subscript") }),
            h("button", { class: "oc-rbtn", type: "button", icon: "highlight", text: "Resaltar", onclick: highlightPicker }),
            h("button", { class: "oc-rbtn", type: "button", icon: "textColor", text: "Color", onclick: fontColorPicker }),
            h("button", { class: "oc-rbtn", type: "button", icon: "clearFmt", text: "Borrar", onclick: () => execCmd("format.clear") })
          )
        ),
        h(
          "div",
          { class: "oc-rgroup" },
          h("div", { class: "oc-rgroup-title", text: "Párrafo y alineación" }),
          h(
            "div",
            { class: "oc-rrow" },
            h("button", { class: "oc-rbtn", type: "button", icon: "alignLeft", onclick: () => execCmd("para.alignLeft") }),
            h("button", { class: "oc-rbtn", type: "button", icon: "alignCenter", onclick: () => execCmd("para.alignCenter") }),
            h("button", { class: "oc-rbtn", type: "button", icon: "alignRight", onclick: () => execCmd("para.alignRight") }),
            h("button", { class: "oc-rbtn", type: "button", icon: "alignJustify", onclick: () => execCmd("para.justify") }),
            h("button", { class: "oc-rbtn", type: "button", icon: "bullets", onclick: () => execCmd("para.bullets") }),
            h("button", { class: "oc-rbtn", type: "button", icon: "numbering", onclick: () => execCmd("para.numbering") }),
            h("button", { class: "oc-rbtn", type: "button", icon: "outdent", onclick: () => execCmd("para.outdent") }),
            h("button", { class: "oc-rbtn", type: "button", icon: "indent", onclick: () => execCmd("para.indent") }),
            h("button", { class: "oc-rbtn", type: "button", text: "1.0×", onclick: () => execCmd("para.single") }),
            h("button", { class: "oc-rbtn", type: "button", text: "1.5×", onclick: () => execCmd("para.oneAndHalf") }),
            h("button", { class: "oc-rbtn", type: "button", text: "2.0×", onclick: () => execCmd("para.double") })
          )
        ),
        h(
          "div",
          { class: "oc-rgroup" },
          h("div", { class: "oc-rgroup-title", text: "Estilos de Word" }),
          h(
            "div",
            { class: "oc-rrow" },
            h("button", { class: "oc-rbtn", type: "button", text: "Normal", onclick: () => execCmd("para.normal") }),
            h("button", { class: "oc-rbtn", type: "button", text: "Título 1", onclick: () => execCmd("para.heading1") }),
            h("button", { class: "oc-rbtn", type: "button", text: "Título 2", onclick: () => execCmd("para.heading2") }),
            h("button", { class: "oc-rbtn", type: "button", text: "Título 3", onclick: () => execCmd("para.heading3") }),
            h("button", { class: "oc-rbtn", type: "button", text: "Todo", onclick: () => execCmd("select.all") })
          )
        ),
      ];
    }

    if (tab === "insertar") {
      return [
        h(
          "div",
          { class: "oc-rgrid-3" },
          rtile("table", "Tabla", insertTableSheet),
          rtile("image", "Imagen", () => { closeDrawer(); $("oc-img").click(); }),
          rtile("pageBreak", "Salto pág.", () => execCmd("insert.pageBreak")),
          rtile("blank", "Pág. blanca", () => execCmd("insert.blankPage")),
          rtile("textBox", "Cuadro texto", () => execCmd("insert.textBox")),
          rtile("shape", "Forma", insertShapeSheet),
          rtile("link", "Vínculo", insertLinkSheet),
          rtile("comment", "Comentario", addCommentSheet),
          rtile("file", "Núm. página", () => execCmd("insert.pageNumber"))
        ),
      ];
    }

    if (tab === "disposicion") {
      return [
        h(
          "div",
          { class: "oc-rgroup" },
          h("div", { class: "oc-rgroup-title", text: "Orientación y columnas" }),
          h(
            "div",
            { class: "oc-rrow" },
            h("button", { class: "oc-rbtn", type: "button", text: "Vertical", onclick: () => execCmd("layout.orientation", { orientation: "portrait" }) }),
            h("button", { class: "oc-rbtn", type: "button", text: "Horizontal", onclick: () => execCmd("layout.orientation", { orientation: "landscape" }) }),
            h("button", { class: "oc-rbtn", type: "button", text: "1 col", onclick: () => execCmd("layout.columns", { count: 1 }) }),
            h("button", { class: "oc-rbtn", type: "button", text: "2 cols", onclick: () => execCmd("layout.columns", { count: 2 }) }),
            h("button", { class: "oc-rbtn", type: "button", text: "3 cols", onclick: () => execCmd("layout.columns", { count: 3 }) })
          )
        ),
        h(
          "div",
          { class: "oc-rgroup" },
          h("div", { class: "oc-rgroup-title", text: "Márgenes y tamaño" }),
          h(
            "div",
            { class: "oc-rrow" },
            h("button", { class: "oc-rbtn", type: "button", text: "Normal", onclick: () => execCmd("layout.margins", { preset: "normal" }) }),
            h("button", { class: "oc-rbtn", type: "button", text: "Estrecho", onclick: () => execCmd("layout.margins", { preset: "narrow" }) }),
            h("button", { class: "oc-rbtn", type: "button", text: "Ancho", onclick: () => execCmd("layout.margins", { preset: "wide" }) }),
            h("button", { class: "oc-rbtn", type: "button", text: "Carta", onclick: () => execCmd("layout.size", { preset: "letter" }) }),
            h("button", { class: "oc-rbtn", type: "button", text: "A4", onclick: () => execCmd("layout.size", { preset: "a4" }) })
          )
        ),
      ];
    }

    if (tab === "revisar") {
      return [
        h(
          "div",
          { class: "oc-rgrid-2" },
          rtile("info", "Contar palabras", showWordCount),
          rtile("comment", "Nuevo comentario", addCommentSheet),
          rtile("edit", "Control de cambios", () => execCmd("review.trackChanges")),
          rtile("check", "Aceptar cambios", () => execCmd("review.acceptAll")),
          rtile("close", "Rechazar cambios", () => execCmd("review.rejectAll")),
          rtile("search", "Ortografía", () => execCmd("review.spelling"))
        ),
      ];
    }

    return [
      h(
        "div",
        { class: "oc-rgrid-2" },
        rtile("mobileView", "Vista móvil", async () => { if (!S.doc.mobileView) await toggleMobileView(); }),
        rtile("printLayout", "Diseño impresión", async () => { if (S.doc.mobileView) await toggleMobileView(); }),
        rtile("moon", "Página oscura", async () => {
          prefs.paginaOscura = !prefs.paginaOscura;
          savePrefs();
          await execCmd("view.darkMode", { value: prefs.paginaOscura });
        }),
        rtile("zoom", "Ajustar al ancho", fitPageWidth),
        rtile("eye", "Modo lectura", () => execCmd("view.readMode")),
        rtile("desktop", "Modo PC completo", () => { closeDrawer(); setDesktop(true); })
      ),
    ];
  }

  function rtile(ic, label, onclick) {
    return h("button", { class: "oc-rtile", type: "button", icon: ic, onclick }, h("span", { text: label }));
  }

  async function execCmd(cmd, params) {
    try {
      const res = await wc.run(cmd, params || {});
      canvas.focus({ preventScroll: true });
      const st = await wc.run("format.state").catch(() => null);
      if (st) { S.fmt = st; updateQuickBarState(); }
      return res;
    } catch (e) {
      toast("No disponible ahora: " + (e.message || e), 3000);
      return null;
    }
  }

  // ---------- Selectores y hojas contextuales ----------
  function fontPicker() {
    const fonts = ["Calibri", "Arial", "Georgia", "Times New Roman", "Cambria", "Verdana", "Trebuchet MS", "Courier New"];
    sheet({
      title: "Tipo de fuente",
      build: (close) => [list(fonts.map((f) => ({ i: S.fmt.font === f ? "check" : "edit", t: f, run: () => { close(); execCmd("format.font", { font: f }); } })))],
    });
  }

  function fontSizePicker() {
    const sizes = [9, 10, 11, 12, 14, 16, 18, 22, 26, 32, 40];
    sheet({
      title: "Tamaño de fuente",
      build: (close) => [
        h("div", { class: "oc-rrow" }, sizes.map((sz) => h("button", { class: "oc-rbtn", type: "button", text: sz + " pt", onclick: () => { close(); execCmd("format.size", { size: sz }); } }))),
      ],
    });
  }

  function highlightPicker() {
    const colors = [
      ["yellow", "Amarillo", "#fff176"],
      ["brightGreen", "Verde", "#80e27e"],
      ["turquoise", "Cian", "#4dd0e1"],
      ["pink", "Rosa", "#ff80ab"],
      ["blue", "Azul", "#82b1ff"],
      ["red", "Rojo", "#ff8a80"],
      ["darkYellow", "Ámbar", "#ffd54f"],
      ["gray25", "Gris", "#cfd8dc"],
      ["none", "Sin resaltado", ""],
    ];
    sheet({
      title: "Color de resaltado",
      build: (close) => [colorGrid(colors, (v) => { close(); execCmd("format.highlight", { color: v }); })],
    });
  }

  function fontColorPicker() {
    const colors = [
      ["auto", "Automático", ""],
      ["#000000", "Negro", "#000000"],
      ["#595959", "Gris", "#595959"],
      ["#185ABD", "Azul Word", "#185ABD"],
      ["#C43E1C", "Rojo teja", "#C43E1C"],
      ["#D92D20", "Rojo", "#D92D20"],
      ["#107C41", "Verde", "#107C41"],
      ["#E8590C", "Naranja", "#E8590C"],
      ["#6741D9", "Morado", "#6741D9"],
      ["#0C8599", "Turquesa", "#0C8599"],
      ["#F59F00", "Dorado", "#F59F00"],
      ["#FFFFFF", "Blanco", "#FFFFFF"],
    ];
    sheet({
      title: "Color de fuente",
      build: (close) => [colorGrid(colors, (v) => { close(); execCmd("format.color", { color: v }); })],
    });
  }

  function insertTableSheet() {
    let rows = 3;
    let cols = 3;
    const rInp = h("input", { class: "oc-field", type: "number", inputmode: "numeric", value: "3", min: "1", max: "20" });
    const cInp = h("input", { class: "oc-field", type: "number", inputmode: "numeric", value: "3", min: "1", max: "12" });
    sheet({
      title: "Insertar tabla",
      build: (close) => [
        h("div", { class: "oc-btn-row" },
          h("button", { class: "oc-secondary", type: "button", text: "2 × 2", onclick: () => { close(); execCmd("insert.table", { rows: 2, cols: 2 }); } }),
          h("button", { class: "oc-secondary", type: "button", text: "3 × 3", onclick: () => { close(); execCmd("insert.table", { rows: 3, cols: 3 }); } }),
          h("button", { class: "oc-secondary", type: "button", text: "4 × 4", onclick: () => { close(); execCmd("insert.table", { rows: 4, cols: 4 }); } })
        ),
        h("div", { class: "oc-label", text: "Filas" }), rInp,
        h("div", { class: "oc-label", text: "Columnas" }), cInp,
        h("button", {
          class: "oc-primary",
          type: "button",
          text: "Insertar tabla",
          onclick: () => {
            rows = Math.max(1, Math.min(20, parseInt(rInp.value, 10) || 3));
            cols = Math.max(1, Math.min(12, parseInt(cInp.value, 10) || 3));
            close();
            execCmd("insert.table", { rows, cols });
          },
        }),
      ],
    });
  }

  function insertShapeSheet() {
    const shapes = [
      ["rect", "Rectángulo"],
      ["roundRect", "Rectángulo redondeado"],
      ["ellipse", "Óvalo / Círculo"],
      ["rightArrow", "Flecha derecha"],
      ["line", "Línea"],
      ["star5", "Estrella"],
    ];
    sheet({
      title: "Insertar forma",
      build: (close) => [list(shapes.map(([preset, t]) => ({ i: "shape", t, run: () => { close(); execCmd("insert.shape", { preset }); } })))],
    });
  }

  function insertLinkSheet() {
    const urlInp = h("input", { class: "oc-field", type: "url", placeholder: "https://ejemplo.com" });
    const txtInp = h("input", { class: "oc-field", type: "text", placeholder: "Texto visible (opcional)" });
    sheet({
      title: "Insertar vínculo",
      build: (close) => [
        h("div", { class: "oc-label", text: "Dirección web" }), urlInp,
        h("div", { class: "oc-label", text: "Texto" }), txtInp,
        h("button", {
          class: "oc-primary",
          type: "button",
          text: "Insertar",
          onclick: () => {
            const url = urlInp.value.trim();
            if (!url) return;
            close();
            execCmd("insert.link", txtInp.value.trim() ? { url, text: txtInp.value.trim() } : { url });
          },
        }),
      ],
    });
    urlInp.focus();
  }

  function addCommentSheet() {
    const inp = h("input", { class: "oc-field", type: "text", placeholder: "Escribe tu comentario…" });
    sheet({
      title: "Nuevo comentario",
      build: (close) => {
        const go = () => {
          const text = inp.value.trim();
          if (!text) return;
          close();
          execCmd("review.newComment", { text });
        };
        inp.addEventListener("keydown", (e) => { if (e.key === "Enter") go(); });
        return [inp, h("button", { class: "oc-primary", type: "button", text: "Añadir comentario", onclick: go })];
      },
    });
    inp.focus();
  }

  async function insertImageFromFile(file) {
    try {
      const buf = await file.arrayBuffer();
      let bin = "";
      const bytes = new Uint8Array(buf);
      for (let i = 0; i < bytes.length; i += 0x8000) {
        bin += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
      }
      const b64 = btoa(bin);
      await wc.run("insert.picture", { base64: b64, name: file.name || "imagen.png" });
      toast("Imagen insertada");
    } catch (e) {
      // Fallback mediante drop de archivo al lienzo si la firma difiere
      try {
        const dt = new DataTransfer();
        dt.items.add(file);
        canvas.dispatchEvent(new DragEvent("drop", { bubbles: true, cancelable: true, dataTransfer: dt }));
        toast("Imagen enviada al documento");
      } catch (e2) {
        toast("No se pudo insertar la imagen: " + e.message, 4000);
      }
    }
  }

  async function showWordCount() {
    try {
      const res = await wc.run("review.wordCount");
      const w = res && (res.words != null ? res.words : res.wordCount);
      const c = res && (res.characters != null ? res.characters : res.chars);
      const p = res && (res.pages != null ? res.pages : S.doc.pages);
      toast((w != null ? w + " palabras" : "Conteo abierto") + (c != null ? " · " + c + " caracteres" : "") + (p ? " · " + p + " pág." : ""), 4500);
    } catch (e) {
      toast("No se pudo contar palabras");
    }
  }

  function zoomAndPageSheet() {
    sheet({
      title: "Zoom y vista",
      build: (close) => [
        h("div", { class: "oc-btn-row" },
          h("button", { class: "oc-secondary", type: "button", text: "Ancho", onclick: () => { close(); fitPageWidth(); } }),
          h("button", { class: "oc-secondary", type: "button", text: "100%", onclick: () => { close(); S.zoomPct = 100; wc.run("view.zoom", { value: 100 }); updatePagePill(); } }),
          h("button", { class: "oc-secondary", type: "button", text: "150%", onclick: () => { close(); S.zoomPct = 150; wc.run("view.zoom", { value: 150 }); updatePagePill(); } }),
          h("button", { class: "oc-secondary", type: "button", text: "200%", onclick: () => { close(); S.zoomPct = 200; wc.run("view.zoom", { value: 200 }); updatePagePill(); } })
        ),
        list([
          { i: "mobileView", t: S.doc.mobileView ? "Cambiar a Diseño de impresión" : "Cambiar a Vista móvil", run: () => { close(); toggleMobileView(); } },
          { i: "moon", t: prefs.paginaOscura ? "Desactivar página oscura" : "Activar página oscura", run: () => { close(); prefs.paginaOscura = !prefs.paginaOscura; savePrefs(); execCmd("view.darkMode", { value: prefs.paginaOscura }); } },
        ]),
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
          { i: "share", t: "Compartir como Word (.docx)", small: "Guardar en Archivos", run: async () => { close(); busy(true, "Preparando .docx…"); try { const f = await currentFile(); busy(false); if (f) shareFile(f); } catch (e) { busy(false); toast(e.message); } } },
          { i: "pdf", t: "Exportar como PDF", small: ".pdf", run: () => { close(); execCmd("file.exportPdf"); toast("Generando PDF…"); } },
          { i: "print", t: "Imprimir (AirPrint)", run: () => { close(); execCmd("ui.printSubmit").catch(() => execCmd("file.print")); } },
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
          { i: "info", t: "Contar palabras", run: () => { close(); showWordCount(); } },
          { i: "pdf", t: "Exportar a PDF", run: () => { close(); execCmd("file.exportPdf"); } },
          { i: "desktop", t: "Interfaz completa (PC)", small: "Todas las opciones", run: () => { close(); setDesktop(true); } },
        ]),
      ],
    });
  }

  function settingsSheet() {
    const authorInp = h("input", { class: "oc-field", type: "text", value: prefs.autor, placeholder: "Tu nombre para comentarios y cambios" });
    authorInp.addEventListener("change", () => {
      prefs.autor = authorInp.value.trim().slice(0, 120);
      savePrefs();
      if (prefs.autor && S.ready) wc.run("file.setAuthor", { name: prefs.autor }).catch(() => {});
    });
    sheet({
      title: "Ajustes de WordCraft",
      build: (close) => [
        h("div", { class: "oc-label", text: "Apariencia" }),
        seg([["auto", "Sistema"], ["claro", "Claro"], ["oscuro", "Oscuro"]], () => prefs.tema, (v) => {
          prefs.tema = v;
          savePrefs();
          applyTheme();
          if (S.ready) wc.run("ui.theme", { value: v === "oscuro" ? "dark" : v === "claro" ? "light" : "system" }).catch(() => {});
        }),
        h("div", { class: "oc-label", text: "Vista al abrir documentos" }),
        seg([["movil", "Vista móvil"], ["impresion", "Diseño de impresión"]], () => prefs.vista, (v) => { prefs.vista = v; savePrefs(); }),
        h("div", { class: "oc-label", text: "Tu nombre de autor" }),
        authorInp,
        h("div", { class: "oc-list" }, swRow("Mantener pantalla encendida", () => prefs.pantallaActiva, (v) => { prefs.pantallaActiva = v; savePrefs(); })),
        list([
          { i: "refresh", t: "Buscar actualizaciones", run: () => { close(); window.ocCheckUpdate(); } },
          { i: "apps", t: "Ir al portal OfficeCraft", small: "Word · PowerPoint · Excel", run: () => { location.href = "../"; } },
        ]),
        h("p", { text: "WordCraft para iPhone · Motor de código abierto (github.com/storytold/wordcraft) " + (S.appVersion ? "v" + S.appVersion : "") + ". Tus archivos no salen del teléfono." }),
      ],
    });
  }

  let backBtn = null;
  async function setDesktop(on) {
    S.desktop = on;
    body.classList.toggle("oc-desktop", on);
    if (on) {
      closeDrawer();
      await wc.run("ui.collapseRibbon", { value: false }).catch(() => {});
      await wc.run("ui.interfaceSize", { percent: 80 }).catch(() => {});
      applyCanvasCrop(null);
      backBtn = backBtn || h("button", { class: "oc-back-mobile", type: "button", text: "Volver a vista iPhone", onclick: () => setDesktop(false) });
      body.append(backBtn);
    } else {
      if (backBtn) backBtn.remove();
      await wc.run("ui.collapseRibbon", { value: true }).catch(() => {});
      await wc.run("ui.interfaceSize", { value: "auto" }).catch(() => {});
      await sleep(120);
      applyCanvasCrop(await wc.ui().catch(() => null));
    }
  }

  window.ocOnUpdate = () => {
    if (!body.classList.contains("oc-in-editor")) location.reload();
    else toast("Hay una versión nueva de WordCraft: se aplicará al volver a Inicio.", 5500);
  };

  window.wordcraftMobile = { openCreateSheet, createFromTemplate, toggleMobileView, toggleDrawer, setDesktop, goHome, state: () => S };
})();
