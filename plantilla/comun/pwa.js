// OfficeCraft para iPhone: núcleo compartido para WordCraft, DeckCraft y GridCraft.
// Maneja pantalla de carga, service worker offline con actualización automática,
// hoja de compartir de iOS, almacenamiento local IndexedDB compartido, íconos y hojas modales.
(function () {
  "use strict";

  const CFG = window.OC_APP || { id: "app", name: "OfficeCraft", ext: ".bin" };
  const APP_NAME = CFG.name || "OfficeCraft";
  const splash = document.getElementById("oc-splash");
  const bar = document.getElementById("oc-bar");
  const statusEl = document.getElementById("oc-status");
  const errorBox = document.getElementById("oc-error");
  const SIN_RED =
    "Revisa tu conexión a internet y vuelve a abrir la app. " +
    "Después de la primera instalación ya no hace falta conexión.";

  function setStatus(text) {
    if (statusEl) statusEl.textContent = text;
  }
  function setProgress(fraction) {
    if (bar) bar.style.width = Math.max(3, Math.min(100, fraction * 100)).toFixed(1) + "%";
  }
  function showError(text) {
    if (!splash || splash.hidden) {
      console.error("[" + APP_NAME + "]", text);
      return;
    }
    if (errorBox) {
      errorBox.hidden = false;
      errorBox.textContent = text;
    }
    setStatus("No se pudo abrir " + APP_NAME);
  }
  function hideSplash() {
    if (!splash || splash.hidden) return;
    splash.classList.add("oc-out");
    setTimeout(() => {
      splash.hidden = true;
    }, 340);
  }

  function toast(text, ms) {
    document.querySelectorAll(".oc-toast").forEach((n) => n.remove());
    const el = document.createElement("div");
    el.className = "oc-toast";
    el.setAttribute("role", "status");
    el.textContent = text;
    document.body.appendChild(el);
    const dur = ms || 3000;
    setTimeout(() => {
      el.style.opacity = "0";
    }, dur);
    setTimeout(() => el.remove(), dur + 380);
  }

  window.ocToast = toast;
  window.ocHideSplash = hideSplash;
  window.ocShowError = showError;

  window.addEventListener(
    "error",
    (e) => {
      const t = e && e.target;
      if (t && t !== window && (t.src || t.href)) {
        if (t.tagName === "IMG" || t.tagName === "LINK") return;
        showError("No se pudo descargar " + (t.src || t.href).split("/").pop() + ".\n" + SIN_RED);
        return;
      }
      if (splash && !splash.hidden) {
        showError((e && e.message) || "Error desconocido al cargar.");
      }
    },
    true
  );
  window.addEventListener("unhandledrejection", (e) => {
    const r = e && e.reason;
    if (splash && !splash.hidden) {
      showError((r && (r.message || String(r))) || "Error desconocido al cargar.");
    }
  });

  // ---------- Service worker + actualizaciones automáticas ----------
  const sw = "serviceWorker" in navigator && window.isSecureContext ? navigator.serviceWorker : null;
  const hadController = !!(sw && sw.controller);
  let registration = null;
  let installFailed = null;
  let lastUpdateCheck = Date.now();

  if (sw) {
    sw.addEventListener("message", (event) => {
      const m = event.data || {};
      if (m.type === "progress" && m.total && splash && !splash.hidden) {
        setProgress(m.loaded / m.total);
        const pct = Math.min(100, Math.round((m.loaded / m.total) * 100));
        setStatus("Instalando " + APP_NAME + " en este iPhone… " + pct + "%");
      } else if (m.type === "install-failed") {
        installFailed = m.message || "error";
      } else if (m.type === "activated" && hadController) {
        window.ocUpdateReady = true;
        if (window.ocOnUpdate) window.ocOnUpdate(m.version || "");
        else toast("Nueva versión de " + APP_NAME + " lista. Se aplicará al volver a Inicio.", 5500);
      }
    });
    sw.register("./sw.js", { scope: "./" })
      .then((reg) => {
        registration = reg;
      })
      .catch((e) => {
        installFailed = String((e && e.message) || e);
      });
    if (navigator.storage && navigator.storage.persist) {
      navigator.storage.persist().catch(() => {});
    }
    document.addEventListener("visibilitychange", () => {
      if (
        document.visibilityState === "visible" &&
        registration &&
        Date.now() - lastUpdateCheck > 20 * 60 * 1000
      ) {
        lastUpdateCheck = Date.now();
        registration.update().catch(() => {});
      }
    });
  }

  window.ocCheckUpdate = async function () {
    if (!registration) {
      toast("Las actualizaciones se revisan al instalar la app desde Safari.");
      return;
    }
    try {
      await registration.update();
      if (window.ocUpdateReady) {
        if (window.ocOnUpdate) window.ocOnUpdate();
        else location.reload();
      } else if (registration.installing || registration.waiting) {
        toast("Descargando actualización de " + APP_NAME + " en segundo plano…");
      } else {
        toast(APP_NAME + " ya está en la última versión");
      }
    } catch (e) {
      toast("No se pudo buscar actualizaciones (sin conexión)");
    }
  };

  function waitForInstall(timeoutMs) {
    return new Promise((resolve, reject) => {
      if (!sw || sw.controller) return resolve();
      let done = false;
      function finish(ok, why) {
        if (done) return;
        done = true;
        clearInterval(poll);
        ok ? resolve() : reject(new Error(why));
      }
      sw.addEventListener("controllerchange", () => finish(true));
      const started = Date.now();
      const poll = setInterval(() => {
        if (sw.controller) finish(true);
        else if (installFailed) finish(false, installFailed);
        else if (Date.now() - started > timeoutMs) finish(false, "timeout");
      }, 220);
    });
  }

  window.ocFetchWasm = async function (url, expectedBytes) {
    if (sw && !sw.controller) {
      setStatus("Instalando " + APP_NAME + " para uso sin conexión…");
      try {
        await waitForInstall(8 * 60 * 1000);
      } catch (e) {
        console.warn(APP_NAME + ": sin instalación offline (" + e.message + "), se abre directo");
      }
    }
    setStatus("Iniciando " + APP_NAME + "…");
    let response;
    try {
      response = await fetch(url);
    } catch (e) {
      throw new Error("No se pudo descargar " + APP_NAME + ".\n" + SIN_RED);
    }
    if (!response.ok) {
      throw new Error("No se pudo descargar " + APP_NAME + " (" + response.status + ").\n" + SIN_RED);
    }
    if (!response.body || typeof ReadableStream !== "function") return response;
    const total = expectedBytes || Number(response.headers.get("Content-Length")) || 0;
    let loaded = 0;
    const reader = response.body.getReader();
    setProgress(0.05);
    const stream = new ReadableStream({
      async pull(controller) {
        const chunk = await reader.read();
        if (chunk.done) {
          setProgress(1);
          setStatus("Preparando interfaz…");
          controller.close();
          return;
        }
        loaded += chunk.value.byteLength;
        if (total) setProgress(loaded / total);
        controller.enqueue(chunk.value);
      },
      cancel(reason) {
        return reader.cancel(reason);
      },
    });
    return new Response(stream, { headers: { "Content-Type": "application/wasm" } });
  };

  // ---------- Compartir / Descargar en iOS ----------
  const isIOS =
    /iPad|iPhone|iPod/.test(navigator.userAgent) ||
    (navigator.platform === "MacIntel" && navigator.maxTouchPoints > 1);
  const blobs = new Map();
  const createURL = URL.createObjectURL.bind(URL);
  const revokeURL = URL.revokeObjectURL.bind(URL);
  const anchorClick = HTMLAnchorElement.prototype.click;

  const MIME_TYPES = {
    docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    doc: "application/msword",
    odt: "application/vnd.oasis.opendocument.text",
    rtf: "application/rtf",
    pptx: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    potx: "application/vnd.openxmlformats-officedocument.presentationml.template",
    ppsx: "application/vnd.openxmlformats-officedocument.presentationml.slideshow",
    deckcraft: "application/zip",
    xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    xlsm: "application/vnd.ms-excel.sheet.macroEnabled.12",
    xlsb: "application/vnd.ms-excel.sheet.binary.macroEnabled.12",
    ods: "application/vnd.oasis.opendocument.spreadsheet",
    csv: "text/csv",
    tsv: "text/tab-separated-values",
    pdf: "application/pdf",
    png: "image/png",
    jpg: "image/jpeg",
    jpeg: "image/jpeg",
    gif: "image/gif",
    txt: "text/plain",
    md: "text/markdown",
    html: "text/html",
    json: "application/json",
  };
  window.ocMimeFor = function (name) {
    const ext = String(name || "").split(".").pop().toLowerCase();
    return MIME_TYPES[ext] || "application/octet-stream";
  };

  URL.createObjectURL = function (obj) {
    const url = createURL(obj);
    if (obj instanceof Blob) {
      blobs.set(url, obj);
      if (blobs.size > 24) blobs.delete(blobs.keys().next().value);
    }
    return url;
  };
  URL.revokeObjectURL = function (url) {
    revokeURL(url);
    setTimeout(() => blobs.delete(url), 60000);
  };

  function downloadBlob(blob, name) {
    const a = document.createElement("a");
    a.href = createURL(blob);
    a.download = name;
    document.body.appendChild(a);
    anchorClick.call(a);
    a.remove();
    setTimeout(() => revokeURL(a.href), 60000);
  }

  window.ocShareFile = async function (file) {
    if (!(isIOS && navigator.share && navigator.canShare)) {
      downloadBlob(file, file.name);
      return true;
    }
    let ok = false;
    try {
      ok = navigator.canShare({ files: [file] });
    } catch (e) {
      ok = false;
    }
    if (!ok) {
      downloadBlob(file, file.name);
      return true;
    }
    try {
      await navigator.share({ files: [file] });
      return true;
    } catch (err) {
      if (err && err.name === "AbortError") return true;
      if (err && err.name === "NotAllowedError") return false;
      downloadBlob(file, file.name);
      return true;
    }
  };

  HTMLAnchorElement.prototype.click = function () {
    const blob = this.download ? blobs.get(this.href) : null;
    if (!blob) return anchorClick.call(this);
    const name = this.download;
    const file = new File([blob], name, { type: blob.type || window.ocMimeFor(name) });
    try {
      if (window.ocOnDownloadFile && window.ocOnDownloadFile(file)) return;
    } catch (e) {
      console.warn("ocOnDownloadFile", e);
    }
    window.ocShareFile(file).then((done) => {
      if (!done && window.ocOfferShare) window.ocOfferShare(file);
      else if (!done) downloadBlob(file, name);
    });
  };

  // ---------- Almacenamiento compartido (IndexedDB: officecraft-movil) ----------
  const MAX_PER_APP = 120;
  let dbPromise = null;
  function openDb() {
    if (!dbPromise) {
      dbPromise = new Promise((resolve, reject) => {
        const r = indexedDB.open("officecraft-movil", 1);
        r.onupgradeneeded = () => {
          const d = r.result;
          if (!d.objectStoreNames.contains("meta")) d.createObjectStore("meta", { keyPath: "id" });
          if (!d.objectStoreNames.contains("datos")) d.createObjectStore("datos");
          if (!d.objectStoreNames.contains("carpetas")) d.createObjectStore("carpetas", { keyPath: "id" });
        };
        r.onsuccess = () => resolve(r.result);
        r.onerror = () => reject(r.error);
      });
    }
    return dbPromise;
  }
  function reqDone(req) {
    return new Promise((resolve, reject) => {
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
  }
  async function txStore(name, mode) {
    return (await openDb()).transaction(name, mode).objectStore(name);
  }

  window.ocStore = {
    newId: () => Date.now().toString(36) + Math.random().toString(36).slice(2, 8),
    async listAll() {
      const rows = await reqDone((await txStore("meta", "readonly")).getAll());
      return rows.sort((a, b) => (b.abierto || 0) - (a.abierto || 0));
    },
    async list(appId) {
      const all = await this.listAll();
      return appId ? all.filter((m) => m.app === appId) : all;
    },
    async getMeta(id) {
      return reqDone((await txStore("meta", "readonly")).get(id));
    },
    async blob(id) {
      return reqDone((await txStore("datos", "readonly")).get(id));
    },
    async putMeta(meta) {
      return reqDone((await txStore("meta", "readwrite")).put(meta));
    },
    async save(meta, blob) {
      if (!meta.app) meta.app = CFG.id;
      const t = (await openDb()).transaction(["meta", "datos"], "readwrite");
      t.objectStore("meta").put(meta);
      t.objectStore("datos").put(blob, meta.id);
      await new Promise((resolve, reject) => {
        t.oncomplete = resolve;
        t.onerror = () => reject(t.error);
        t.onabort = () => reject(t.error);
      });
      await this.trim(meta.app);
    },
    async remove(id) {
      const t = (await openDb()).transaction(["meta", "datos"], "readwrite");
      t.objectStore("meta").delete(id);
      t.objectStore("datos").delete(id);
      await new Promise((resolve, reject) => {
        t.oncomplete = resolve;
        t.onerror = () => reject(t.error);
      });
    },
    async trim(appId) {
      const all = await this.list(appId);
      if (all.length <= MAX_PER_APP) return;
      const nonStar = all.filter((m) => !m.destacado);
      const excess = all.length - MAX_PER_APP;
      for (const old of nonStar.slice(-excess)) await this.remove(old.id);
    },
    async listFolders(appId) {
      try {
        const rows = await reqDone((await txStore("carpetas", "readonly")).getAll());
        const filtered = appId ? rows.filter((f) => !f.app || f.app === appId) : rows;
        return filtered.sort((a, b) => (a.nombre || "").localeCompare(b.nombre || "", "es"));
      } catch (e) {
        return [];
      }
    },
    async saveFolder(folder) {
      if (!folder.app) folder.app = CFG.id;
      return reqDone((await txStore("carpetas", "readwrite")).put(folder));
    },
    async removeFolder(id, appId) {
      const list = await this.list(appId);
      for (const m of list) {
        if (m.carpeta === id) {
          delete m.carpeta;
          await this.putMeta(m);
        }
      }
      return reqDone((await txStore("carpetas", "readwrite")).delete(id));
    },
  };

  // ---------- Wake Lock ----------
  let wakeLock = null;
  window.ocSyncWakeLock = async function (want) {
    const active = !!(want && document.visibilityState === "visible");
    if (!active) {
      if (wakeLock) {
        wakeLock.release().catch(() => {});
        wakeLock = null;
      }
      return;
    }
    if (!wakeLock && navigator.wakeLock && navigator.wakeLock.request) {
      try {
        wakeLock = await navigator.wakeLock.request("screen");
        wakeLock.addEventListener("release", () => {
          wakeLock = null;
        });
      } catch (e) {}
    }
  };

  // ---------- Íconos vectoriales (estilo SF Symbols / Office iOS) ----------
  const ICONS = {
    home: "M3 10.5L12 3l9 7.5V20a1 1 0 0 1-1 1h-5.5v-6.5h-5V21H4a1 1 0 0 1-1-1z",
    folder: "M3 6a2 2 0 0 1 2-2h4.2l2 2H19a2 2 0 0 1 2 2v10a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z",
    folderPlus: "M3 6a2 2 0 0 1 2-2h4.2l2 2H19a2 2 0 0 1 2 2v10a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z M12 10.5v5 M9.5 13h5",
    star: "M12 3.2l2.6 5.3 5.9.8-4.3 4.1 1 5.8-5.2-2.7-5.2 2.7 1-5.8-4.3-4.1 5.9-.8z",
    plus: "M12 5v14 M5 12h14",
    back: "M15 5l-7 7 7 7",
    chevRight: "M9 6l6 6-6 6",
    chevUp: "M6 15l6-6 6 6",
    chevDown: "M6 9l6 6 6-6",
    close: "M6 6l12 12 M18 6L6 18",
    check: "M5 12l5 5 9-10",
    search: "M10.5 17a6.5 6.5 0 1 0 0-13 6.5 6.5 0 0 0 0 13z M20 20l-4.8-4.8",
    share: "M12 3v12 M8 7l4-4 4 4 M7 10H5.5a1 1 0 0 0-1 1v9a1 1 0 0 0 1 1h13a1 1 0 0 0 1-1v-9a1 1 0 0 0-1-1H17",
    more: "M5.5 12h.01 M12 12h.01 M18.5 12h.01",
    settings: "M12 15.2a3.2 3.2 0 1 0 0-6.4 3.2 3.2 0 0 0 0 6.4z M12 2.8v2.6 M12 18.6v2.6 M4 7.4l2.3 1.3 M17.7 15.3l2.3 1.3 M4 16.6l2.3-1.3 M17.7 8.7L20 7.4",
    open: "M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z M14 3v5h5 M12 11v7 M8.5 14.5h7",
    file: "M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z M14 3v5h5 M8.5 13h7 M8.5 16.5h5",
    blank: "M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z M14 3v5h5",
    undo: "M9 14L4 9l5-5 M4 9h11a5 5 0 0 1 0 10h-3",
    redo: "M15 14l5-5-5-5 M20 9H9a5 5 0 0 0 0 10h3",
    bold: "M7 5h6a3.5 3.5 0 0 1 0 7H7z M7 12h7a3.5 3.5 0 0 1 0 7H7z",
    italic: "M10 5h8 M6 19h8 M14 5l-4 14",
    underline: "M7 4v6a5 5 0 0 0 10 0V4 M5 20h14",
    strike: "M5 12h14 M16 7.5A4 3 0 0 0 12 5c-2.5 0-4 1.2-4 2.8 0 1.4 1.2 2.2 3 2.8 M8 16.5c.6 1.6 2.1 2.5 4 2.5 2.5 0 4.2-1.3 4.2-3.1",
    highlight: "M9 15l-2 5h5l1-2 M14 4l6 6-7 7-6-6z",
    textColor: "M6 17L12 4l6 13 M8.2 12.5h7.6 M4 20h16",
    fillColor: "M12 3l7 7-7 7-7-7z M19 14c1.2 1.6 2 2.8 2 4a2 2 0 1 1-4 0c0-1.2.8-2.4 2-4z M4 20h11",
    alignLeft: "M4 6h16 M4 10h11 M4 14h16 M4 18h11",
    alignCenter: "M4 6h16 M6.5 10h11 M4 14h16 M6.5 18h11",
    alignRight: "M4 6h16 M9 10h11 M4 14h16 M9 18h11",
    alignJustify: "M4 6h16 M4 10h16 M4 14h16 M4 18h16",
    bullets: "M9 6h11 M9 12h11 M9 18h11 M5 6h.01 M5 12h.01 M5 18h.01",
    numbering: "M10 6h10 M10 12h10 M10 18h10 M4 5h2v4 M4 12h2l-2 2h2 M4 17h2v2H4",
    indent: "M4 6h16 M10 10h10 M10 14h10 M4 18h16 M4 9.5l3 2.5-3 2.5",
    outdent: "M4 6h16 M10 10h10 M10 14h10 M4 18h16 M7 9.5L4 12l3 2.5",
    mobileView: "M7 2h10a2 2 0 0 1 2 2v16a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2z M8 7h8 M8 10.5h8 M8 14h5",
    printLayout: "M6 3h12a1 1 0 0 1 1 1v16a1 1 0 0 1-1 1H6a1 1 0 0 1-1-1V4a1 1 0 0 1 1-1z M8.5 8h7 M8.5 11.5h7 M8.5 15h4.5",
    table: "M4 5h16v14H4z M4 10h16 M4 14.5h16 M9.5 5v14 M15 5v14",
    image: "M4 5h16v14H4z M4 16l5-5 4 4 2-2 5 5 M15.5 9.5h.01",
    camera: "M4 8h3l1.5-2.5h7L17 8h3a1 1 0 0 1 1 1v10a1 1 0 0 1-1 1H4a1 1 0 0 1-1-1V9a1 1 0 0 1 1-1z M12 16.5a3.5 3.5 0 1 0 0-7 3.5 3.5 0 0 0 0 7z",
    shape: "M4 5h9v9H4z M15 19a5 5 0 1 0 0-10 5 5 0 0 0 0 10z",
    textBox: "M4 5h16v14H4z M8 9h8 M12 9v7",
    chart: "M4 20h16 M7 17v-6 M12 17V7 M17 17v-9",
    comment: "M4 5a1 1 0 0 1 1-1h14a1 1 0 0 1 1 1v10a1 1 0 0 1-1 1H9l-5 4z M8 9h8 M8 12h5",
    link: "M10 14a4 4 0 0 0 5.7 0l3-3a4 4 0 0 0-5.7-5.7l-1 1 M14 10a4 4 0 0 0-5.7 0l-3 3a4 4 0 0 0 5.7 5.7l1-1",
    play: "M7 5l12 7-12 7z",
    notes: "M5 4h14v16H5z M8 9h8 M8 13h8 M8 17h5",
    slide: "M3 5h18v12H3z M12 17v4 M8 21h8",
    slidePlus: "M3 5h18v12H3z M12 17v4 M8 21h8 M12 8.5v5 M9.5 11h5",
    sorter: "M4 5h6.5v5.5H4z M13.5 5H20v5.5h-6.5z M4 13.5h6.5V19H4z M13.5 13.5H20V19h-6.5z",
    draw: "M3 18c3 0 4-3 6-6s4-6 6-6 3 2 2 4-5 5-3 8 5 1 7-1",
    eraser: "M8 20h12 M5 15l9-9 5 5-9 9H9z",
    theme: "M12 3a9 9 0 1 0 0 18 2 2 0 0 0 1.6-3.2 1.8 1.8 0 0 1 1.4-2.8H17a4 4 0 0 0 4-4c0-4.4-4-8-9-8z",
    transition: "M4 6h10v12H4z M14 9l6 3-6 3",
    sigma: "M18 5H7l6 7-6 7h11",
    fx: "M6 19l2-11a3 3 0 0 1 3-2.5h2 M5 11h6 M14 12l5 7 M19 12l-5 7",
    filter: "M4 5h16l-6 7.5V19l-4-2v-4.5z",
    sortAsc: "M4 7h10 M4 12h7 M4 17h4 M17 18V6 M14 9l3-3 3 3",
    sortDesc: "M4 7h10 M4 12h7 M4 17h4 M17 6v12 M14 15l3 3 3-3",
    sort: "M4 7h11 M4 12h8 M4 17h5 M17 7v11 M14 15l3 3 3-3",
    borders: "M4 4h16v16H4z M4 12h16 M12 4v16",
    merge: "M4 5h16v14H4z M9 12h6 M11 10l-2 2 2 2 M13 10l2 2-2 2",
    wrap: "M4 7h16 M4 12h12a3 3 0 1 1 0 6h-4 M14 16l-2 2 2 2",
    freeze: "M4 4h16v16H4z M4 9.5h16 M9.5 4v16",
    percent: "M19 5L5 19 M7.5 8a1.5 1.5 0 1 0 0-3 1.5 1.5 0 0 0 0 3z M16.5 19a1.5 1.5 0 1 0 0-3 1.5 1.5 0 0 0 0 3z",
    dollar: "M12 3v18 M16 7.5c0-1.7-1.8-3-4-3s-4 1.3-4 3 1.5 2.5 4 3 4 1.5 4 3.5-1.8 3-4 3-4-1.3-4-3",
    listView: "M4 6h16 M4 12h16 M4 18h16",
    gridView: "M4 4h6.5v6.5H4z M13.5 4H20v6.5h-6.5z M4 13.5h6.5V20H4z M13.5 13.5H20V20h-6.5z",
    rename: "M4 17h4l11-11-4-4L4 13z M13.5 5.5l4 4 M3 21h18",
    copy: "M8 8h12v12H8z M16 8V4H4v12h4",
    trash: "M4 7h16 M10 11v6 M14 11v6 M6 7l1 13h10l1-13 M9 7V4h6v3",
    pdf: "M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z M14 3v5h5 M8 15h8",
    print: "M6 9V4h12v5 M6 18H4a1 1 0 0 1-1-1v-6a1 1 0 0 1 1-1h16a1 1 0 0 1 1 1v6a1 1 0 0 1-1 1h-2 M6 14h12v7H6z",
    desktop: "M3 5h18v12H3z M8 21h8 M12 17v4",
    refresh: "M20 11a8 8 0 1 0-2.3 5.7 M20 4v7h-7",
    iphone: "M7 2h10a2 2 0 0 1 2 2v16a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2z M11 18.5h2",
    apps: "M4 4h4v4H4z M10 4h4v4h-4z M16 4h4v4h-4z M4 10h4v4H4z M10 10h4v4h-4z M16 10h4v4h-4z M4 16h4v4H4z M10 16h4v4h-4z M16 16h4v4h-4z",
    moon: "M20 14.5A8 8 0 0 1 9.5 4a8.5 8.5 0 1 0 10.5 10.5z",
    sun: "M12 16a4 4 0 1 0 0-8 4 4 0 0 0 0 8z M12 2v2 M12 20v2 M4.9 4.9l1.4 1.4 M17.7 17.7l1.4 1.4 M2 12h2 M20 12h2 M4.9 19.1l1.4-1.4 M17.7 6.3l1.4-1.4",
    clearFmt: "M6 5h12 M12 5l-3 14 M15 14l5 5 M20 14l-5 5",
    edit: "M4 20h4L19 9a2.8 2.8 0 0 0-4-4L4 16z M13.5 6.5l4 4",
    eye: "M2.5 12s3.5-6.5 9.5-6.5 9.5 6.5 9.5 6.5-3.5 6.5-9.5 6.5S2.5 12 2.5 12z M12 15a3 3 0 1 0 0-6 3 3 0 0 0 0 6z",
    info: "M12 21a9 9 0 1 0 0-18 9 9 0 0 0 0 18z M12 11v5 M12 8h.01",
    keyboard: "M3 7h18v10H3z M7 10h.01 M10 10h.01 M14 10h.01 M17 10h.01 M8 14h8",
    pageBreak: "M4 5h16v5H4z M4 14h16v5H4z M2 12h4 M10 12h4 M18 12h4",
    zoom: "M10.5 17a6.5 6.5 0 1 0 0-13 6.5 6.5 0 0 0 0 13z M20 20l-4.8-4.8 M10.5 7.5v6 M7.5 10.5h6",
  };

  const BOLD_ICONS = { more: 3.1, bullets: 2.5 };

  function icon(name, filled) {
    const w = BOLD_ICONS[name] ? ' style="stroke-width:' + BOLD_ICONS[name] + '"' : "";
    const cls = filled ? "ic ic-fill" : "ic";
    return '<svg class="' + cls + '" viewBox="0 0 24 24" aria-hidden="true"' + w + '><path d="' + (ICONS[name] || "") + '"/></svg>';
  }

  function h(tag, props, ...kids) {
    const el = document.createElement(tag);
    for (const [k, v] of Object.entries(props || {})) {
      if (v == null || v === false) continue;
      if (k === "class") el.className = v;
      else if (k === "text") el.textContent = v;
      else if (k === "icon") el.insertAdjacentHTML("beforeend", icon(v));
      else if (k === "iconFilled") el.insertAdjacentHTML("beforeend", icon(v, true));
      else if (k.startsWith("on")) el.addEventListener(k.slice(2), v);
      else el.setAttribute(k, v === true ? "" : v);
    }
    for (const kid of kids.flat()) {
      if (kid == null || kid === false) continue;
      el.append(kid.nodeType ? kid : document.createTextNode(String(kid)));
    }
    return el;
  }

  function sheet({ title, text, build, onClose, compact }) {
    const wrap = h("div", {
      class: "oc-sheet-wrap" + (compact ? " oc-sheet-compact" : ""),
      role: "dialog",
      "aria-modal": "true",
      "aria-label": title || "Opciones",
    });
    let closed = false;
    const close = () => {
      if (closed) return;
      closed = true;
      wrap.remove();
      document.removeEventListener("keydown", onKey);
      if (onClose) onClose();
    };
    const onKey = (e) => {
      if (e.key === "Escape") close();
    };
    const panel = h(
      "div",
      { class: "oc-sheet" },
      h("div", { class: "oc-grabber" }),
      title
        ? h(
            "h3",
            {},
            h("span", { text: title }),
            h("button", { class: "oc-icon-btn", type: "button", "aria-label": "Cerrar", icon: "close", onclick: close })
          )
        : null,
      text ? h("p", { text }) : null,
      build(close)
    );
    wrap.append(h("div", { class: "oc-scrim", onclick: close }), panel);
    document.addEventListener("keydown", onKey);
    document.body.append(wrap);
    return close;
  }

  function list(items) {
    return h(
      "div",
      { class: "oc-list" },
      items.map((it) =>
        h(
          "button",
          {
            class: "oc-item" + (it.danger ? " oc-danger" : ""),
            type: "button",
            icon: it.i,
            onclick: it.run,
          },
          h("span", { text: it.t }),
          it.small ? h("small", { text: it.small }) : null
        )
      )
    );
  }

  function seg(options, get, onChange) {
    const wrap = h("div", { class: "oc-seg", role: "group" });
    const draw = () =>
      wrap.replaceChildren(
        ...options.map(([value, label]) =>
          h("button", {
            type: "button",
            "aria-pressed": String(get() === value),
            text: label,
            onclick: () => {
              onChange(value);
              draw();
            },
          })
        )
      );
    draw();
    return wrap;
  }

  function swRow(label, get, onToggle) {
    const btn = h("button", {
      class: "oc-switch",
      type: "button",
      role: "switch",
      "aria-checked": String(!!get()),
      "aria-label": label,
    });
    btn.addEventListener("click", () => {
      onToggle(!get());
      btn.setAttribute("aria-checked", String(!!get()));
    });
    return h("div", { class: "oc-toggle-row" }, h("span", { text: label }), btn);
  }

  function colorGrid(colors, onPick) {
    return h(
      "div",
      { class: "oc-color-grid" },
      colors.map(([val, label, bg]) => {
        const dot = h("button", {
          class: "oc-color-dot" + (val === "none" || val === "auto" ? " oc-color-clear" : ""),
          type: "button",
          "aria-label": label,
          title: label,
          onclick: () => onPick(val),
        });
        if (bg) dot.style.background = bg;
        return dot;
      })
    );
  }

  function sizeText(bytes) {
    const b = Number(bytes) || 0;
    if (b < 1024) return b + " B";
    if (b < 1024 * 1024) return Math.round(b / 1024) + " KB";
    return (b / 1024 / 1024).toFixed(1).replace(".", ",") + " MB";
  }

  function describe(meta, extra) {
    const d = new Date(meta.modificado || meta.abierto || Date.now());
    const today = new Date();
    const when =
      d.toDateString() === today.toDateString()
        ? "Hoy, " + d.toLocaleTimeString("es", { hour: "2-digit", minute: "2-digit" })
        : d.toLocaleDateString("es", {
            day: "numeric",
            month: "short",
            year: d.getFullYear() === today.getFullYear() ? undefined : "numeric",
          });
    return when + " · " + sizeText(meta.tamano || 0) + (extra ? " · " + extra : "");
  }

  function makeSearchBox(placeholder, value, onChange, autoFocus) {
    const input = h("input", {
      type: "search",
      placeholder,
      value: value || "",
      autocomplete: "off",
      autocorrect: "off",
      autocapitalize: "off",
      spellcheck: "false",
      "aria-label": placeholder,
    });
    const clearBtn = h("button", {
      class: "oc-clear-btn",
      type: "button",
      "aria-label": "Limpiar búsqueda",
      icon: "close",
      hidden: !value,
      onclick: () => {
        input.value = "";
        clearBtn.hidden = true;
        onChange("");
        input.focus();
      },
    });
    input.addEventListener("input", () => {
      clearBtn.hidden = !input.value;
      onChange(input.value);
    });
    const box = h("label", { class: "oc-search-box" }, h("span", { icon: "search" }), input, clearBtn);
    if (autoFocus) setTimeout(() => input.focus(), 35);
    return box;
  }

  window.ocUI = {
    icon,
    h,
    sheet,
    list,
    seg,
    swRow,
    colorGrid,
    sizeText,
    describe,
    makeSearchBox,
    sleep: (ms) => new Promise((r) => setTimeout(r, ms)),
  };
})();
