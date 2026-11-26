(function () {
  "use strict";

  if ("serviceWorker" in navigator && window.isSecureContext) {
    navigator.serviceWorker.register("./sw.js", { scope: "./" }).catch(() => {});
  }

  let filter = "all";
  let query = "";
  let allFiles = [];

  const searchInp = document.getElementById("pt-search");
  const chipsWrap = document.getElementById("pt-chips");
  const recentsEl = document.getElementById("pt-recents");
  const installBtn = document.getElementById("pt-install-btn");

  installBtn.addEventListener("click", () => {
    const g = document.getElementById("pt-guide");
    if (g) g.scrollIntoView({ behavior: "smooth" });
  });

  searchInp.addEventListener("input", () => {
    query = searchInp.value.trim().toLowerCase();
    render();
  });

  chipsWrap.addEventListener("click", (e) => {
    const btn = e.target.closest("button[data-f]");
    if (!btn) return;
    filter = btn.getAttribute("data-f");
    chipsWrap.querySelectorAll("button").forEach((b) => {
      b.setAttribute("aria-pressed", String(b === btn));
    });
    render();
  });

  function readRecents() {
    return new Promise((resolve) => {
      try {
        const req = indexedDB.open("officecraft-movil", 1);
        req.onupgradeneeded = () => {
          const d = req.result;
          if (!d.objectStoreNames.contains("meta")) d.createObjectStore("meta", { keyPath: "id" });
          if (!d.objectStoreNames.contains("datos")) d.createObjectStore("datos");
          if (!d.objectStoreNames.contains("carpetas")) d.createObjectStore("carpetas", { keyPath: "id" });
        };
        req.onsuccess = () => {
          const db = req.result;
          if (!db.objectStoreNames.contains("meta")) return resolve([]);
          const tx = db.transaction("meta", "readonly");
          const g = tx.objectStore("meta").getAll();
          g.onsuccess = () => resolve((g.result || []).sort((a, b) => (b.abierto || 0) - (a.abierto || 0)));
          g.onerror = () => resolve([]);
        };
        req.onerror = () => resolve([]);
      } catch (e) {
        resolve([]);
      }
    });
  }

  function sizeText(b) {
    b = Number(b) || 0;
    if (b < 1024) return b + " B";
    if (b < 1024 * 1024) return Math.round(b / 1024) + " KB";
    return (b / 1024 / 1024).toFixed(1).replace(".", ",") + " MB";
  }

  function fmtDate(ts) {
    const d = new Date(ts || Date.now());
    const today = new Date();
    return d.toDateString() === today.toDateString()
      ? "Hoy, " + d.toLocaleTimeString("es", { hour: "2-digit", minute: "2-digit" })
      : d.toLocaleDateString("es", { day: "numeric", month: "short" });
  }

  function render() {
    let items = allFiles;
    if (filter !== "all") items = items.filter((m) => m.app === filter);
    if (query) items = items.filter((m) => (m.nombre || "").toLowerCase().includes(query));

    if (!items.length) {
      recentsEl.innerHTML =
        '<div class="pt-empty">' +
        (query
          ? "No se encontraron archivos con ese nombre."
          : "Aquí aparecerán tus documentos recientes de WordCraft, DeckCraft y GridCraft guardados en este iPhone.") +
        "</div>";
      return;
    }

    const ul = document.createElement("ul");
    ul.className = "pt-list";
    for (const m of items.slice(0, 30)) {
      const app = m.app === "deck" ? "deck" : m.app === "grid" ? "grid" : "word";
      const tag = app === "deck" ? "PPT" : app === "grid" ? "XLS" : "DOC";
      const a = document.createElement("a");
      a.className = "pt-row";
      a.href = "./" + app + "/?open=" + encodeURIComponent(m.id);

      const badge = document.createElement("span");
      badge.className = "pt-badge " + app;
      badge.textContent = tag;

      const txt = document.createElement("span");
      txt.className = "pt-row-txt";
      const b = document.createElement("b");
      b.textContent = m.nombre || "Sin título";
      const sm = document.createElement("small");
      sm.textContent = fmtDate(m.modificado || m.abierto) + " · " + sizeText(m.tamano);
      txt.append(b, sm);

      a.append(badge, txt);
      const li = document.createElement("li");
      li.append(a);
      ul.append(li);
    }
    recentsEl.replaceChildren(ul);
  }

  readRecents().then((rows) => {
    allFiles = rows;
    render();
  });
})();
