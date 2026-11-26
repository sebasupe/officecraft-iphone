"""Arma el sitio completo de OfficeCraft para iPhone (Portal + WordCraft + DeckCraft + GridCraft)
a partir de los builds web oficiales sin modificar los binarios WASM.

Cada aplicación (word/, deck/, grid/) queda como una PWA independiente con su propio
manifest.webmanifest, ícono oficial de iPhone, service worker de alcance propio (./)
y caché offline que reutiliza archivos WASM con huella cuando solo cambia la interfaz.

Uso:
    python scripts/preparar.py <carpeta web/> <carpeta assets/> <carpeta salida _site/>
"""
import hashlib
import json
import re
import shutil
import sys
from pathlib import Path

from PIL import Image, ImageDraw

RAIZ = Path(__file__).resolve().parent.parent
PLANTILLA = RAIZ / "plantilla"

# Archivos de texto/licencias que no hacen falta para arrancar sin conexión.
NO_PRECARGAR = {
    "README.md",
    "ATTRIBUTION.md",
    "NOTICE",
    "LICENSE-APACHE",
    "LICENSE-MIT",
    "VERSION",
    ".digest",
    ".nojekyll",
}
SOBRAN = ("_headers", ".htaccess", "HOSTING.md", ".digest")

APPS = {
    "word": {
        "web_dir": "wordcraft",
        "name": "WordCraft",
        "short_name": "WordCraft",
        "description": "Procesador de textos Word (.docx) de código abierto para iPhone.",
        "icon_file": "wordcraft-1024.png",
        "glue_re": r"wordcraft-web-[0-9a-f]+\.js",
        "wasm_re": r"wordcraft-web-([0-9a-f]+)_bg\.wasm",
        "bg": "#f2f4f8",
        "theme": "#185abd",
        "cache_prefix": "wordcraft-",
    },
    "deck": {
        "web_dir": "deckcraft",
        "name": "DeckCraft",
        "short_name": "DeckCraft",
        "description": "Presentaciones PowerPoint (.pptx) de código abierto para iPhone.",
        "icon_file": "deckcraft-1024.png",
        "glue_re": r"deckcraft-web-[0-9a-f]+\.js",
        "wasm_re": r"deckcraft-web-([0-9a-f]+)_bg\.wasm",
        "bg": "#f5f4f6",
        "theme": "#c43e1c",
        "cache_prefix": "deckcraft-",
    },
    "grid": {
        "web_dir": "gridcraft",
        "name": "GridCraft",
        "short_name": "GridCraft",
        "description": "Hojas de cálculo Excel (.xlsx) de código abierto para iPhone.",
        "icon_file": "gridcraft-1024.png",
        "glue_re": r"gridcraft-web-[0-9a-f]+\.js",
        "wasm_re": r"gridcraft-web-([0-9a-f]+)_bg\.wasm",
        "bg": "#f2f5f3",
        "theme": "#107c41",
        "cache_prefix": "gridcraft-",
    },
}


def uno(carpeta: Path, patron: str) -> Path:
    hallados = [p for p in carpeta.iterdir() if re.fullmatch(patron, p.name)]
    if len(hallados) != 1:
        sys.exit(f"error: esperaba un archivo {patron} en {carpeta}, hallé {len(hallados)}")
    return hallados[0]


def generar_iconos(png: Path, destino: Path) -> None:
    destino.mkdir(parents=True, exist_ok=True)
    im = Image.open(png).convert("RGBA")
    # Recorta márgenes transparentes y rellena las esquinas con el color base del ícono
    # porque iOS pinta de negro la transparencia y luego aplica su propia máscara squircle.
    bbox = im.getchannel("A").point(lambda a: 255 if a > 200 else 0).getbbox()
    if bbox:
        im = im.crop(bbox)
    w, h = im.size
    # Muestrea un pixel opaco cerca del borde superior-central (o izquierdo-central)
    color = im.getpixel((w // 2, max(1, int(h * 0.04))))
    if color[3] < 200:
        color = im.getpixel((max(1, int(w * 0.08)), h // 2))
    fondo = Image.new("RGBA", im.size, (color[0], color[1], color[2], 255))
    lleno = Image.alpha_composite(fondo, im).convert("RGB")
    for nombre, lado in (
        ("apple-touch-icon.png", 180),
        ("icon-192.png", 192),
        ("icon-512.png", 512),
    ):
        lleno.resize((lado, lado), Image.LANCZOS).save(destino / nombre, optimize=True)
    lado = 512
    m = Image.new("RGB", (lado, lado), color[:3])
    interior = lleno.resize((int(lado * 0.82), int(lado * 0.82)), Image.LANCZOS)
    m.paste(interior, ((lado - interior.width) // 2, (lado - interior.height) // 2))
    m.save(destino / "icon-maskable-512.png", optimize=True)


def generar_icono_portal(assets: Path, destino: Path) -> None:
    """Crea un ícono limpio de suite OfficeCraft (mosaico con los 3 íconos oficiales)."""
    destino.mkdir(parents=True, exist_ok=True)
    size = 512
    canvas = Image.new("RGB", (size, size), (24, 32, 56))
    draw = ImageDraw.Draw(canvas)
    # Fondo degradado azul profundo elegante
    for y in range(size):
        t = y / (size - 1)
        r = int(20 + 24 * t)
        g = int(42 + 28 * t)
        b = int(98 + 46 * t)
        draw.line([(0, y), (size, y)], fill=(r, g, b))

    def tile_from(path: Path, box_size: int) -> Image.Image:
        im = Image.open(path).convert("RGBA")
        bbox = im.getchannel("A").point(lambda a: 255 if a > 200 else 0).getbbox()
        if bbox:
            im = im.crop(bbox)
        return im.resize((box_size, box_size), Image.LANCZOS)

    # Disposición limpia en 3 tarjetas redondeadas (Word arriba-izq, Deck arriba-der, Grid abajo-centro)
    t_size = 186
    w_im = tile_from(assets / "wordcraft-1024.png", t_size)
    d_im = tile_from(assets / "deckcraft-1024.png", t_size)
    g_im = tile_from(assets / "gridcraft-1024.png", t_size)

    canvas_rgba = canvas.convert("RGBA")
    canvas_rgba.alpha_composite(w_im, (52, 56))
    canvas_rgba.alpha_composite(d_im, (size - 52 - t_size, 56))
    canvas_rgba.alpha_composite(g_im, ((size - t_size) // 2, size - 56 - t_size))
    final = canvas_rgba.convert("RGB")

    for nombre, lado in (
        ("apple-touch-icon.png", 180),
        ("icon-192.png", 192),
        ("icon-512.png", 512),
        ("icon-maskable-512.png", 512),
    ):
        final.resize((lado, lado), Image.LANCZOS).save(destino / nombre, optimize=True)


def preparar_app(clave: str, cfg: dict, web_root: Path, assets: Path, salida_root: Path) -> dict:
    release = web_root / cfg["web_dir"]
    salida = salida_root / clave
    tpl_app = PLANTILLA / clave
    tpl_comun = PLANTILLA / "comun"

    if not (release / "index.html").is_file():
        sys.exit(f"error: {release} no tiene index.html")

    original = (release / "index.html").read_text(encoding="utf-8")
    glue = uno(release, cfg["glue_re"])
    wasm = uno(release, cfg["wasm_re"])
    build = re.search(cfg["wasm_re"], wasm.name).group(1)
    ver_file = release / "VERSION"
    app_ver = ver_file.read_text(encoding="utf-8").strip() if ver_file.is_file() else build

    # Conservar modulepreload del glue y snippets; omitir preload del .wasm para no bajarlo dos veces
    cabecera = [
        t
        for t in re.findall(r"<link\b[^>]*>", original)
        if not ('rel="preload"' in t and ".wasm" in t) and 'rel="icon"' not in t
    ]

    if salida.exists():
        shutil.rmtree(salida)
    shutil.copytree(release, salida)
    for sobra in SOBRAN:
        (salida / sobra).unlink(missing_ok=True)
    # Quita archivos OFL/licencias gigantes solo del precache (quedan accesibles en el sitio)
    generar_iconos(assets / cfg["icon_file"], salida / "icons")

    # Copia recursos comunes DENTRO del alcance de la app (./) para 100% offline independiente en iOS
    for f in sorted(tpl_comun.iterdir()):
        if f.is_file() and f.name != "sw.js":
            shutil.copy2(f, salida / f.name)

    # Copia los archivos específicos de la app
    for f in sorted(tpl_app.iterdir()):
        if f.is_file() and f.name != "index.html":
            shutil.copy2(f, salida / f.name)

    html = (tpl_app / "index.html").read_text(encoding="utf-8")
    html = (
        html.replace("__HEAD_LINKS__", "\n    ".join(cabecera))
        .replace("__GLUE__", glue.name)
        .replace("__WASM__", wasm.name)
        .replace("__WASM_BYTES__", str(wasm.stat().st_size))
        .replace("__APP_VERSION__", app_ver)
    )
    (salida / "index.html").write_text(html, encoding="utf-8")

    manifest = {
        "name": cfg["name"],
        "short_name": cfg["short_name"],
        "description": cfg["description"],
        "id": "./",
        "start_url": "./",
        "scope": "./",
        "display": "standalone",
        "orientation": "any",
        "lang": "es",
        "background_color": cfg["bg"],
        "theme_color": cfg["theme"],
        "icons": [
            {"src": "icons/icon-192.png", "sizes": "192x192", "type": "image/png"},
            {"src": "icons/icon-512.png", "sizes": "512x512", "type": "image/png"},
            {
                "src": "icons/icon-maskable-512.png",
                "sizes": "512x512",
                "type": "image/png",
                "purpose": "maskable",
            },
        ],
    }
    (salida / "manifest.webmanifest").write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    precache = ["./", "./index.html"]
    for f in sorted(salida.rglob("*")):
        rel = f.relative_to(salida).as_posix()
        if (
            f.is_dir()
            or f.name in NO_PRECARGAR
            or f.name.startswith("OFL-")
            or rel.startswith(("licenses/", "models/"))
            or rel in ("index.html", "sw.js")
        ):
            continue
        precache.append("./" + rel)

    huella = hashlib.sha256()
    for rel in precache[2:]:
        huella.update(rel.encode("utf-8"))
        huella.update((salida / rel[2:]).read_bytes())
    huella.update(html.encode("utf-8"))
    sw_tpl = (tpl_comun / "sw.js").read_text(encoding="utf-8")
    huella.update(sw_tpl.encode("utf-8"))

    sw = (
        sw_tpl.replace("__APP_NAME__", cfg["name"])
        .replace("__CACHE_PREFIX__", cfg["cache_prefix"])
        .replace("__VERSION__", build + "-" + huella.hexdigest()[:10])
        .replace("__ENGINE_VER__", app_ver)
        .replace("__WASM__", wasm.name)
        .replace("__WASM_BYTES__", str(wasm.stat().st_size))
        .replace("__PRECACHE__", json.dumps(precache, indent=2))
    )
    (salida / "sw.js").write_text(sw, encoding="utf-8")

    tam = sum(f.stat().st_size for f in salida.rglob("*") if f.is_file())
    print(f"  [{clave}] {cfg['name']} {app_ver} (build {build}): {len(precache)} archivos offline, {tam / 1e6:.1f} MB")
    return {
        "key": clave,
        "name": cfg["name"],
        "version": app_ver,
        "build": build,
        "wasmBytes": wasm.stat().st_size,
        "files": len(precache),
    }


def preparar_portal(assets: Path, salida: Path, resumen: dict) -> None:
    tpl_portal = PLANTILLA / "portal"
    generar_icono_portal(assets, salida / "icons")
    for f in sorted(tpl_portal.iterdir()):
        if f.is_file() and f.name not in ("index.html", "sw.js"):
            shutil.copy2(f, salida / f.name)

    html = (tpl_portal / "index.html").read_text(encoding="utf-8")
    html = (
        html.replace("__WORD_VER__", resumen["word"]["version"])
        .replace("__DECK_VER__", resumen["deck"]["version"])
        .replace("__GRID_VER__", resumen["grid"]["version"])
        .replace("__VERSIONS_JSON__", json.dumps(resumen, ensure_ascii=False))
    )
    (salida / "index.html").write_text(html, encoding="utf-8")

    manifest = {
        "name": "OfficeCraft",
        "short_name": "OfficeCraft",
        "description": "WordCraft, DeckCraft y GridCraft para iPhone: documentos, presentaciones y hojas de cálculo sin conexión.",
        "id": "./",
        "start_url": "./",
        "scope": "./",
        "display": "standalone",
        "orientation": "any",
        "lang": "es",
        "background_color": "#f4f5f8",
        "theme_color": "#185abd",
        "icons": [
            {"src": "icons/icon-192.png", "sizes": "192x192", "type": "image/png"},
            {"src": "icons/icon-512.png", "sizes": "512x512", "type": "image/png"},
            {
                "src": "icons/icon-maskable-512.png",
                "sizes": "512x512",
                "type": "image/png",
                "purpose": "maskable",
            },
        ],
    }
    (salida / "manifest.webmanifest").write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    portal_files = [
        "./",
        "./index.html",
        "./portal.css",
        "./portal.js",
        "./manifest.webmanifest",
        "./icons/icon-192.png",
        "./icons/icon-512.png",
        "./icons/apple-touch-icon.png",
        "./word/icons/icon-192.png",
        "./deck/icons/icon-192.png",
        "./grid/icons/icon-192.png",
    ]
    h = hashlib.sha256()
    for rel in portal_files[2:]:
        p = salida / rel[2:]
        if p.is_file():
            h.update(rel.encode("utf-8"))
            h.update(p.read_bytes())
    h.update(html.encode("utf-8"))
    sw_tpl = (tpl_portal / "sw.js").read_text(encoding="utf-8")
    h.update(sw_tpl.encode("utf-8"))
    sw = sw_tpl.replace("__VERSION__", h.hexdigest()[:12]).replace(
        "__PRECACHE__", json.dumps(portal_files, indent=2)
    )
    (salida / "sw.js").write_text(sw, encoding="utf-8")


def main() -> None:
    if len(sys.argv) < 4:
        sys.exit(__doc__)
    web_root = Path(sys.argv[1]).resolve()
    assets = Path(sys.argv[2]).resolve()
    salida = Path(sys.argv[3]).resolve()

    if salida.exists():
        shutil.rmtree(salida)
    salida.mkdir(parents=True, exist_ok=True)
    (salida / ".nojekyll").write_text("", encoding="utf-8")

    print(f"Armando OfficeCraft en {salida}…")
    resumen = {}
    for clave, cfg in APPS.items():
        resumen[clave] = preparar_app(clave, cfg, web_root, assets, salida)

    preparar_portal(assets, salida, resumen)
    (salida / "versions.json").write_text(
        json.dumps(resumen, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    tam = sum(f.stat().st_size for f in salida.rglob("*") if f.is_file())
    print(f"Listo: {salida} ({tam / 1e6:.1f} MB totales)")


if __name__ == "__main__":
    main()
