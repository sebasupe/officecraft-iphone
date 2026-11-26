"""Baja los últimos builds web oficiales de WordCraft, DeckCraft y GridCraft
desde sus releases oficiales en GitHub (github.com/storytold/*).

Verifica el tamaño y la huella SHA-256 que publica GitHub para cada archivo.

Uso:
    python scripts/descargar_web.py [carpeta destino, por defecto web/] [wordcraft deckcraft gridcraft]
"""
import hashlib
import json
import os
import shutil
import sys
import tempfile
import urllib.request
import zipfile
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
APPS = {
    "wordcraft": {
        "repo": "storytold/wordcraft",
        "prefix": "wordcraft-web-",
        "env": "WORDCRAFT_VERSION",
    },
    "deckcraft": {
        "repo": "storytold/deckcraft",
        "prefix": "deckcraft-web-",
        "env": "DECKCRAFT_VERSION",
    },
    "gridcraft": {
        "repo": "storytold/gridcraft",
        "prefix": "gridcraft-web-",
        "env": "GRIDCRAFT_VERSION",
    },
}


def pedir(url: str, accept: str = "application/vnd.github+json"):
    req = urllib.request.Request(
        url,
        headers={"Accept": accept, "User-Agent": "officecraft-iphone-build"},
    )
    token = os.environ.get("GH_TOKEN")
    if token and url.startswith("https://api.github.com/"):
        req.add_header("Authorization", f"Bearer {token}")
    return urllib.request.urlopen(req, timeout=180)


def descargar_uno(nombre: str, destino_raiz: Path) -> str:
    cfg = APPS[nombre]
    api = f"https://api.github.com/repos/{cfg['repo']}/releases/latest"
    release = json.load(pedir(api))
    tag = release["tag_name"]
    version = tag.lstrip("vV")
    asset = next(
        (
            a
            for a in release["assets"]
            if a["name"].startswith(cfg["prefix"]) and a["name"].endswith(".zip")
        ),
        None,
    )
    if asset is None:
        sys.exit(f"error: no encontré {cfg['prefix']}*.zip en {cfg['repo']} {tag}")

    print(f"{nombre} {version}: {asset['name']} ({asset['size'] / 1e6:.2f} MB)")

    destino = destino_raiz / nombre
    ver_file = destino / "VERSION"
    digest_file = destino / ".digest"
    digest = (asset.get("digest") or "").lower()
    if (
        destino.is_dir()
        and (destino / "index.html").is_file()
        and ver_file.is_file()
        and ver_file.read_text(encoding="utf-8").strip() == version
        and (not digest or (digest_file.is_file() and digest_file.read_text(encoding="utf-8").strip() == digest))
    ):
        print(f"  ya descargado y verificado en {destino}")
        return version

    with tempfile.TemporaryDirectory() as tmp:
        zip_path = Path(tmp) / asset["name"]
        with pedir(asset["browser_download_url"], accept="application/octet-stream") as r, open(zip_path, "wb") as f:
            shutil.copyfileobj(r, f, 1 << 20)
        if zip_path.stat().st_size != asset["size"]:
            sys.exit(f"error: la descarga de {nombre} no tiene el tamaño esperado")
        if digest.startswith("sha256:"):
            h = hashlib.sha256(zip_path.read_bytes()).hexdigest()
            if h != digest.split(":", 1)[1]:
                sys.exit(f"error: la huella SHA-256 de {nombre} no coincide")
            print("  huella SHA-256 verificada")

        if destino.exists():
            shutil.rmtree(destino)
        extraido = Path(tmp) / "x"
        with zipfile.ZipFile(zip_path) as z:
            for item in z.namelist():
                if item.startswith("/") or ".." in Path(item).parts:
                    sys.exit(f"error: ruta no permitida en el zip de {nombre}: {item}")
            z.extractall(extraido)
        carpeta = next(p.parent for p in extraido.rglob("index.html"))
        destino.parent.mkdir(parents=True, exist_ok=True)
        shutil.move(str(carpeta), str(destino))

    for sobra in ("_headers", ".htaccess", "HOSTING.md"):
        (destino / sobra).unlink(missing_ok=True)
    ver_file.write_text(version, encoding="utf-8")
    if digest:
        digest_file.write_text(digest, encoding="utf-8")

    env = os.environ.get("GITHUB_ENV")
    if env:
        with open(env, "a", encoding="utf-8") as f:
            f.write(f"{cfg['env']}={version}\n")
    print(f"  listo: {destino}")
    return version


def main() -> None:
    destino_raiz = Path(sys.argv[1]).resolve() if len(sys.argv) > 1 else RAIZ / "web"
    objetivos = sys.argv[2:] if len(sys.argv) > 2 else list(APPS.keys())
    for obj in objetivos:
        if obj not in APPS:
            sys.exit(f"error: app desconocida '{obj}'. Usa: {', '.join(APPS.keys())}")
    destino_raiz.mkdir(parents=True, exist_ok=True)
    versiones = {}
    for obj in objetivos:
        versiones[obj] = descargar_uno(obj, destino_raiz)
    (destino_raiz / "versions.json").write_text(
        json.dumps(versiones, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print("Descargas completadas:", versiones)


if __name__ == "__main__":
    main()
