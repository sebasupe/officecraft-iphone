"""Servidor HTTP local para probar OfficeCraft (WordCraft, DeckCraft, GridCraft)
en el navegador o desde el iPhone en la misma red Wi-Fi.

Sirve .wasm con tipo MIME application/wasm y sin caché agresiva en desarrollo.

Uso:
    python scripts/servidor.py [carpeta, por defecto _site] [puerto, por defecto 8765]
"""
import http.server
import mimetypes
import socketserver
import sys
from pathlib import Path

mimetypes.add_type("application/wasm", ".wasm")
mimetypes.add_type("application/manifest+json", ".webmanifest")
mimetypes.add_type("text/javascript", ".js")


class Handler(http.server.SimpleHTTPRequestHandler):
    def end_headers(self):
        self.send_header("Cache-Control", "no-store")
        super().end_headers()

    def log_message(self, fmt, *args):
        if len(args) >= 2 and str(args[1]).startswith(("4", "5")):
            super().log_message(fmt, *args)


class ThreadedServer(socketserver.ThreadingMixIn, http.server.HTTPServer):
    daemon_threads = True
    allow_reuse_address = True


def main():
    carpeta = Path(sys.argv[1]).resolve() if len(sys.argv) > 1 else Path(__file__).resolve().parent.parent / "_site"
    puerto = int(sys.argv[2]) if len(sys.argv) > 2 else 8765
    if not carpeta.is_dir():
        sys.exit(f"error: no existe {carpeta}. Ejecuta primero scripts/preparar.py")
    handler = lambda *a, **kw: Handler(*a, directory=str(carpeta), **kw)
    with ThreadedServer(("127.0.0.1", puerto), handler) as httpd:
        print(f"Sirviendo {carpeta} en http://127.0.0.1:{puerto}/")
        print(f"  Portal:    http://127.0.0.1:{puerto}/")
        print(f"  WordCraft: http://127.0.0.1:{puerto}/word/")
        print(f"  DeckCraft: http://127.0.0.1:{puerto}/deck/")
        print(f"  GridCraft: http://127.0.0.1:{puerto}/grid/")
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            pass


if __name__ == "__main__":
    main()
