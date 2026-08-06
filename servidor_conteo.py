"""servidor_conteo.py — Servidor de RED para el Conteo de Controlados (modo online).

A diferencia de conteo_controlados_app.py (Streamlit, historial en localStorage,
pensado para 1 computador), este servidor guarda el conteo EN DISCO
(Conteo_Controlados/estado_compartido.json) para que VARIAS COMPUTADORAS puedan
llenar el MISMO conteo al mismo tiempo — típicamente una persona por farmacia
(AT Cerrada / AT Abierta / Urgencia), cada una desde su propio equipo.

No agrega dependencias nuevas: usa únicamente http.server (librería estándar).
Sirve el mismo conteo_controlados.html de siempre, pero le inyecta
window.CONTEO_ONLINE = true para que active su capa de sincronización (fetch
a /api/estado, /api/conteo, /api/historial/...). Sin esa bandera (Streamlit o
el HTML abierto suelto por doble clic) todo sigue funcionando solo con
localStorage, como antes — este servidor es un modo adicional, no un reemplazo.

Concurrencia: cada request corre en su propio hilo (ThreadingHTTPServer); un
lock protege la lectura/escritura de estado_compartido.json para que dos
computadoras guardando al mismo tiempo no corrompan el archivo. El historial
(guardar/eliminar) usa endpoints atómicos (agregar UNA entrada, eliminar por
su "guardadoEn") en vez de reemplazar el arreglo completo, para no perder
entradas guardadas por otra computadora entre medio.

Correr:   py servidor_conteo.py                 (puerto 8504, red 0.0.0.0)
          py servidor_conteo.py --port 8504
O:        SERVIDOR_CONTEO_ONLINE.bat
"""
import argparse
import json
import socket
import sys
import threading
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

MAESTRO_DIR = Path(__file__).parent
HTML_FILE   = MAESTRO_DIR / "conteo_controlados.html"
CONFIG_FILE = MAESTRO_DIR / "controlados_config.json"
STOCK_FILE  = MAESTRO_DIR / "Conteo_Controlados" / "stock_sistema.json"
ACTA_VENC_FILE = MAESTRO_DIR / "Conteo_Controlados" / "acta_vencimiento.json"
ESTADO_FILE = MAESTRO_DIR / "Conteo_Controlados" / "estado_compartido.json"

_lock = threading.Lock()  # protege lectura/escritura de ESTADO_FILE entre hilos


def _cargar_json(path, default=None):
    try:
        return json.loads(Path(path).read_text(encoding="utf-8"))
    except Exception:
        return default


def _estado_vacio():
    return {"conteos": {}, "historial": {}, "actualizado": {}}


def _leer_estado():
    return _cargar_json(ESTADO_FILE, None) or _estado_vacio()


def _guardar_estado(estado):
    ESTADO_FILE.parent.mkdir(parents=True, exist_ok=True)
    ESTADO_FILE.write_text(json.dumps(estado, ensure_ascii=False, indent=2), encoding="utf-8")


def _ip_local():
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(("8.8.8.8", 80))  # no envía nada real; solo resuelve la interfaz de salida
        return s.getsockname()[0]
    except Exception:
        return "127.0.0.1"
    finally:
        s.close()


class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        print(f"  {self.address_string()} - {fmt % args}")

    def _json(self, status, payload):
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _leer_body(self):
        largo = int(self.headers.get("Content-Length", 0) or 0)
        if not largo:
            return {}
        return json.loads(self.rfile.read(largo).decode("utf-8"))

    # ── GET ──────────────────────────────────────────────────────────────────
    def do_GET(self):
        ruta = urlparse(self.path).path
        if ruta in ("/", "/index.html"):
            self._servir_html()
        elif ruta == "/api/estado":
            with _lock:
                estado = _leer_estado()
            self._json(200, estado)
        else:
            self._json(404, {"error": "no encontrado"})

    def _servir_html(self):
        if not HTML_FILE.exists():
            self._json(500, {"error": "conteo_controlados.html no existe"})
            return
        config = _cargar_json(CONFIG_FILE, {})
        stock = _cargar_json(STOCK_FILE, None)
        acta_venc = _cargar_json(ACTA_VENC_FILE, None)
        html = HTML_FILE.read_text(encoding="utf-8")
        inyeccion = (
            "<script>\n"
            f"window.CONTROLADOS_CONFIG = {json.dumps(config, ensure_ascii=False)};\n"
            f"window.STOCK_SISTEMA_INYECTADO = "
            f"{json.dumps(stock, ensure_ascii=False) if stock else 'null'};\n"
            f"window.ACTA_VENCIMIENTO_INYECTADO = "
            f"{json.dumps(acta_venc, ensure_ascii=False) if acta_venc else 'null'};\n"
            "window.CONTEO_ONLINE = true;\n"
            "</script>\n"
        )
        html = html.replace("<body>", "<body>\n" + inyeccion, 1)
        body = html.encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    # ── POST ─────────────────────────────────────────────────────────────────
    def do_POST(self):
        ruta = urlparse(self.path).path
        try:
            data = self._leer_body()
        except Exception:
            self._json(400, {"error": "JSON inválido"})
            return

        if ruta == "/api/conteo":
            self._api_conteo(data)
        elif ruta == "/api/historial/agregar":
            self._api_historial_agregar(data)
        elif ruta == "/api/historial/eliminar":
            self._api_historial_eliminar(data)
        else:
            self._json(404, {"error": "no encontrado"})

    def _api_conteo(self, data):
        fid, conteo = data.get("farmaciaId"), data.get("conteo")
        if not fid or conteo is None:
            self._json(400, {"error": "faltan farmaciaId/conteo"})
            return
        with _lock:
            estado = _leer_estado()
            estado["conteos"][fid] = conteo
            ahora = datetime.now().isoformat(timespec="seconds")
            estado["actualizado"][fid] = ahora
            _guardar_estado(estado)
        self._json(200, {"ok": True, "actualizado": ahora})

    def _api_historial_agregar(self, data):
        fid, entry = data.get("farmaciaId"), data.get("entry")
        if not fid or entry is None:
            self._json(400, {"error": "faltan farmaciaId/entry"})
            return
        with _lock:
            estado = _leer_estado()
            estado["historial"].setdefault(fid, [])
            estado["historial"][fid].insert(0, entry)
            _guardar_estado(estado)
            historial = estado["historial"]
        self._json(200, {"ok": True, "historial": historial})

    def _api_historial_eliminar(self, data):
        fid, guardado_en = data.get("farmaciaId"), data.get("guardadoEn")
        if not fid or not guardado_en:
            self._json(400, {"error": "faltan farmaciaId/guardadoEn"})
            return
        with _lock:
            estado = _leer_estado()
            lista = estado["historial"].get(fid, [])
            estado["historial"][fid] = [h for h in lista if h.get("guardadoEn") != guardado_en]
            _guardar_estado(estado)
            historial = estado["historial"]
        self._json(200, {"ok": True, "historial": historial})


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--port", type=int, default=8504)
    ap.add_argument("--host", default="0.0.0.0")
    args = ap.parse_args()

    if not HTML_FILE.exists():
        print(f"[ERROR] No se encuentra {HTML_FILE.name} en {MAESTRO_DIR}")
        sys.exit(1)

    ESTADO_FILE.parent.mkdir(parents=True, exist_ok=True)
    if not ESTADO_FILE.exists():
        _guardar_estado(_estado_vacio())

    srv = ThreadingHTTPServer((args.host, args.port), Handler)
    ip = _ip_local()
    print("=" * 62)
    print("  CONTEO DE CONTROLADOS — SERVIDOR ONLINE (varias computadoras)")
    print("=" * 62)
    print(f"  Este computador:     http://localhost:{args.port}")
    print(f"  Otros computadores:  http://{ip}:{args.port}")
    print()
    print("  Todos los que abran esa dirección llenan el MISMO conteo,")
    print("  en vivo (poll cada 4 s). Deja esta ventana abierta mientras")
    print("  se use. Ctrl+C para detener.")
    print("=" * 62)
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        print("\nServidor detenido.")


if __name__ == "__main__":
    main()
