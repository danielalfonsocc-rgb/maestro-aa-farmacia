"""servidor_conteo_internet.py — Conteo de Controlados accesible por INTERNET.

Levanta servidor_conteo.py (red local, puerto 8504) y ADEMÁS abre un túnel de
Cloudflare (cloudflared, "quick tunnel") para que sea accesible desde CUALQUIER
red, no solo la del hospital — sin abrir puertos en el router y sin comprar un
dominio. Decisión explícita del usuario (2026-07-25): SIN clave — cualquiera
con el link puede ver/editar el conteo de controlados.

El quick tunnel es gratis y no requiere cuenta de Cloudflare, pero la URL
(https://<palabras-al-azar>.trycloudflare.com) CAMBIA cada vez que este
proceso se reinicia (no hay forma de fijarla sin comprar un dominio propio —
ver memoria del proyecto). Por eso este script:
  1. Levanta el servidor de red (mismo código que servidor_conteo.py) en un hilo.
  2. Lanza cloudflared, lee su salida y extrae la URL nueva.
  3. La escribe en Conteo_Controlados/url_publica.txt y la imprime en pantalla,
     para no tener que rebuscarla en el log cada vez.
  4. Si cloudflared se cae (corte de red, etc.), lo reintenta solo.

Correr:   py servidor_conteo_internet.py
O:        SERVIDOR_CONTEO_INTERNET.bat
"""
import re
import shutil
import subprocess
import sys
import threading
import time
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

import servidor_conteo as sc  # reutiliza Handler / ESTADO_FILE / _ip_local, etc.

MAESTRO_DIR = Path(__file__).parent
URL_FILE = MAESTRO_DIR / "Conteo_Controlados" / "url_publica.txt"
PUERTO = 8504

CLOUDFLARED_CANDIDATOS = (
    shutil.which("cloudflared"),
    r"C:\Program Files (x86)\cloudflared\cloudflared.exe",
    r"C:\Program Files\cloudflared\cloudflared.exe",
)
RE_URL = re.compile(r"https://[a-z0-9-]+\.trycloudflare\.com")


def _cloudflared_exe():
    for c in CLOUDFLARED_CANDIDATOS:
        if c and Path(c).exists():
            return c
    return None


def _iniciar_servidor_local():
    """Corre el mismo servidor de red de servidor_conteo.py, en este proceso."""
    from http.server import ThreadingHTTPServer

    sc.ESTADO_FILE.parent.mkdir(parents=True, exist_ok=True)
    if not sc.ESTADO_FILE.exists():
        sc._guardar_estado(sc._estado_vacio())

    srv = ThreadingHTTPServer(("0.0.0.0", PUERTO), sc.Handler)
    ip = sc._ip_local()
    print(f"  Red local (hospital):  http://{ip}:{PUERTO}")
    srv.serve_forever()


def _tunel_cloudflare(exe):
    """Lanza cloudflared, imprime/guarda la URL nueva, y se reintenta solo si cae."""
    intento = 0
    while True:
        intento += 1
        if intento > 1:
            print(f"\n  [túnel] reconectando (intento {intento})...")
        proc = subprocess.Popen(
            [exe, "tunnel", "--url", f"http://localhost:{PUERTO}"],
            stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            text=True, encoding="utf-8", errors="replace",
            bufsize=1,
        )
        encontrada = False
        for linea in proc.stdout:
            m = RE_URL.search(linea)
            if m and not encontrada:
                encontrada = True
                url = m.group(0)
                URL_FILE.parent.mkdir(parents=True, exist_ok=True)
                URL_FILE.write_text(url + "\n", encoding="utf-8")
                print("\n" + "=" * 62)
                print("  INTERNET (desde cualquier red):")
                print(f"  {url}")
                print(f"  (también guardada en {URL_FILE})")
                print("=" * 62 + "\n")
        proc.wait()
        print(f"  [túnel] cloudflared terminó (código {proc.returncode}); reintentando en 5s...")
        time.sleep(5)


def main():
    if not sc.HTML_FILE.exists():
        print(f"[ERROR] No se encuentra {sc.HTML_FILE.name} en {MAESTRO_DIR}")
        sys.exit(1)

    exe = _cloudflared_exe()
    if not exe:
        print("[ERROR] No encontré cloudflared.exe. Instálalo con:")
        print('  winget install --id Cloudflare.cloudflared -e')
        sys.exit(1)

    print("=" * 62)
    print("  CONTEO DE CONTROLADOS — SERVIDOR + TÚNEL A INTERNET")
    print("=" * 62)

    hilo_srv = threading.Thread(target=_iniciar_servidor_local, daemon=True)
    hilo_srv.start()
    time.sleep(1)  # dar tiempo a que el servidor local levante antes del túnel

    try:
        _tunel_cloudflare(exe)
    except KeyboardInterrupt:
        print("\nDetenido.")


if __name__ == "__main__":
    main()
