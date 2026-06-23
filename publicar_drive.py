#!/usr/bin/env python3
"""
publicar_drive.py — Sube las salidas de Farmacia AA a Google Drive.

Sube los mismos archivos que publicar_escritorio.py EXCEPTO los datos de
pacientes (Recetas Cheque, que contienen RUT y están sujetos a la Ley 19.628).

Estructura en Drive:
  Farmacia AA/
    1 - App Pedidos/        Consolidado_AA_MAESTRO.xlsx + Resumen_Pedidos_AA.xlsx
    2 - Gestion Territorial/  último rango al frente + Historial/<rango>/
    4 - Auditoria Prescripcion/  Auditoria_Prescripcion_Resumen.xlsx
    5 - Reposicion/         Reposicion_DiasHabiles_AA.xlsx

Primera vez (requiere Google Cloud credentials.json):
  py publicar_drive.py --setup

Uso normal (token ya generado):
  py publicar_drive.py           # sube todo
  py publicar_drive.py --solo-app  --solo-gt  --solo-auditoria  --solo-rep
"""
import argparse, base64, glob, hashlib, io, json, os, re, sys
from datetime import datetime

for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

WORK_DIR = os.path.dirname(os.path.abspath(__file__))
CREDS_FILE = os.path.join(WORK_DIR, "credentials.json")
TOKEN_FILE = os.path.join(WORK_DIR, "token_drive.json")

NOMBRE_RAIZ = "Farmacia AA"
SUB_APP     = "1 - App Pedidos"
SUB_GT      = "2 - Gestion Territorial"
SUB_AUDIT   = "4 - Auditoria Prescripcion"
SUB_REP     = "5 - Reposicion"

# IDs de carpetas ya creadas en Drive (evita duplicados en búsquedas)
_FOLDER_CACHE_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_drive_folders.json")

_RANGO_RE = re.compile(r"^\d{2}-\d{2}-\d{4}_\d{2}-\d{2}-\d{4}$")
SCOPES = ["https://www.googleapis.com/auth/drive"]


# ── OAuth ─────────────────────────────────────────────────────────────────────
def _get_service():
    try:
        from googleapiclient.discovery import build
        from google.oauth2.credentials import Credentials
        from google.auth.transport.requests import Request
        from google_auth_oauthlib.flow import InstalledAppFlow
    except ImportError:
        print("[ERROR] Faltan dependencias de Google Drive:")
        print("  py -m pip install google-api-python-client google-auth-oauthlib google-auth-httplib2")
        sys.exit(1)

    creds = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists(CREDS_FILE):
                print(f"[ERROR] No se encontró credentials.json en:\n  {WORK_DIR}")
                print("\nPasos para obtenerlo:")
                print("  1. Abre console.cloud.google.com")
                print("  2. Crea un proyecto → Habilita 'Google Drive API'")
                print("  3. Credenciales → 'OAuth 2.0 Client ID' (tipo: Desktop App)")
                print("  4. Descarga el JSON y guárdalo como credentials.json en:")
                print(f"     {WORK_DIR}")
                print("  5. Ejecuta:  py publicar_drive.py --setup")
                sys.exit(1)
            flow = InstalledAppFlow.from_client_secrets_file(CREDS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, "w", encoding="utf-8") as fh:
            fh.write(creds.to_json())
        print(f"  Token guardado: {TOKEN_FILE}")

    return build("drive", "v3", credentials=creds, cache_discovery=False)


# ── API Drive helpers ─────────────────────────────────────────────────────────
def _buscar_carpeta(service, nombre, parent_id=None):
    q = f"name='{nombre}' and mimeType='application/vnd.google-apps.folder' and trashed=false"
    if parent_id:
        q += f" and '{parent_id}' in parents"
    r = service.files().list(q=q, fields="files(id,name)", pageSize=1).execute()
    files = r.get("files", [])
    return files[0]["id"] if files else None


def _crear_carpeta(service, nombre, parent_id=None):
    meta = {
        "name": nombre,
        "mimeType": "application/vnd.google-apps.folder",
    }
    if parent_id:
        meta["parents"] = [parent_id]
    f = service.files().create(body=meta, fields="id").execute()
    return f["id"]


def _obtener_o_crear_carpeta(service, nombre, parent_id=None, cache=None):
    clave = (nombre, parent_id)
    if cache is not None and clave in cache:
        return cache[clave]
    fid = _buscar_carpeta(service, nombre, parent_id)
    if not fid:
        fid = _crear_carpeta(service, nombre, parent_id)
        print(f"  [Drive] Carpeta creada: {nombre}")
    if cache is not None:
        cache[clave] = fid
    return fid


def _md5(path):
    h = hashlib.md5()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def _buscar_archivo(service, nombre, parent_id):
    q = f"name='{nombre}' and '{parent_id}' in parents and trashed=false"
    r = service.files().list(q=q, fields="files(id,md5Checksum)", pageSize=1).execute()
    files = r.get("files", [])
    return (files[0]["id"], files[0].get("md5Checksum")) if files else (None, None)


def _subir(service, local_path, folder_id, nuevo_nombre=None, stats=None):
    """Sube local_path a folder_id en Drive. Salta si MD5 coincide."""
    from googleapiclient.http import MediaFileUpload

    nombre = nuevo_nombre or os.path.basename(local_path)
    md5_local = _md5(local_path)
    fid, md5_drive = _buscar_archivo(service, nombre, folder_id)

    if fid and md5_drive == md5_local:
        if stats:
            stats["skip"] += 1
        return "skip"

    mime, _ = _guess_mime(local_path)
    media = MediaFileUpload(local_path, mimetype=mime, resumable=True)
    if fid:
        service.files().update(fileId=fid, body={"name": nombre}, media_body=media).execute()
        if stats:
            stats["ok"] += 1
        return "updated"
    else:
        meta = {"name": nombre, "parents": [folder_id]}
        service.files().create(body=meta, media_body=media, fields="id").execute()
        if stats:
            stats["ok"] += 1
        return "created"


def _guess_mime(path):
    ext = os.path.splitext(path)[1].lower()
    table = {
        ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        ".pdf":  "application/pdf",
        ".csv":  "text/csv",
        ".json": "application/json",
        ".txt":  "text/plain",
        ".png":  "image/png",
    }
    return table.get(ext, "application/octet-stream"), ext


def _mas_reciente(patron):
    cand = [f for f in glob.glob(patron)
            if not os.path.basename(f).startswith("~$")]
    return max(cand, key=os.path.getmtime) if cand else None


# ── Sincronizadores ───────────────────────────────────────────────────────────
def sync_app(service, raiz_id, stats):
    fid = _obtener_o_crear_carpeta(service, SUB_APP, raiz_id)
    for patron, canon in [
        ("Consolidado_AA_MAESTRO*.xlsx", "Consolidado_AA_MAESTRO.xlsx"),
        ("Resumen_Pedidos_AA*.xlsx",      "Resumen_Pedidos_AA.xlsx"),
    ]:
        src = _mas_reciente(os.path.join(WORK_DIR, patron))
        if src:
            r = _subir(service, src, fid, nuevo_nombre=canon, stats=stats)
            print(f"  {canon}: {r}")
        else:
            print(f"  {canon}: no encontrado, omitido")


def sync_gt(service, raiz_id, stats, cache):
    src_base = os.path.join(WORK_DIR, "out_gt")
    if not os.path.isdir(src_base):
        print("  [GT] sin out_gt todavía")
        return
    rangos = sorted([d for d in glob.glob(os.path.join(src_base, "*"))
                     if os.path.isdir(d) and _RANGO_RE.match(os.path.basename(d))],
                    key=os.path.getmtime)
    if not rangos:
        print("  [GT] sin rangos en out_gt")
        return

    fid_gt   = _obtener_o_crear_carpeta(service, SUB_GT,      raiz_id, cache)
    fid_hist = _obtener_o_crear_carpeta(service, "Historial", fid_gt,  cache)

    for d in rangos:
        nombre_rango = os.path.basename(d)
        fid_r = _obtener_o_crear_carpeta(service, nombre_rango, fid_hist, cache)
        for f in glob.glob(os.path.join(d, "*.xlsx")) + glob.glob(os.path.join(d, "*.json")):
            r = _subir(service, f, fid_r, stats=stats)
            if r != "skip":
                print(f"  GT/{nombre_rango}/{os.path.basename(f)}: {r}")

    # Último rango al frente (directamente bajo 2 - Gestion Territorial)
    ultimo = rangos[-1]
    for f in glob.glob(os.path.join(ultimo, "*.xlsx")):
        r = _subir(service, f, fid_gt, stats=stats)
        if r != "skip":
            print(f"  GT (frente)/{os.path.basename(f)}: {r}")

    print(f"  [GT] {len(rangos)} rango(s) · último «{os.path.basename(ultimo)}» al frente")


def sync_auditoria(service, raiz_id, stats):
    fid = _obtener_o_crear_carpeta(service, SUB_AUDIT, raiz_id)
    src = _mas_reciente(os.path.join(WORK_DIR, "Auditoria_Prescripcion_Resumen*.xlsx"))
    if src:
        r = _subir(service, src, fid, nuevo_nombre="Auditoria_Prescripcion_Resumen.xlsx", stats=stats)
        print(f"  Auditoria_Prescripcion_Resumen.xlsx: {r}")
    else:
        print("  Auditoria_Prescripcion_Resumen.xlsx: no encontrado, omitido")


def sync_reposicion(service, raiz_id, stats):
    fid = _obtener_o_crear_carpeta(service, SUB_REP, raiz_id)
    src = _mas_reciente(os.path.join(WORK_DIR, "Reposicion_DiasHabiles_AA*.xlsx"))
    if src:
        r = _subir(service, src, fid, nuevo_nombre="Reposicion_DiasHabiles_AA.xlsx", stats=stats)
        print(f"  Reposicion_DiasHabiles_AA.xlsx: {r}")
    else:
        print("  Reposicion_DiasHabiles_AA.xlsx: no encontrado, omitido")


# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    ap = argparse.ArgumentParser(description="Sube salidas Farmacia AA a Google Drive")
    ap.add_argument("--setup", action="store_true",
                    help="Fuerza el flujo OAuth (primera vez o token expirado)")
    ap.add_argument("--solo-app",      action="store_true")
    ap.add_argument("--solo-gt",       action="store_true")
    ap.add_argument("--solo-auditoria",action="store_true")
    ap.add_argument("--solo-rep",      action="store_true")
    a = ap.parse_args()

    if a.setup and os.path.exists(TOKEN_FILE):
        os.remove(TOKEN_FILE)
        print("[SETUP] Token eliminado — se iniciará flujo OAuth desde cero.")

    print("\n[Drive] Conectando con Google Drive...")
    svc = _get_service()

    # Carpeta raíz "Farmacia AA" — usa IDs pre-conocidos si existen
    cache = {}
    known = {}
    if os.path.exists(_FOLDER_CACHE_FILE):
        try:
            with open(_FOLDER_CACHE_FILE, encoding="utf-8") as _f:
                known = json.load(_f)
        except Exception:
            pass

    if NOMBRE_RAIZ in known:
        raiz_id = known[NOMBRE_RAIZ]
        # Pre-carga sub-carpetas fijas en cache para evitar búsquedas API
        for sub in (SUB_APP, SUB_GT, SUB_AUDIT, SUB_REP):
            if sub in known:
                cache[(sub, raiz_id)] = known[sub]
    else:
        raiz_id = _obtener_o_crear_carpeta(svc, NOMBRE_RAIZ, cache=cache)

    stats = {"ok": 0, "skip": 0, "fail": 0}

    todos = not any([a.solo_app, a.solo_gt, a.solo_auditoria, a.solo_rep])

    print(f"\n[Drive] Subiendo a «{NOMBRE_RAIZ}» (id={raiz_id[:8]}…)")
    ts = datetime.now().strftime("%H:%M:%S")
    print(f"        {ts}\n")

    if todos or a.solo_app:
        print(f"  ── {SUB_APP} ──")
        sync_app(svc, raiz_id, stats)

    if todos or a.solo_gt:
        print(f"\n  ── {SUB_GT} ──")
        sync_gt(svc, raiz_id, stats, cache)

    if todos or a.solo_auditoria:
        print(f"\n  ── {SUB_AUDIT} ──")
        sync_auditoria(svc, raiz_id, stats)

    if todos or a.solo_rep:
        print(f"\n  ── {SUB_REP} ──")
        sync_reposicion(svc, raiz_id, stats)

    total = stats["ok"] + stats["skip"] + stats["fail"]
    print(f"\n[Drive] Listo: {stats['ok']} subidos · {stats['skip']} sin cambios · "
          f"{stats['fail']} errores  ({total} total)")
    if stats["fail"]:
        sys.exit(1)


if __name__ == "__main__":
    main()
