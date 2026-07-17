#!/usr/bin/env python3
"""
publicar_drive.py — Sube las salidas de Farmacia AA a Google Drive.

Sube los mismos archivos que publicar_escritorio.py EXCEPTO los datos de
pacientes (Recetas Cheque, que contienen RUT y están sujetos a la Ley 19.628).

Todo .xlsx se sube convertido a Google Sheets nativo (no como archivo Excel
crudo) — se abre editable en el navegador sin descargar. PDF/JSON no se tocan.

Estructura en Drive:
  Farmacia AA/
    1 - App Pedidos/          Consolidado_AA_MAESTRO.xlsx + Resumen_Pedidos_AA.xlsx
    2 - Gestion Territorial/
        <ESTAB DESTINO>/      últimas planillas + letrero por establecimiento
        Historial/
            <fecha>/
                <ESTAB DESTINO>/  planillas + letrero históricos
                Cruce_GT_Clasificacion.xlsx
    3 - Pedido Fusionado/     Pedido_Fusion_AA.xlsx
    4 - Auditoria Prescripcion/  Auditoria_Prescripcion_Resumen.xlsx
    6 - Centinela/<Sxx>/      centinela_Sxx.json + centinela_Sxx.pdf por semana
    7 - Programacion AA/      Resumen_Programacion_AA.xlsx (conteo vs programación)

Primera vez (requiere Google Cloud credentials.json):
  py publicar_drive.py --setup

Uso normal (token ya generado):
  py publicar_drive.py           # sube todo
  py publicar_drive.py --solo-app  --solo-gt  --solo-pedido  --solo-auditoria  --solo-centinela  --solo-programacion
"""
import argparse, glob, hashlib, json, os, re, sys
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
SUB_PEDIDO  = "3 - Pedido Fusionado"
SUB_AUDIT   = "4 - Auditoria Prescripcion"
SUB_CENTINELA = "6 - Centinela"
SUB_PROG    = "7 - Programacion AA"

# IDs de carpetas ya creadas en Drive (evita duplicados en búsquedas)
_FOLDER_CACHE_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_drive_folders.json")

_RANGO_RE = re.compile(r"^\d{2}-\d{2}-\d{4}_\d{2}-\d{2}-\d{4}$")

# Detecta el establecimiento de destino en un nombre de archivo GT
# Patrones: <DESTINO>_Planilla.xlsx, <DESTINO>_Letrero.pdf,
#           <DESTINO>_Controlados_Planilla.xlsx, <DESTINO>_Verificacion.xlsx
_TIPO_GT_RE = re.compile(
    r"^(.+?)_(Planilla|Letrero|Controlados_Planilla|Verificacion)\.(xlsx|pdf)$",
    re.IGNORECASE,
)

def _destino_de_archivo(nombre):
    """Devuelve el nombre del establecimiento (con espacios) o None si es archivo general."""
    m = _TIPO_GT_RE.match(nombre)
    return m.group(1).replace("_", " ") if m else None
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
    nombre_q = nombre.replace("'", "\\'")
    q = f"name='{nombre_q}' and mimeType='application/vnd.google-apps.folder' and trashed=false"
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
    nombre_q = nombre.replace("'", "\\'")
    q = f"name='{nombre_q}' and '{parent_id}' in parents and trashed=false"
    r = service.files().list(q=q, fields="files(id,md5Checksum,mimeType)", pageSize=1).execute()
    files = r.get("files", [])
    if not files:
        return None, None, None
    f = files[0]
    return f["id"], f.get("md5Checksum"), f.get("mimeType")


SHEETS_MIME = "application/vnd.google-apps.spreadsheet"


def _subir(service, local_path, folder_id, nuevo_nombre=None, stats=None):
    """Sube local_path a folder_id en Drive. Los .xlsx se suben como Google Sheets
    NATIVO (Drive los convierte al vuelo) — se abren editables en el navegador,
    sin descargar ni necesitar Excel.

    Probado en vivo (2026-07-14) — la conversión tiene dos restricciones reales
    de la API de Drive, no documentadas con claridad:
      1) Solo funciona en subida multipart simple, NO resumable
         ("Invalid MIME type provided for the uploaded content" si resumable=True).
      2) `files.update()` NO puede convertir un archivo EXISTENTE que todavía es
         xlsx crudo a Sheets (mismo error) — solo `files.create()` puede fijar
         el mimeType destino. Un archivo .xlsx ya subido antes de este cambio se
         migra por única vez: se borra y se vuelve a crear como Sheet (cambia su
         fileId esa vez). Una vez que YA es Sheet, sí se puede actualizar su
         contenido en el mismo fileId sin volver a tocar el mimeType.

    Los Sheets nativos no tienen md5Checksum, así que esos SIEMPRE se re-suben en
    cada corrida (no hay forma barata de detectar 'sin cambios' ya convertido);
    el resto de formatos (pdf/json/txt/png) sigue el dedup por MD5 de siempre.

    IMPORTANTE: Drive QUITA la extensión ".xlsx" del nombre al convertir a Sheets
    nativo (p.ej. "Consolidado_AA_MAESTRO.xlsx" queda guardado como
    "Consolidado_AA_MAESTRO"). Por eso la búsqueda/creación de estos archivos usa
    el nombre SIN extensión — si se buscara con ".xlsx" nunca encontraría el
    archivo ya convertido y crearía un duplicado nuevo en cada corrida (bug real,
    visto en vivo: 6 duplicados en 2 corridas antes de este fix).

    Captura cualquier error de subida (red, cuota, archivo bloqueado, etc.) para
    que una falla puntual no aborte el resto de la sincronización."""
    from googleapiclient.http import MediaFileUpload

    nombre_local = nuevo_nombre or os.path.basename(local_path)
    try:
        mime, ext = _guess_mime(local_path)
        a_sheets = ext == ".xlsx"
        nombre = (nombre_local[:-len(ext)] if a_sheets and nombre_local.lower().endswith(ext)
                  else nombre_local)

        md5_local = _md5(local_path)
        fid, md5_drive, mime_drive = _buscar_archivo(service, nombre, folder_id)
        if a_sheets and fid is None and nombre_local != nombre:
            # No estaba con el nombre sin extensión (ya convertido) — puede seguir
            # en crudo con ".xlsx" en el nombre (subido antes de este fix).
            fid, md5_drive, mime_drive = _buscar_archivo(service, nombre_local, folder_id)
        ya_convertido = mime_drive == SHEETS_MIME

        if fid and md5_drive and md5_drive == md5_local and (not a_sheets or ya_convertido):
            if stats:
                stats["skip"] += 1
            return "skip"

        media = MediaFileUpload(local_path, mimetype=mime, resumable=not a_sheets)

        if fid and a_sheets and not ya_convertido:
            # Migración única: no se puede convertir un xlsx crudo con update().
            service.files().delete(fileId=fid).execute()
            fid = None

        if fid:
            body = {"name": nombre}   # ya es Sheet (o no aplica conversión): sin mimeType
            service.files().update(fileId=fid, body=body, media_body=media).execute()
            if stats:
                stats["ok"] += 1
            return "updated"
        else:
            meta = {"name": nombre, "parents": [folder_id]}
            if a_sheets:
                meta["mimeType"] = SHEETS_MIME
            service.files().create(body=meta, media_body=media, fields="id").execute()
            if stats:
                stats["ok"] += 1
            return "created"
    except Exception as e:
        print(f"  [error] {nombre}: {e}")
        if stats:
            stats["fail"] += 1
        return "fail"


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
def sync_app(service, raiz_id, stats, cache=None):
    fid = _obtener_o_crear_carpeta(service, SUB_APP, raiz_id, cache)
    for patron, canon in [
        ("Consolidado_AA_MAESTRO*.xlsx", "Consolidado_AA_MAESTRO.xlsx"),
        ("Resumen_Pedidos_AA*.xlsx",      "Resumen_Pedidos_AA.xlsx"),
        ("SGLI_Historico_*.xlsx",         "SGLI_Historico.xlsx"),
    ]:
        src = _mas_reciente(os.path.join(WORK_DIR, patron))
        if src:
            r = _subir(service, src, fid, nuevo_nombre=canon, stats=stats)
            print(f"  {canon}: {r}")
        else:
            print(f"  {canon}: no encontrado, omitido")


def _fecha_fin_rango(nombre_carpeta):
    """Extrae la fecha de fin del nombre de carpeta 'DD-MM-YYYY_DD-MM-YYYY' como 'DD-MM-YYYY'.
    Si el formato no coincide, cae a mtime de la carpeta."""
    m = re.match(r"(\d{2}-\d{2}-\d{4})_(\d{2}-\d{2}-\d{4})$", nombre_carpeta)
    return m.group(2) if m else None


def _subir_gt_rango(service, rango_dir, parent_fid, stats, cache, prefijo_log):
    """Sube todos los archivos de un rango GT a parent_fid, organizados por establecimiento.
    Archivos con destino conocido van a <parent>/<ESTAB>/. Archivos generales van directo."""
    archivos = sorted(
        glob.glob(os.path.join(rango_dir, "*.xlsx")) +
        glob.glob(os.path.join(rango_dir, "*.pdf"))
    )
    estabs_vistos = set()
    for f in archivos:
        nb = os.path.basename(f)
        if nb.startswith("~$"):
            continue
        destino = _destino_de_archivo(nb)
        if destino:
            estabs_vistos.add(destino)
            fid_estab = _obtener_o_crear_carpeta(service, destino, parent_fid, cache)
            r = _subir(service, f, fid_estab, stats=stats)
            if r != "skip":
                print(f"  {prefijo_log}/{destino}/{nb}: {r}")
        else:
            # Archivo general (Cruce_GT_Clasificacion, etc.) — al nivel del rango
            r = _subir(service, f, parent_fid, stats=stats)
            if r != "skip":
                print(f"  {prefijo_log}/{nb}: {r}")
    return sorted(estabs_vistos)


_GT_SYNC_STATE_FILE = os.path.join(WORK_DIR, "_gt_historial_sync.json")


def _cargar_gt_sincronizados():
    """Nombres de carpetas out_gt/<rango> que YA quedaron completas en Drive
    Historial en una corrida anterior. Un rango es inmutable una vez que su
    ventana de fechas dejó de ser la más nueva (AUTO_SSASUR no vuelve a
    generar el mismo rango dos veces en uso normal), así que no hace falta
    volver a listarlo ni re-subir sus archivos nunca más."""
    if os.path.exists(_GT_SYNC_STATE_FILE):
        try:
            with open(_GT_SYNC_STATE_FILE, encoding="utf-8") as f:
                return set(json.load(f))
        except Exception:
            pass
    return set()


def _guardar_gt_sincronizados(nombres):
    with open(_GT_SYNC_STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(sorted(nombres), f, ensure_ascii=False, indent=2)


def sync_gt(service, raiz_id, stats, cache):
    src_base = os.path.join(WORK_DIR, "out_gt")
    if not os.path.isdir(src_base):
        print("  [GT] sin out_gt todavía")
        return
    def _clave_cronologica(d):
        fin = _fecha_fin_rango(os.path.basename(d))
        return datetime.strptime(fin, "%d-%m-%Y") if fin else datetime.fromtimestamp(os.path.getmtime(d))

    rangos = sorted(
        [d for d in glob.glob(os.path.join(src_base, "*"))
         if os.path.isdir(d) and _RANGO_RE.match(os.path.basename(d))],
        key=_clave_cronologica,
    )
    if not rangos:
        print("  [GT] sin rangos en out_gt")
        return

    fid_gt   = _obtener_o_crear_carpeta(service, SUB_GT,      raiz_id, cache)
    fid_hist = _obtener_o_crear_carpeta(service, "Historial", fid_gt,  cache)

    # ── Historial: un subfolder por rango. Los rangos ANTERIORES al último ya
    # quedaron con su ventana de fechas fija — si ya se subieron completos en
    # una corrida previa, se SALTAN enteros (ni se listan sus archivos) en vez
    # de re-subirlos. Antes esto no pasaba: como los .xlsx se convierten a
    # Google Sheets nativo (sin md5Checksum), el dedup por hash no aplicaba y
    # CADA corrida de AUTO_SSASUR volvía a re-subir TODO el histórico
    # acumulado (9 rangos y creciendo, cada uno traslapando ~13 de sus 14 días
    # con el anterior). El último rango (ventana viva, todavía puede cambiar)
    # se sigue subiendo siempre.
    sincronizados = _cargar_gt_sincronizados()
    nuevos_sincronizados = set(sincronizados)
    anteriores, ultimo = rangos[:-1], rangos[-1]
    saltados = 0
    for d in anteriores:
        nombre_rango = os.path.basename(d)
        if nombre_rango in sincronizados:
            saltados += 1
            continue
        fecha_rango = _fecha_fin_rango(nombre_rango) or \
                      datetime.fromtimestamp(os.path.getmtime(d)).strftime("%d-%m-%Y")
        fid_r = _obtener_o_crear_carpeta(service, fecha_rango, fid_hist, cache)
        fails_antes = stats["fail"]
        _subir_gt_rango(service, d, fid_r, stats, cache,
                        prefijo_log=f"Historial/{fecha_rango}")
        if stats["fail"] == fails_antes:
            nuevos_sincronizados.add(nombre_rango)
    if saltados:
        print(f"  [GT] {saltados} rango(s) histórico(s) ya sincronizados — omitidos")

    fecha_ultimo = _fecha_fin_rango(os.path.basename(ultimo)) or \
                   datetime.fromtimestamp(os.path.getmtime(ultimo)).strftime("%d-%m-%Y")
    fid_r_ultimo = _obtener_o_crear_carpeta(service, fecha_ultimo, fid_hist, cache)
    _subir_gt_rango(service, ultimo, fid_r_ultimo, stats, cache,
                    prefijo_log=f"Historial/{fecha_ultimo}")

    _guardar_gt_sincronizados(nuevos_sincronizados)

    # ── Frente: último rango organizado por establecimiento bajo GT raíz ──
    estabs = _subir_gt_rango(service, ultimo, fid_gt, stats, cache,
                             prefijo_log="frente")
    estabs_str = ", ".join(estabs) if estabs else "(ninguno)"
    print(f"  [GT] {len(rangos)} rango(s) · último «{fecha_ultimo}» · "
          f"establecimientos: {estabs_str}")


def sync_auditoria(service, raiz_id, stats, cache=None):
    fid = _obtener_o_crear_carpeta(service, SUB_AUDIT, raiz_id, cache)
    src = _mas_reciente(os.path.join(WORK_DIR, "Auditoria_Prescripcion_Resumen*.xlsx"))
    if src:
        r = _subir(service, src, fid, nuevo_nombre="Auditoria_Prescripcion_Resumen.xlsx", stats=stats)
        print(f"  Auditoria_Prescripcion_Resumen.xlsx: {r}")
    else:
        print("  Auditoria_Prescripcion_Resumen.xlsx: no encontrado, omitido")


def sync_pedido(service, raiz_id, stats, cache=None):
    fid = _obtener_o_crear_carpeta(service, SUB_PEDIDO, raiz_id, cache)
    src = _mas_reciente(os.path.join(WORK_DIR, "Pedido_Fusion_AA*.xlsx"))
    if src:
        r = _subir(service, src, fid, nuevo_nombre="Pedido_Fusion_AA.xlsx", stats=stats)
        print(f"  Pedido_Fusion_AA.xlsx: {r}")
    else:
        print("  Pedido_Fusion_AA.xlsx: no encontrado, omitido")


def sync_centinela(service, raiz_id, stats, cache=None):
    fid = _obtener_o_crear_carpeta(service, SUB_CENTINELA, raiz_id, cache)
    src_base = os.path.join(WORK_DIR, "Centinela_Reportes")
    if not os.path.isdir(src_base):
        print("  [Centinela] sin Centinela_Reportes todavía")
        return
    semanas = sorted(
        d for d in glob.glob(os.path.join(src_base, "*")) if os.path.isdir(d)
    )
    if not semanas:
        print("  [Centinela] sin semanas en Centinela_Reportes")
        return
    for d in semanas:
        nombre_semana = os.path.basename(d)
        fid_semana = _obtener_o_crear_carpeta(service, nombre_semana, fid, cache)
        archivos = sorted(
            glob.glob(os.path.join(d, "*.json")) + glob.glob(os.path.join(d, "*.pdf"))
        )
        for f in archivos:
            r = _subir(service, f, fid_semana, stats=stats)
            if r != "skip":
                print(f"  {nombre_semana}/{os.path.basename(f)}: {r}")
    print(f"  [Centinela] {len(semanas)} semana(s): {', '.join(os.path.basename(d) for d in semanas)}")


def sync_programacion(service, raiz_id, stats, cache=None):
    fid = _obtener_o_crear_carpeta(service, SUB_PROG, raiz_id, cache)
    # El Resumen (post-conteo) es la salida final; mientras no exista, se sube
    # igual la planilla del ciclo (pre-conteo) para tenerla disponible fuera
    # del equipo local aunque el conteo físico todavía no se haya hecho.
    src = _mas_reciente(os.path.join(WORK_DIR, "Programacion_AA", "Resumen_Programacion_AA*.xlsx"))
    if src:
        r = _subir(service, src, fid, nuevo_nombre="Resumen_Programacion_AA.xlsx", stats=stats)
        print(f"  Resumen_Programacion_AA.xlsx: {r}")
        return
    src = _mas_reciente(os.path.join(WORK_DIR, "Programacion_AA", "Programacion_AA_*.xlsx"))
    if src:
        r = _subir(service, src, fid, nuevo_nombre="Programacion_AA.xlsx", stats=stats)
        print(f"  Programacion_AA.xlsx: {r}  (planilla del ciclo, sin conteo aplicado todavía)")
    else:
        print("  Programacion_AA: no encontrado, omitido (corre programacion_aa.py)")


# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    ap = argparse.ArgumentParser(description="Sube salidas Farmacia AA a Google Drive")
    ap.add_argument("--setup", action="store_true",
                    help="Fuerza el flujo OAuth (primera vez o token expirado)")
    ap.add_argument("--solo-app",      action="store_true")
    ap.add_argument("--solo-gt",       action="store_true")
    ap.add_argument("--solo-pedido",   action="store_true")
    ap.add_argument("--solo-auditoria",action="store_true")
    ap.add_argument("--solo-centinela",action="store_true")
    ap.add_argument("--solo-programacion", action="store_true")
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
        # Pre-carga sub-carpetas fijas para evitar búsquedas API
        for sub in (SUB_APP, SUB_GT, SUB_PEDIDO, SUB_AUDIT, SUB_CENTINELA, SUB_PROG):
            if sub in known:
                cache[(sub, raiz_id)] = known[sub]
        # Pre-carga carpetas de Historial + rangos si están en el JSON
        gt_id = known.get(SUB_GT)
        hist_key = f"{SUB_GT}/Historial"
        hist_id = known.get(hist_key)
        if gt_id and hist_id:
            cache[("Historial", gt_id)] = hist_id
            for full_key, fid in known.items():
                if full_key.startswith(f"{hist_key}/"):
                    rango = full_key[len(hist_key) + 1:]
                    cache[(rango, hist_id)] = fid
    else:
        raiz_id = _obtener_o_crear_carpeta(svc, NOMBRE_RAIZ, cache=cache)

    stats = {"ok": 0, "skip": 0, "fail": 0}

    todos = not any([a.solo_app, a.solo_gt, a.solo_pedido, a.solo_auditoria,
                     a.solo_centinela, a.solo_programacion])

    print(f"\n[Drive] Subiendo a «{NOMBRE_RAIZ}» (id={raiz_id[:8]}…)")
    ts = datetime.now().strftime("%H:%M:%S")
    print(f"        {ts}\n")

    if todos or a.solo_app:
        print(f"  ── {SUB_APP} ──")
        sync_app(svc, raiz_id, stats, cache)

    if todos or a.solo_gt:
        print(f"\n  ── {SUB_GT} ──")
        sync_gt(svc, raiz_id, stats, cache)

    if todos or a.solo_pedido:
        print(f"\n  ── {SUB_PEDIDO} ──")
        sync_pedido(svc, raiz_id, stats, cache)

    if todos or a.solo_auditoria:
        print(f"\n  ── {SUB_AUDIT} ──")
        sync_auditoria(svc, raiz_id, stats, cache)

    if todos or a.solo_centinela:
        print(f"\n  ── {SUB_CENTINELA} ──")
        sync_centinela(svc, raiz_id, stats, cache)

    if todos or a.solo_programacion:
        print(f"\n  ── {SUB_PROG} ──")
        sync_programacion(svc, raiz_id, stats, cache)

    total = stats["ok"] + stats["skip"] + stats["fail"]
    print(f"\n[Drive] Listo: {stats['ok']} subidos · {stats['skip']} sin cambios · "
          f"{stats['fail']} errores  ({total} total)")

    # Persiste IDs de carpetas nuevas descubiertas en esta corrida
    nuevos = {nombre: fid for (nombre, pid), fid in cache.items() if pid == raiz_id}
    nuevos[NOMBRE_RAIZ] = raiz_id
    # También persiste la jerarquía SUB_GT/Historial/<rango> (no son hijos directos
    # de la raíz, así que el filtro por pid==raiz_id de arriba no las captura — sin
    # esto, esas carpetas se re-buscan vía API en cada ejecución para siempre).
    gt_folder_id = nuevos.get(SUB_GT)
    hist_folder_id = cache.get(("Historial", gt_folder_id)) if gt_folder_id else None
    if hist_folder_id:
        nuevos[f"{SUB_GT}/Historial"] = hist_folder_id
        for (nombre, pid), fid in cache.items():
            if pid == hist_folder_id:
                nuevos[f"{SUB_GT}/Historial/{nombre}"] = fid
    if os.path.exists(_FOLDER_CACHE_FILE):
        try:
            with open(_FOLDER_CACHE_FILE, encoding="utf-8") as _f:
                existing = json.load(_f)
        except Exception:
            existing = {}
    else:
        existing = {}
    merged = {**existing, **{k: v for k, v in nuevos.items() if k not in existing}}
    if merged != existing:
        with open(_FOLDER_CACHE_FILE, "w", encoding="utf-8") as _f:
            json.dump(merged, _f, ensure_ascii=False, indent=2)
        print(f"  [cache] {_FOLDER_CACHE_FILE.split(os.sep)[-1]} actualizado con {len(merged)-len(existing)} entradas nuevas")

    if stats["fail"]:
        sys.exit(1)


if __name__ == "__main__":
    main()
