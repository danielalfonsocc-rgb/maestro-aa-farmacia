#!/usr/bin/env python3
"""
publicar_drive.py — Sube las salidas de Farmacia AA a Google Drive.

Sube los mismos archivos que publicar_escritorio.py EXCEPTO los datos de
pacientes (Recetas Cheque, que contienen RUT y están sujetos a la Ley 19.628).

EXCEPCIÓN — Clozapina: el usuario autorizó explícitamente (2026-07-23) subir
Clozapina_Reportes/ (RUT + nombre + diagnóstico psiquiátrico implícito) a su
PROPIA carpeta de nivel superior en Drive ("Clozapina - CONFIDENCIAL"), fuera
de la carpeta raíz "Farmacia AA" para no heredar su lista de compartidos.
Solo sube con el flag explícito --solo-clozapina, JAMÁS como parte de una
sincronización general (mismo criterio que [[recetas-cheque-drive-excepcion]]
y [[gt-sheets-rut-drive-excepcion]] en memoria del proyecto).

Todo .xlsx se sube convertido a Google Sheets nativo (no como archivo Excel
crudo) — se abre editable en el navegador sin descargar. PDF/JSON no se tocan.

Estructura en Drive:
  Farmacia AA/
    1 - App Pedidos/          Consolidado_AA_MAESTRO.xlsx + Resumen_Pedidos_AA.xlsx
    2 - Gestion Territorial/
        <ESTAB DESTINO>/
            Revisión de Solicitudes/   feedback + recetas solicitadas
            Nóminas de Envío/
                <DD-MM-YYYY>/          planillas + letrero de cada envío (historial completo;
                                       sin carpetas Historial/Acumulados desde el 21-07-2026)
    3 - Pedido Fusionado/     Pedido_Fusion_AA.xlsx + Pedido_Fusion_Simple_AA.xlsx
    4 - Auditoria Prescripcion/  Auditoria_Prescripcion_Resumen.xlsx
    6 - Centinela/<Sxx>/      centinela_Sxx.json + centinela_Sxx.pdf por semana
    7 - Programacion AA/      Resumen_Programacion_AA.xlsx (conteo vs programación)
    8 - Servicios Farmaceuticos/<MES AÑO>/  Resumen_Servicios_Farmaceuticos_<mes>_<año>.xlsx
                              (agregado QF × actividad, SIN RUT — el reporte crudo de
                              Agenda Médica con RUT de paciente nunca se sube, igual que
                              informe_completo_recetas*.csv y reporte_de_stock*.xlsx)
    9 - Centinela Inyectables SM/<fecha>/  reporte_inyectables_sm_<fecha>.html
                              (stock de antipsicóticos de depósito, SIN RUT de paciente)

Primera vez (requiere Google Cloud credentials.json):
  py publicar_drive.py --setup

Uso normal (token ya generado):
  py publicar_drive.py           # sube todo
  py publicar_drive.py --solo-app  --solo-gt  --solo-pedido  --solo-auditoria  --solo-centinela  --solo-programacion  --solo-centinela-sm
"""
import argparse, glob, hashlib, json, os, re, sys
from datetime import datetime

import gt_maestro as _GM  # solo para MESES_ES (mismo esquema de carpetas que agregar_gt_manual.py)

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
SUB_SERVICIOS = "8 - Servicios Farmaceuticos"
SUB_CENTINELA_SM = "9 - Centinela Inyectables SM"

# Clozapina: RUT + nombre + diagnóstico psiquiátrico implícito (más sensible
# que el RUT genérico de Recetas Cheque/GT) — por eso NO va dentro de la
# carpeta raíz "Farmacia AA" (que ya está compartida con más gente), sino en
# su PROPIA carpeta de nivel superior en el Drive de la cuenta autenticada,
# privada por defecto (nadie más la ve salvo que la QF la comparta a mano).
# Autorizado explícitamente por el usuario 2026-07-23 — solo corre con el
# flag --solo-clozapina, NUNCA como parte de la sincronización general.
NOMBRE_RAIZ_CLOZAPINA = "Clozapina - CONFIDENCIAL"

# IDs de carpetas ya creadas en Drive (evita duplicados en búsquedas). El de
# Clozapina va en un archivo APARTE (gitignored) — el nombre "Clozapina -
# CONFIDENCIAL" no debe aparecer en _drive_folders.json porque ese SÍ está
# versionado en el repo público (revelaría que se maneja data psiquiátrica).
_FOLDER_CACHE_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_drive_folders.json")
_FOLDER_CACHE_FILE_CLOZAPINA = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_drive_folders_clozapina.json")

# Pedido por el usuario 04-08-2026: copias SELLADAS de recetas GT (con
# anotaciones de anular/pendientes/observación escritas encima) van a su
# PROPIA carpeta Drive privada, fuera de "Farmacia AA" — mismo patrón que
# Clozapina. Solo corre con --solo-gt-confidencial, NUNCA con la
# sincronización general ni junto con --solo-gt (que sube las recetas
# limpias, sin escritura, a la carpeta normal).
NOMBRE_RAIZ_GT_CONFIDENCIAL = "Gestion Territorial - CONFIDENCIAL (anotado)"
_FOLDER_CACHE_FILE_GT_CONFIDENCIAL = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                                   "_drive_folders_gt_confidencial.json")

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
            try:
                creds.refresh(Request())
            except Exception as e:
                # invalid_grant: refresh token vencido o revocado (apps OAuth en
                # modo "Testing" caducan el grant a los 7 días; también pasa al
                # re-consentir el mismo client para otro scope). Antes esto
                # reventaba TODA la sincronización con un traceback — ahora se
                # relanza el flujo OAuth en el navegador para re-autorizar.
                print(f"  [aviso] Token de Drive vencido/revocado ({e}) — re-autorizando en el navegador...")
                creds = None
        if not creds or not creds.valid:
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


def _subir(service, local_path, folder_id, nuevo_nombre=None, stats=None, md5_local=None):
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

        if md5_local is None:
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

        migrar = fid and a_sheets and not ya_convertido

        if fid and not migrar:
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
            if migrar:
                # Solo se borra el xlsx crudo DESPUÉS de crear el Sheet con éxito,
                # para no dejar una ventana sin ninguna versión del archivo en Drive
                # si el create() fallara.
                service.files().delete(fileId=fid).execute()
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
        ".html": "text/html",
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


NOMINAS_ENVIO = "Nóminas de Envío"


def _depositar_arbol_local(rango_dir):
    """Copia las planillas por establecimiento de un rango out_gt/ al árbol local
    ordenado: 04_Farmacia_Gestion_Territorial/<ESTAB>/Nóminas de Envío/<mes>/<fecha>/,
    donde <fecha> es el día en que se generó el archivo (= día del envío) y <mes>
    es "JULIO 2026" etc. — mismo esquema de 2 niveles que agregar_gt_manual.py
    (antes esta función depositaba directo en Nóminas de Envío/<fecha>/, sin mes,
    y quedaba desordenado frente a lo que sí anida agregar_gt_manual.py; detectado
    30-07-2026, requería correr _reorganizar_nominas_por_mes.py a mano después).
    Dedup por MD5; si existe un archivo distinto con el mismo nombre, se agrega
    el rango como sufijo. Devuelve cuántos archivos copió."""
    import shutil
    if not os.path.isdir(GT_SOLICITUDES_DIR):
        return 0
    local_de_drive = {v: k for k, v in _SOLICITUDES_A_DRIVE.items()}
    depositados = 0
    for f in sorted(glob.glob(os.path.join(rango_dir, "*.xlsx")) +
                    glob.glob(os.path.join(rango_dir, "*.pdf"))):
        nb = os.path.basename(f)
        if nb.startswith("~$"):
            continue
        destino = _destino_de_archivo(nb)
        carpeta_local = local_de_drive.get(destino) if destino else None
        if not carpeta_local:
            continue
        f_mtime = datetime.fromtimestamp(os.path.getmtime(f))
        fecha = f_mtime.strftime("%d-%m-%Y")
        carpeta_mes = f"{_GM.MESES_ES[f_mtime.month - 1]} {f_mtime.year}"
        ddir = os.path.join(GT_SOLICITUDES_DIR, carpeta_local, NOMINAS_ENVIO, carpeta_mes, fecha)
        os.makedirs(ddir, exist_ok=True)
        dest = os.path.join(ddir, nb)
        if os.path.exists(dest):
            if _md5(dest) == _md5(f):
                continue
            base, ext = os.path.splitext(nb)
            dest = os.path.join(ddir, f"{base} (rango {os.path.basename(rango_dir)}){ext}")
            if os.path.exists(dest) and _md5(dest) == _md5(f):
                continue
        shutil.copy2(f, dest)
        depositados += 1
    return depositados


_GT_SYNC_STATE_FILE = os.path.join(WORK_DIR, "_gt_historial_sync.json")


def _cargar_gt_sincronizados():
    """Nombres de carpetas out_gt/<rango> ya depositadas en el árbol local
    ordenado en una corrida anterior. Un rango es inmutable una vez que su
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

    fid_gt = _obtener_o_crear_carpeta(service, SUB_GT, raiz_id, cache)

    # Desde el 21-07-2026 Drive ya NO tiene carpeta Historial ni Acumulados (el
    # usuario las disolvió): las planillas de cada rango out_gt/ se depositan en
    # el árbol local ordenado (<ESTAB>/Nóminas de Envío/<fecha>) y llegan a
    # Drive vía sync_gt_solicitudes. Los Cruce_GT_Clasificacion quedan solo en
    # out_gt/ local. _gt_historial_sync.json marca los rangos ya depositados
    # (los rangos anteriores al último son inmutables — no se vuelven a tocar).
    sincronizados = _cargar_gt_sincronizados()
    nuevos_sincronizados = set(sincronizados)
    anteriores, ultimo = rangos[:-1], rangos[-1]
    saltados = 0
    for d in anteriores:
        nombre_rango = os.path.basename(d)
        if nombre_rango in sincronizados:
            saltados += 1
            continue
        _depositar_arbol_local(d)
        nuevos_sincronizados.add(nombre_rango)
    if saltados:
        print(f"  [GT] {saltados} rango(s) histórico(s) ya depositados — omitidos")

    nombre_ultimo = os.path.basename(ultimo)
    depositados = _depositar_arbol_local(ultimo)
    if depositados:
        print(f"  [GT] {depositados} planilla(s) del rango «{nombre_ultimo}» "
              f"depositadas en el árbol local por establecimiento")

    _guardar_gt_sincronizados(nuevos_sincronizados)

    # (El antiguo paso "frente" — últimas planillas sueltas bajo
    #  <ESTAB>/Nóminas de Envío/ — se eliminó el 21-07-2026: dejaba archivos
    #  sueltos que duplicaban la subcarpeta de fecha del mismo envío. El envío
    #  vigente ahora se encuentra en la subcarpeta de fecha más reciente.)
    estabs = sorted({_destino_de_archivo(os.path.basename(f)) or ""
                     for f in glob.glob(os.path.join(ultimo, "*"))} - {""})
    estabs_str = ", ".join(estabs) if estabs else "(ninguno)"
    print(f"  [GT] {len(rangos)} rango(s) · último «{nombre_ultimo}» · "
          f"establecimientos: {estabs_str}")

    sync_gt_solicitudes(service, fid_gt, stats, cache)


# Carpeta hermana del repo (ver CLAUDE.md) donde viven las solicitudes GT ya
# analizadas: PDFs de recetas individuales/combinadas + feedback, organizados
# por establecimiento (subcarpetas creadas 19-07-2026).
GT_SOLICITUDES_DIR = os.path.join(os.path.dirname(WORK_DIR), "04_Farmacia_Gestion_Territorial")

# Mapea el nombre de la subcarpeta LOCAL (convención "TIPO_LUGAR" de
# gt_maestro, ej. HOSPITAL_TOLTEN) al nombre de carpeta que YA existe en
# Drive bajo "2 - Gestion Territorial" (convención "LUGAR TIPO" heredada de
# SSASUR/skill_gt, ej. "TOLTEN HOSP") — para reusar la carpeta existente en
# vez de crear una duplicada con un nombre ligeramente distinto.
_SOLICITUDES_A_DRIVE = {
    "CESFAM_FREIRE": "CESFAM FREIRE",
    "CESFAM_HUALPIN": "CESFAM HUALPIN",
    "CESFAM_QUEPE": "CESFAM QUEPE",
    "CESFAM_TEODORO_SCHMIDT": "CESFAM TEODORO SCHMIDT",
    "DSM_GORBEA": "GORBEA DSM",
    "DSM_LONCOCHE": "LONCOCHE DSM",
    "DSM_TOLTEN": "TOLTEN DSM",
    "HOSPITAL_GORBEA": "GORBEA HOSP",
    "HOSPITAL_LONCOCHE": "LONCOCHE HOSP",
    "HOSPITAL_TOLTEN": "TOLTEN HOSP",
    "PSR_COMUY": "PSR COMUY",
    "PSR_QUEULE": "PSR QUEULE",
    "PSR_LOS_GALPONES": "PSR LOS GALPONES",
}


# Estado local del espejo del árbol GT (relpath → md5 ya subido). Sin esto,
# los .xlsx del árbol (convertidos a Sheets nativo, sin md5Checksum en Drive)
# se re-subirían TODOS en cada corrida — mismo bug que ya se arregló para
# Historial con _gt_historial_sync.json. El árbol ordenado por establecimiento/
# fecha (21-07-2026) tiene ~178 archivos y crece, así que acá el dedup se hace
# contra este JSON local en vez de contra Drive.
_GT_ARBOL_STATE_FILE = os.path.join(WORK_DIR, "_gt_arbol_sync.json")


def _cargar_arbol_sincronizado():
    if os.path.exists(_GT_ARBOL_STATE_FILE):
        try:
            with open(_GT_ARBOL_STATE_FILE, encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def sync_gt_solicitudes(service, fid_gt, stats, cache):
    """Refleja en Drive las dos carpetas locales por establecimiento
    (04_Farmacia_Gestion_Territorial/<carpeta local>/):
      · "Revisión de Solicitudes" (el PDF combinado con TODAS las recetas
        solicitadas + feedback; los PDF individuales por paciente quedan
        aparte en la subcarpeta "PDFs individuales" para no tapar el
        combinado)
      · "Nóminas de Envío" (una subcarpeta DD-MM-YYYY por cada envío hecho;
        historial completo ordenado el 21-07-2026)
    Reusa el mismo nombre de carpeta que ya existe en Drive (ver
    _SOLICITUDES_A_DRIVE) en vez de crear una carpeta duplicada."""
    if not os.path.isdir(GT_SOLICITUDES_DIR):
        return
    estado = _cargar_arbol_sincronizado()
    subidos = 0
    for carpeta_local, estab_drive in _SOLICITUDES_A_DRIVE.items():
        local_dir = os.path.join(GT_SOLICITUDES_DIR, carpeta_local)
        if not os.path.isdir(local_dir):
            continue
        fid_estab = _obtener_o_crear_carpeta(service, estab_drive, fid_gt, cache)
        for nombre_carpeta in ("Revisión de Solicitudes", NOMINAS_ENVIO):
            local_sub = os.path.join(local_dir, nombre_carpeta)
            if not os.path.isdir(local_sub):
                continue
            fid_sub = _obtener_o_crear_carpeta(service, nombre_carpeta, fid_estab, cache)
            subidos += _subir_arbol(service, local_sub, fid_sub, stats, cache,
                                    prefijo_log=f"{estab_drive}/{nombre_carpeta}",
                                    estado=estado)
    with open(_GT_ARBOL_STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(estado, f, ensure_ascii=False, indent=2, sort_keys=True)
    if subidos == 0:
        print("  [GT] Revisión de Solicitudes / Nóminas de Envío: nada nuevo que subir")


def _subir_arbol(service, local_dir, fid_padre, stats, cache, prefijo_log, estado=None):
    """Sube todos los .pdf/.xlsx de local_dir a fid_padre en Drive, y baja
    recursivamente por sus subcarpetas preservando la estructura. Devuelve
    cuántos archivos subió/actualizó (no cuenta los "skip").

    `estado` (relpath→md5) evita hasta la llamada a la API para archivos ya
    subidos sin cambios; si el md5 local coincide con el registrado, ni se
    lista ni se sube."""
    subidos = 0
    for nombre in sorted(os.listdir(local_dir)):
        ruta = os.path.join(local_dir, nombre)
        if os.path.isdir(ruta):
            fid_sub = _obtener_o_crear_carpeta(service, nombre, fid_padre, cache)
            subidos += _subir_arbol(service, ruta, fid_sub, stats, cache,
                                    f"{prefijo_log}/{nombre}", estado=estado)
            continue
        if nombre.startswith("~$") or os.path.splitext(nombre)[1].lower() not in (".pdf", ".xlsx"):
            continue
        clave = f"{prefijo_log}/{nombre}"
        md5_local = _md5(ruta)
        if estado is not None and estado.get(clave) == md5_local:
            stats["skip"] += 1
            continue
        r = _subir(service, ruta, fid_padre, stats=stats, md5_local=md5_local)
        if r != "fail" and estado is not None:
            estado[clave] = md5_local
        if r != "skip":
            print(f"  {prefijo_log}/{nombre}: {r}")
            subidos += 1
    return subidos


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

    src = _mas_reciente(os.path.join(WORK_DIR, "Pedido_Fusion_Simple_AA*.xlsx"))
    if src:
        r = _subir(service, src, fid, nuevo_nombre="Pedido_Fusion_Simple_AA.xlsx", stats=stats)
        print(f"  Pedido_Fusion_Simple_AA.xlsx: {r}")
    else:
        print("  Pedido_Fusion_Simple_AA.xlsx: no encontrado, omitido")


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
    # Si existe una planilla de ciclo NUEVO (más reciente que el último Resumen),
    # subirla aunque haya Resumen — el nuevo ciclo aún no tiene conteo aplicado.
    # Si el Resumen es el más reciente (o no hay planilla nueva), sube el Resumen.
    resumen = _mas_reciente(os.path.join(WORK_DIR, "Programacion_AA", "Resumen_Programacion_AA*.xlsx"))
    planilla = _mas_reciente(os.path.join(WORK_DIR, "Programacion_AA", "Programacion_AA_*.xlsx"))
    ciclo_nuevo = (
        planilla and resumen and
        os.path.getmtime(planilla) > os.path.getmtime(resumen)
    )
    if resumen and not ciclo_nuevo:
        r = _subir(service, resumen, fid, nuevo_nombre="Resumen_Programacion_AA.xlsx", stats=stats)
        print(f"  Resumen_Programacion_AA.xlsx: {r}")
        return
    if planilla:
        r = _subir(service, planilla, fid, nuevo_nombre="Programacion_AA.xlsx", stats=stats)
        etiqueta = "(nuevo ciclo — conteo pendiente)" if ciclo_nuevo else "(planilla del ciclo, sin conteo aplicado todavía)"
        print(f"  Programacion_AA.xlsx: {r}  {etiqueta}")
    else:
        print("  Programacion_AA: no encontrado, omitido (corre programacion_aa.py)")


def sync_servicios(service, raiz_id, stats, cache=None):
    """Sube Servicios_Farmaceuticos/<MES AÑO>/Resumen_Servicios_Farmaceuticos_*.xlsx
    — el agregado QF × actividad (sin RUT). Una subcarpeta por mes, mismo
    patrón que sync_centinela (una subcarpeta por semana)."""
    fid = _obtener_o_crear_carpeta(service, SUB_SERVICIOS, raiz_id, cache)
    src_base = os.path.join(WORK_DIR, "Servicios_Farmaceuticos")
    if not os.path.isdir(src_base):
        print("  [Servicios Farmaceuticos] sin Servicios_Farmaceuticos/ todavía")
        return
    meses = sorted(
        d for d in glob.glob(os.path.join(src_base, "*")) if os.path.isdir(d)
    )
    if not meses:
        print("  [Servicios Farmaceuticos] sin meses en Servicios_Farmaceuticos/")
        return
    for d in meses:
        nombre_mes = os.path.basename(d)
        fid_mes = _obtener_o_crear_carpeta(service, nombre_mes, fid, cache)
        for f in sorted(glob.glob(os.path.join(d, "Resumen_Servicios_Farmaceuticos_*.xlsx"))):
            r = _subir(service, f, fid_mes, stats=stats)
            if r != "skip":
                print(f"  {nombre_mes}/{os.path.basename(f)}: {r}")
    print(f"  [Servicios Farmaceuticos] {len(meses)} mes(es): "
          f"{', '.join(os.path.basename(d) for d in meses)}")


def sync_centinela_sm(service, raiz_id, stats, cache=None):
    """Sube Centinela_Inyectables_SM/<fecha>/reporte_inyectables_sm_<fecha>.html
    — stock de antipsicóticos de depósito (salud mental ambulatoria), SIN RUT
    ni nombre de paciente (solo cifras agregadas de stock). Una subcarpeta por
    fecha en que el reporte se regeneró, mismo patrón que sync_centinela()
    (semanas) / sync_servicios() (meses)."""
    fid = _obtener_o_crear_carpeta(service, SUB_CENTINELA_SM, raiz_id, cache)
    src_base = os.path.join(WORK_DIR, "Centinela_Inyectables_SM")
    if not os.path.isdir(src_base):
        print("  [Centinela Inyectables SM] sin Centinela_Inyectables_SM/ todavía")
        return
    fechas = sorted(
        d for d in glob.glob(os.path.join(src_base, "*")) if os.path.isdir(d)
    )
    if not fechas:
        print("  [Centinela Inyectables SM] sin fechas en Centinela_Inyectables_SM/")
        return
    for d in fechas:
        nombre_fecha = os.path.basename(d)
        fid_fecha = _obtener_o_crear_carpeta(service, nombre_fecha, fid, cache)
        for f in sorted(glob.glob(os.path.join(d, "*.html"))):
            r = _subir(service, f, fid_fecha, stats=stats)
            if r != "skip":
                print(f"  {nombre_fecha}/{os.path.basename(f)}: {r}")
    print(f"  [Centinela Inyectables SM] {len(fechas)} fecha(s): "
          f"{', '.join(os.path.basename(d) for d in fechas)}")


def sync_clozapina(service, stats, cache=None):
    """Sube los reportes mensuales de Clozapina (Clozapina_Reportes/*.xlsx) a su
    PROPIA carpeta de nivel superior en Drive (NOMBRE_RAIZ_CLOZAPINA) — no
    dentro de 'Farmacia AA'. Solo corre con --solo-clozapina (ver main())."""
    src_dir = os.path.join(WORK_DIR, "Clozapina_Reportes")
    if not os.path.isdir(src_dir):
        print("  [Clozapina] sin Clozapina_Reportes/ todavía — corre AUTO_SSASUR.py --clozapina primero")
        return
    fid = _obtener_o_crear_carpeta(service, NOMBRE_RAIZ_CLOZAPINA, parent_id=None, cache=cache)
    archivos = sorted(glob.glob(os.path.join(src_dir, "Consolidado_Hemogramas_Clozapina_*.xlsx")))
    if not archivos:
        print("  [Clozapina] sin reportes mensuales que subir")
        return
    for f in archivos:
        r = _subir(service, f, fid, stats=stats)
        print(f"  {os.path.basename(f)}: {r}")


def sync_gt_confidencial(service, stats, cache=None):
    """Sube GT_Confidencial/ (copias SELLADAS de recetas GT — con ANULAR/
    pendientes/ya enviada/observación escritos encima por
    descargar_recetas_pdf.py) a su PROPIA carpeta de nivel superior en Drive
    (NOMBRE_RAIZ_GT_CONFIDENCIAL) — no dentro de 'Farmacia AA'. Preserva la
    estructura <carpeta_local>/<fecha>/ vía _subir_arbol. Solo corre con
    --solo-gt-confidencial (ver main()); la carpeta normal de GT
    (--solo-gt) sube las recetas LIMPIAS, sin escritura."""
    src_dir = os.path.join(WORK_DIR, "GT_Confidencial")
    if not os.path.isdir(src_dir):
        print("  [GT Confidencial] sin GT_Confidencial/ todavía — nada que subir")
        return
    fid = _obtener_o_crear_carpeta(service, NOMBRE_RAIZ_GT_CONFIDENCIAL, parent_id=None, cache=cache)
    subidos = _subir_arbol(service, src_dir, fid, stats, cache, NOMBRE_RAIZ_GT_CONFIDENCIAL)
    if subidos == 0:
        print("  [GT Confidencial] nada nuevo que subir")


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
    ap.add_argument("--solo-servicios", action="store_true")
    ap.add_argument("--solo-centinela-sm", action="store_true")
    ap.add_argument("--solo-clozapina", action="store_true",
                    help="Sube Clozapina_Reportes/ a su PROPIA carpeta Drive (privada, fuera de "
                         "'Farmacia AA') — NUNCA corre como parte de una sincronización general, "
                         "solo con este flag explícito.")
    ap.add_argument("--solo-gt-confidencial", action="store_true",
                    help="Sube GT_Confidencial/ (copias SELLADAS de recetas GT) a su PROPIA carpeta "
                         "Drive privada, fuera de 'Farmacia AA' — NUNCA corre como parte de una "
                         "sincronización general ni junto con --solo-gt (que sube las recetas "
                         "limpias, sin escritura, a la carpeta normal).")
    a = ap.parse_args()

    if a.setup and os.path.exists(TOKEN_FILE):
        os.remove(TOKEN_FILE)
        print("[SETUP] Token eliminado — se iniciará flujo OAuth desde cero.")

    print("\n[Drive] Conectando con Google Drive...")
    svc = _get_service()

    cache = {}
    known = {}
    if os.path.exists(_FOLDER_CACHE_FILE):
        try:
            with open(_FOLDER_CACHE_FILE, encoding="utf-8") as _f:
                known = json.load(_f)
        except Exception:
            pass

    # Caché de Clozapina APARTE (gitignored) — nunca se mezcla con known/_FOLDER_CACHE_FILE.
    known_clozapina = {}
    if os.path.exists(_FOLDER_CACHE_FILE_CLOZAPINA):
        try:
            with open(_FOLDER_CACHE_FILE_CLOZAPINA, encoding="utf-8") as _f:
                known_clozapina = json.load(_f)
        except Exception:
            pass
    if NOMBRE_RAIZ_CLOZAPINA in known_clozapina:
        cache[(NOMBRE_RAIZ_CLOZAPINA, None)] = known_clozapina[NOMBRE_RAIZ_CLOZAPINA]

    # Caché de GT Confidencial APARTE (gitignored) — mismo motivo que Clozapina.
    known_gt_confidencial = {}
    if os.path.exists(_FOLDER_CACHE_FILE_GT_CONFIDENCIAL):
        try:
            with open(_FOLDER_CACHE_FILE_GT_CONFIDENCIAL, encoding="utf-8") as _f:
                known_gt_confidencial = json.load(_f)
        except Exception:
            pass
    if NOMBRE_RAIZ_GT_CONFIDENCIAL in known_gt_confidencial:
        cache[(NOMBRE_RAIZ_GT_CONFIDENCIAL, None)] = known_gt_confidencial[NOMBRE_RAIZ_GT_CONFIDENCIAL]

    stats = {"ok": 0, "skip": 0, "fail": 0}

    todos = not any([a.solo_app, a.solo_gt, a.solo_pedido, a.solo_auditoria,
                     a.solo_centinela, a.solo_programacion, a.solo_servicios,
                     a.solo_centinela_sm, a.solo_clozapina, a.solo_gt_confidencial])
    # Clozapina y GT Confidencial viven en su PROPIA carpeta (fuera de
    # "Farmacia AA") — si son lo ÚNICO que se pidió subir, no hace falta
    # resolver/imprimir la raíz "Farmacia AA" en absoluto (evita el mensaje
    # engañoso "Subiendo a «Farmacia AA»" cuando en realidad va a una
    # carpeta totalmente distinta).
    necesita_raiz_farmacia = todos or any([a.solo_app, a.solo_gt, a.solo_pedido,
                                            a.solo_auditoria, a.solo_centinela,
                                            a.solo_programacion, a.solo_servicios,
                                            a.solo_centinela_sm])

    raiz_id = None
    if necesita_raiz_farmacia:
        # Carpeta raíz "Farmacia AA" — usa IDs pre-conocidos si existen
        if NOMBRE_RAIZ in known:
            raiz_id = known[NOMBRE_RAIZ]
            # Pre-carga sub-carpetas fijas para evitar búsquedas API
            for sub in (SUB_APP, SUB_GT, SUB_PEDIDO, SUB_AUDIT, SUB_CENTINELA, SUB_PROG, SUB_SERVICIOS, SUB_CENTINELA_SM):
                if sub in known:
                    cache[(sub, raiz_id)] = known[sub]
        else:
            raiz_id = _obtener_o_crear_carpeta(svc, NOMBRE_RAIZ, cache=cache)

        print(f"\n[Drive] Subiendo a «{NOMBRE_RAIZ}» (id={raiz_id[:8]}…)")
    ts = datetime.now().strftime("%H:%M:%S")
    print(f"        {ts}\n")

    # Cada sync_*() va en su propio try/except: un error de carpeta/API en una
    # sección (p.ej. resolver una de las 12 carpetas de GT) no debe abortar el
    # resto de secciones independientes de la corrida.
    if todos or a.solo_app:
        print(f"  ── {SUB_APP} ──")
        try:
            sync_app(svc, raiz_id, stats, cache)
        except Exception as e:
            print(f"  [error] {SUB_APP}: {e}")
            stats["fail"] += 1

    if todos or a.solo_gt:
        print(f"\n  ── {SUB_GT} ──")
        try:
            sync_gt(svc, raiz_id, stats, cache)
        except Exception as e:
            print(f"  [error] {SUB_GT}: {e}")
            stats["fail"] += 1

    if todos or a.solo_pedido:
        print(f"\n  ── {SUB_PEDIDO} ──")
        try:
            sync_pedido(svc, raiz_id, stats, cache)
        except Exception as e:
            print(f"  [error] {SUB_PEDIDO}: {e}")
            stats["fail"] += 1

    if todos or a.solo_auditoria:
        print(f"\n  ── {SUB_AUDIT} ──")
        try:
            sync_auditoria(svc, raiz_id, stats, cache)
        except Exception as e:
            print(f"  [error] {SUB_AUDIT}: {e}")
            stats["fail"] += 1

    if todos or a.solo_centinela:
        print(f"\n  ── {SUB_CENTINELA} ──")
        try:
            sync_centinela(svc, raiz_id, stats, cache)
        except Exception as e:
            print(f"  [error] {SUB_CENTINELA}: {e}")
            stats["fail"] += 1

    if todos or a.solo_programacion:
        print(f"\n  ── {SUB_PROG} ──")
        try:
            sync_programacion(svc, raiz_id, stats, cache)
        except Exception as e:
            print(f"  [error] {SUB_PROG}: {e}")
            stats["fail"] += 1

    if todos or a.solo_servicios:
        print(f"\n  ── {SUB_SERVICIOS} ──")
        try:
            sync_servicios(svc, raiz_id, stats, cache)
        except Exception as e:
            print(f"  [error] {SUB_SERVICIOS}: {e}")
            stats["fail"] += 1

    if todos or a.solo_centinela_sm:
        print(f"\n  ── {SUB_CENTINELA_SM} ──")
        try:
            sync_centinela_sm(svc, raiz_id, stats, cache)
        except Exception as e:
            print(f"  [error] {SUB_CENTINELA_SM}: {e}")
            stats["fail"] += 1

    # Clozapina: EXPLÍCITAMENTE fuera de "todos" — solo sube con --solo-clozapina.
    if a.solo_clozapina:
        print(f"\n  ── {NOMBRE_RAIZ_CLOZAPINA} (carpeta separada, privada) ──")
        try:
            sync_clozapina(svc, stats, cache)
        except Exception as e:
            print(f"  [error] {NOMBRE_RAIZ_CLOZAPINA}: {e}")
            stats["fail"] += 1

    # GT Confidencial: EXPLÍCITAMENTE fuera de "todos" — solo sube con --solo-gt-confidencial.
    if a.solo_gt_confidencial:
        print(f"\n  ── {NOMBRE_RAIZ_GT_CONFIDENCIAL} (carpeta separada, privada) ──")
        try:
            sync_gt_confidencial(svc, stats, cache)
        except Exception as e:
            print(f"  [error] {NOMBRE_RAIZ_GT_CONFIDENCIAL}: {e}")
            stats["fail"] += 1

    total = stats["ok"] + stats["skip"] + stats["fail"]
    print(f"\n[Drive] Listo: {stats['ok']} subidos · {stats['skip']} sin cambios · "
          f"{stats['fail']} errores  ({total} total)")

    # Persiste el ID de la carpeta de Clozapina en su archivo APARTE (gitignored)
    # — NUNCA en _FOLDER_CACHE_FILE, que sí está versionado en el repo público.
    if (NOMBRE_RAIZ_CLOZAPINA, None) in cache:
        fid_cloz = cache[(NOMBRE_RAIZ_CLOZAPINA, None)]
        if known_clozapina.get(NOMBRE_RAIZ_CLOZAPINA) != fid_cloz:
            with open(_FOLDER_CACHE_FILE_CLOZAPINA, "w", encoding="utf-8") as _f:
                json.dump({NOMBRE_RAIZ_CLOZAPINA: fid_cloz}, _f, ensure_ascii=False, indent=2)

    # Mismo motivo, para GT Confidencial.
    if (NOMBRE_RAIZ_GT_CONFIDENCIAL, None) in cache:
        fid_gtc = cache[(NOMBRE_RAIZ_GT_CONFIDENCIAL, None)]
        if known_gt_confidencial.get(NOMBRE_RAIZ_GT_CONFIDENCIAL) != fid_gtc:
            with open(_FOLDER_CACHE_FILE_GT_CONFIDENCIAL, "w", encoding="utf-8") as _f:
                json.dump({NOMBRE_RAIZ_GT_CONFIDENCIAL: fid_gtc}, _f, ensure_ascii=False, indent=2)

    # Persiste IDs de carpetas de "Farmacia AA" nuevas descubiertas en esta
    # corrida (solo si de verdad se resolvió esa raíz — ver necesita_raiz_farmacia).
    if not necesita_raiz_farmacia:
        if stats["fail"]:
            sys.exit(1)
        return
    nuevos = {nombre: fid for (nombre, pid), fid in cache.items() if pid == raiz_id}
    nuevos[NOMBRE_RAIZ] = raiz_id
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
