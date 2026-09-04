#!/usr/bin/env python3
# -*- coding: utf-8 -*-
r"""
publicar_escritorio.py — Centraliza en el ESCRITORIO las salidas de todos los
procesos de la Farmacia AT Abierta, para revisar los resultados de un vistazo
sin entrar a la carpeta del repositorio.

  Escritorio\Farmacia AA\  (recortada 04-09-2026 a solo 6 categorías — ver
  memoria del proyecto "drive-carpetas-recorte-6-categorias")
    ├── AUTO_SSASUR.lnk                   (acceso directo → AUTO_SSASUR.bat)
    ├── Gestion Territorial.lnk           (acceso directo → GT.bat)
    ├── Recetas Cheque ISP.lnk            (acceso directo → RECETAS_CHEQUE.bat)
    ├── LEEME.txt
    ├── _ultima_sync.txt                  (qué se copió y cuándo)
    ├── 2 - Gestion Territorial\  ÚLTIMO rango al frente + Historial\<rango>\
    ├── 3 - Recetas Cheque\     SOLO un acceso directo a la carpeta local (datos de
    │                           pacientes NO se copian a la nube de OneDrive)
    ├── 5 - Pedido Fusionado\   Pedido_Fusion_AA.xlsx (Farm_Bod + Bod_Farmacos + Dialisis)
    ├── 6 - Centinela\          Centinela_Reportes\<Sxx>\ (json + pdf) por semana
    ├── 9 - Clozapina\          Accesos a carpetas locales de reportes/hemogramas
    │                             (RUT pacientes, NO se copian a la nube)
    ├── 10 - Servicios Farmaceuticos\  Servicios_Farmaceuticos\<MES AÑO>\ (agregado
    │                             QF x actividad, SIN RUT — sí se copia a la nube)
    └── 11 - Centinela Inyectables SM\  Centinela_Inyectables_SM\<fecha>\ (stock
                                    de antipsicóticos de depósito, SIN RUT — sí se copia)

IMPORTANTE: este script COPIA, no mueve. El repositorio sigue siendo la fuente de
verdad — la app Streamlit lee el Consolidado del repo y PUBLICAR_DATOS.bat publica
desde el repo. Aquí solo dejamos copias legibles, ordenadas por proceso.

Uso:
    py publicar_escritorio.py            # sincroniza TODO
    py publicar_escritorio.py --gt           # solo Gestion Territorial (out_gt)
    py publicar_escritorio.py --rch          # solo el acceso directo de recetas cheque
    py publicar_escritorio.py --pedido       # solo Pedido_Fusion_AA.xlsx
    py publicar_escritorio.py --centinela    # solo Centinela_Reportes\
    py publicar_escritorio.py --servicios    # solo Servicios_Farmaceuticos\ (recuento QF)
    py publicar_escritorio.py --centinela-sm # solo Centinela_Inyectables_SM\ (antipsicóticos depósito)
    py publicar_escritorio.py --enlaces      # solo (re)crea carpetas, LEEME y accesos
"""
import os
import re
import sys
import glob
import shutil
import subprocess
from datetime import datetime

# Nombre de carpeta de un rango GT: DD-MM-AAAA_DD-MM-AAAA
_RANGO_RE = re.compile(r"^\d{2}-\d{2}-\d{4}_\d{2}-\d{2}-\d{4}$")

# Detecta el establecimiento de destino en nombres de archivo GT
# "Inyectables_Planilla" (antipsicóticos LAI) y "SaludMental_Planilla"
# (patrón legado) van ANTES que "Planilla" en la alternativa por la misma
# razón que en publicar_drive.py: sin ellas, el archivo igual matcheaba pero
# con la palabra de más pegada al destino (no existe en el mapeo) y se
# perdía en silencio — ver fix 07-09-2026.
_TIPO_GT_RE = re.compile(
    r"^(.+?)_(Inyectables_Planilla|SaludMental_Planilla|Planilla|Letrero|Controlados_Planilla|Verificacion)\.(xlsx|pdf)$",
    re.IGNORECASE,
)

def _destino_de_archivo(nombre):
    m = _TIPO_GT_RE.match(nombre)
    return m.group(1).replace("_", " ") if m else None

# ── Rutas base ───────────────────────────────────────────────────────────────
WORK_DIR = os.path.dirname(os.path.abspath(__file__))

# Carpeta del formulario ISP: fuente única = utils_aa.py (configurable por
# variable de entorno MAESTRO_RCH_DIR — ver utils_aa.py).
sys.path.insert(0, WORK_DIR)
from utils_aa import RCH_DIR, setup_stdout
setup_stdout()  # evita UnicodeEncodeError en consolas cp1252 (mensajes usan →, tildes)
PREFIJO_FORM = "Formulario-Notificacion-Recetas-Cheque"

NOMBRE_CARPETA = "Farmacia AA"
SUB_GT    = "2 - Gestion Territorial"
SUB_RCH    = "3 - Recetas Cheque"
SUB_PEDIDO = "5 - Pedido Fusionado"
SUB_CENTINELA = "6 - Centinela"
SUB_CLOZAPINA = "9 - Clozapina"
SUB_SERVICIOS = "10 - Servicios Farmaceuticos"
SUB_CENTINELA_SM = "11 - Centinela Inyectables SM"

# Iconos para distinguir los accesos directos (shell32.dll, índices clásicos).
_SHELL32 = os.path.join(os.environ.get("SystemRoot", r"C:\Windows"), "System32", "shell32.dll")
def _icon(idx):
    return f"{_SHELL32},{idx}"
_ICON_REFRESH = _icon(238)  # flechas sync  → actualizar/publicar
_ICON_DOC     = _icon(1)    # documento     → planillas GT
_ICON_RUN     = _icon(2)    # aplicación    → proceso recetas cheque
_ICON_FOLDER  = _icon(4)    # carpeta       → abrir carpeta local


def detectar_escritorio():
    """Ruta del Escritorio probando las ubicaciones habituales (OneDrive lo
    redirige en este equipo)."""
    perfil = os.environ.get("USERPROFILE", os.path.expanduser("~"))
    candidatos = []
    od = os.environ.get("OneDrive")
    if od:
        candidatos += [os.path.join(od, "Desktop"), os.path.join(od, "Escritorio")]
    candidatos += [
        os.path.join(perfil, "OneDrive", "Desktop"),
        os.path.join(perfil, "OneDrive", "Escritorio"),
        os.path.join(perfil, "Desktop"),
        os.path.join(perfil, "Escritorio"),
    ]
    for c in candidatos:
        if os.path.isdir(c):
            return c
    destino = os.path.join(perfil, "Desktop")
    os.makedirs(destino, exist_ok=True)
    return destino


DESKTOP = detectar_escritorio()
BASE    = os.path.join(DESKTOP, NOMBRE_CARPETA)


# ── Reporte / log ────────────────────────────────────────────────────────────
class _Reporte:
    def __init__(self):
        self.ok = self.skip = self.fail = 0
        self.log = []

    def say(self, msg):
        print(msg)
        self.log.append(msg)


REP = _Reporte()


# ── Utilidades de copia (incremental, tolerante a archivos abiertos) ──────────
def _es_temporal(nombre):
    return nombre.startswith("~$") or nombre.endswith(".tmp")


def _igual(src, dst):
    """True si dst ya existe y coincide en tamaño y fecha (±2s) → no recopiar."""
    try:
        a, b = os.stat(src), os.stat(dst)
    except OSError:
        return False
    return a.st_size == b.st_size and abs(a.st_mtime - b.st_mtime) <= 2


def _copiar(src, dst_dir, *, nuevo_nombre=None):
    """Copia src a dst_dir conservando metadatos. Salta si ya está igual.
    No aborta si el archivo está abierto en Excel o bloqueado por OneDrive."""
    if not src or not os.path.isfile(src):
        return "none"
    os.makedirs(dst_dir, exist_ok=True)
    destino = os.path.join(dst_dir, nuevo_nombre or os.path.basename(src))
    if _igual(src, destino):
        REP.skip += 1
        return "skip"
    try:
        shutil.copy2(src, destino)
        REP.ok += 1
        return "ok"
    except PermissionError:
        REP.fail += 1
        REP.say(f"  [aviso] abierto/bloqueado, no copiado: {os.path.basename(src)}")
        return "fail"
    except OSError as e:
        REP.fail += 1
        REP.say(f"  [aviso] error copiando {os.path.basename(src)}: {e}")
        return "fail"


def _mas_reciente(patron):
    cand = [f for f in glob.glob(patron) if not _es_temporal(os.path.basename(f))]
    return max(cand, key=os.path.getmtime) if cand else None


def _espejo(src_dir, dst_dir):
    """Copia recursiva incremental de src_dir → dst_dir (ignora temporales)."""
    n = 0
    for raiz, _dirs, files in os.walk(src_dir):
        rel = os.path.relpath(raiz, src_dir)
        destino = dst_dir if rel == "." else os.path.join(dst_dir, rel)
        for f in files:
            if _es_temporal(f):
                continue
            if _copiar(os.path.join(raiz, f), destino) in ("ok", "skip"):
                n += 1
    return n


def _limpiar_archivos_sueltos(dirpath):
    """Borra solo los archivos (no subcarpetas) que cuelgan de dirpath."""
    if not os.path.isdir(dirpath):
        return
    for nombre in os.listdir(dirpath):
        p = os.path.join(dirpath, nombre)
        if os.path.isfile(p):
            try:
                os.remove(p)
            except OSError:
                pass


# ── Sincronizadores por proceso ──────────────────────────────────────────────
def _copiar_rango_por_estab(rango_dir, dst_base):
    """Copia archivos de un rango GT en dst_base/<ESTABLECIMIENTO>/.
    Archivos sin establecimiento conocido van directamente a dst_base."""
    archivos = sorted(
        glob.glob(os.path.join(rango_dir, "*.xlsx")) +
        glob.glob(os.path.join(rango_dir, "*.pdf"))
    )
    estabs = set()
    for f in archivos:
        nb = os.path.basename(f)
        if _es_temporal(nb):
            continue
        destino = _destino_de_archivo(nb)
        if destino:
            estabs.add(destino)
            _copiar(f, os.path.join(dst_base, destino))
        else:
            _copiar(f, dst_base)
    return sorted(estabs)


def sync_gt():
    src = os.path.join(WORK_DIR, "out_gt")
    base_gt = os.path.join(BASE, SUB_GT)
    if not os.path.isdir(src):
        REP.say("[Gestion Territorial] (sin out_gt todavía)")
        return
    rangos = sorted(
        [d for d in glob.glob(os.path.join(src, "*"))
         if os.path.isdir(d) and _RANGO_RE.match(os.path.basename(d))],
        key=os.path.basename,
    )
    if not rangos:
        REP.say("[Gestion Territorial] (sin rangos en out_gt)")
        return

    # 1) Historial: cada rango en su carpeta, organizado por establecimiento
    hist = os.path.join(base_gt, "Historial")
    for d in rangos:
        _copiar_rango_por_estab(d, os.path.join(hist, os.path.basename(d)))

    # 2) Último rango → carpetas por establecimiento en la raíz GT (al frente)
    ultimo = max(rangos, key=os.path.getmtime)
    rango_actual = os.path.basename(ultimo)
    marca = os.path.join(base_gt, "_rango_actual.txt")
    previo = None
    if os.path.isfile(marca):
        try:
            previo = open(marca, encoding="utf-8").read().strip()
        except OSError:
            pass
    if previo != rango_actual:
        # Nuevo rango: limpia archivos sueltos y carpetas de establecimientos viejos
        _limpiar_archivos_sueltos(base_gt)
        for nombre in os.listdir(base_gt):
            p = os.path.join(base_gt, nombre)
            if os.path.isdir(p) and nombre not in ("Historial",) and not _RANGO_RE.match(nombre):
                shutil.rmtree(p, ignore_errors=True)
    estabs = _copiar_rango_por_estab(ultimo, base_gt)
    try:
        with open(marca, "w", encoding="utf-8") as fh:
            fh.write(rango_actual)
    except OSError:
        pass
    estabs_str = ", ".join(estabs) if estabs else "(ninguno)"
    REP.say(f"[Gestion Territorial] último «{rango_actual}» · "
            f"establecimientos: {estabs_str} · {len(rangos)} rango(s) en Historial")


def sync_rch():
    """NO copia el formulario (datos de pacientes) a la nube de OneDrive: deja
    solo un acceso directo a la carpeta local original, y purga cualquier copia
    previa que se hubiera subido."""
    dst = os.path.join(BASE, SUB_RCH)
    os.makedirs(dst, exist_ok=True)
    # Purga: quita formularios que pudieran haberse copiado antes (privacidad).
    purgados = 0
    for viejo in glob.glob(os.path.join(dst, PREFIJO_FORM + "*.xlsx")):
        try:
            os.remove(viejo)
            purgados += 1
        except OSError:
            pass
    if purgados:
        REP.say(f"  [privacidad] quitada(s) {purgados} copia(s) de la nube")
    # Acceso directo a la carpeta local del formulario.
    if os.path.isdir(RCH_DIR):
        _crear_lnk(os.path.join(dst, "Abrir carpeta Recetas Cheque.lnk"), RCH_DIR,
                   "Abre la carpeta LOCAL del formulario ISP (no se sube a la nube)",
                   icono=_ICON_FOLDER)
        REP.say("[Recetas Cheque] acceso directo a la carpeta local (sin subir a la nube)")
    else:
        REP.say(f"[Recetas Cheque] (no existe la carpeta local {RCH_DIR})")
    nota = (
        "RECETAS CHEQUE — datos de pacientes\n"
        "===================================\n\n"
        "El formulario ISP NO se copia a esta carpeta porque contiene datos de\n"
        "pacientes y el Escritorio se sincroniza a la nube de OneDrive.\n\n"
        "Usa el acceso directo «Abrir carpeta Recetas Cheque» para abrir el\n"
        "formulario en su carpeta LOCAL (no sincronizada):\n"
        f"  {RCH_DIR}\n"
    )
    try:
        with open(os.path.join(dst, "LEEME.txt"), "w", encoding="utf-8") as fh:
            fh.write(nota)
    except OSError:
        pass


def sync_clozapina():
    """Mismo criterio que sync_rch(): los reportes mensuales y los PDF de
    hemograma traen RUT + nombre de paciente (Ley 19.628) — NO se copian a la
    carpeta del Escritorio (sincronizada a OneDrive), solo accesos directos
    locales. Un reporte por mes en Clozapina_Reportes/ (julio, agosto, ...) —
    no se acumula todo en un solo archivo."""
    dst = os.path.join(BASE, SUB_CLOZAPINA)
    os.makedirs(dst, exist_ok=True)
    reportes_dir = os.path.join(WORK_DIR, "Clozapina_Reportes")
    hemogramas_dir = os.path.join(WORK_DIR, "_hemogramas_clozapina")
    # Purga cualquier copia de un reporte que se hubiera dejado antes (privacidad).
    purgados = 0
    for viejo in glob.glob(os.path.join(dst, "Consolidado_Hemogramas_Clozapina*.xlsx")):
        try:
            os.remove(viejo)
            purgados += 1
        except OSError:
            pass
    if purgados:
        REP.say(f"  [privacidad] quitada(s) {purgados} copia(s) de la nube")
    if os.path.isdir(reportes_dir) and glob.glob(os.path.join(reportes_dir, "*.xlsx")):
        _crear_lnk(os.path.join(dst, "Abrir Reportes Clozapina.lnk"), reportes_dir,
                   "Abre la carpeta LOCAL con un reporte por mes (no se sube a la nube)", icono=_ICON_FOLDER)
        REP.say("[Clozapina] acceso directo a la carpeta de reportes mensuales (sin subir a la nube)")
    else:
        REP.say("[Clozapina] (aún no generado — corre AUTO_SSASUR.py --clozapina o clozapina_hce_hemogramas.py)")
    if os.path.isdir(hemogramas_dir):
        _crear_lnk(os.path.join(dst, "Abrir carpeta Hemogramas PDF.lnk"), hemogramas_dir,
                   "Abre la carpeta LOCAL de PDF de hemogramas (no se sube a la nube)", icono=_ICON_FOLDER)
    nota = (
        "CLOZAPINA — datos de pacientes\n"
        "===============================\n\n"
        "Los reportes mensuales y los PDF de hemograma NO se copian a esta carpeta\n"
        "porque contienen RUT/nombre de paciente y el Escritorio se sincroniza a la\n"
        "nube de OneDrive.\n\n"
        "Usa los accesos directos para abrir la carpeta de reportes (uno por mes:\n"
        "julio, agosto, ...) y la carpeta de PDF en su ubicación LOCAL (no sincronizada):\n"
        f"  {reportes_dir}\n"
        f"  {hemogramas_dir}\n"
    )
    try:
        with open(os.path.join(dst, "LEEME.txt"), "w", encoding="utf-8") as fh:
            fh.write(nota)
    except OSError:
        pass


def sync_pedido():
    dst = os.path.join(BASE, SUB_PEDIDO)
    src = _mas_reciente(os.path.join(WORK_DIR, "Pedido_Fusion_AA*.xlsx"))
    if not src:
        REP.say("[Pedido Fusionado] (aún no generado — corre pedido_fusion.py)")
        return
    _copiar(src, dst, nuevo_nombre="Pedido_Fusion_AA.xlsx")
    REP.say(f"[Pedido Fusionado] {os.path.basename(src)} → «{SUB_PEDIDO}»")

def sync_centinela():
    dst = os.path.join(BASE, SUB_CENTINELA)
    src = os.path.join(WORK_DIR, "Centinela_Reportes")
    if not os.path.isdir(src):
        REP.say("[Centinela] (aún no se ha generado ningún reporte)")
        return
    n = _espejo(src, dst)
    semanas = sorted(d for d in os.listdir(src) if os.path.isdir(os.path.join(src, d)))
    REP.say(f"[Centinela] {n} archivo(s) · semanas: {', '.join(semanas) if semanas else '(ninguna)'} → «{SUB_CENTINELA}»")


def sync_servicios():
    """Espejo de Servicios_Farmaceuticos/ (agregado QF x actividad, SIN RUT de
    paciente — el reporte crudo de Agenda Médica nunca llega hasta acá) — mismo
    patrón que sync_centinela()."""
    dst = os.path.join(BASE, SUB_SERVICIOS)
    src = os.path.join(WORK_DIR, "Servicios_Farmaceuticos")
    if not os.path.isdir(src):
        REP.say("[Servicios Farmaceuticos] (aún no se ha generado ningún resumen)")
        return
    n = _espejo(src, dst)
    meses = sorted(d for d in os.listdir(src) if os.path.isdir(os.path.join(src, d)))
    REP.say(f"[Servicios Farmaceuticos] {n} archivo(s) · meses: "
            f"{', '.join(meses) if meses else '(ninguno)'} → «{SUB_SERVICIOS}»")


def sync_centinela_sm():
    """Espejo de Centinela_Inyectables_SM/ (stock de antipsicóticos de depósito
    de salud mental ambulatoria — SIN RUT ni nombre de paciente, solo cifras
    agregadas de stock) — mismo patrón que sync_centinela()/sync_servicios()."""
    dst = os.path.join(BASE, SUB_CENTINELA_SM)
    src = os.path.join(WORK_DIR, "Centinela_Inyectables_SM")
    if not os.path.isdir(src):
        REP.say("[Centinela Inyectables SM] (aún no se ha generado ningún reporte)")
        return
    n = _espejo(src, dst)
    fechas = sorted(d for d in os.listdir(src) if os.path.isdir(os.path.join(src, d)))
    REP.say(f"[Centinela Inyectables SM] {n} archivo(s) · fechas: "
            f"{', '.join(fechas) if fechas else '(ninguna)'} → «{SUB_CENTINELA_SM}»")


# ── Estructura, LEEME y accesos directos ─────────────────────────────────────
LEEME = """\
========================================================================
  FARMACIA AT ABIERTA — Resultados (Hospital de Pitrufquén)
========================================================================

Esta carpeta reúne, en el Escritorio, las SALIDAS de todos los procesos
para revisarlas sin abrir la carpeta del programa. Se actualizan solas
cada vez que corres cada proceso.

  AUTO_SSASUR                 ← descarga de SSASUR + recalcula + publica (todo)
  Gestion Territorial         ← solo descarga y genera planillas GT
  Recetas Cheque ISP          ← solo actualiza el registro ISP del mes
  Sincronizar Todo            ← publica lo YA generado en Escritorio+GitHub+Drive
                                 + Recetas Cheque ISP a Drive (carpeta CONFIDENCIAL,
                                 RUT pacientes — excepción autorizada 2026-06-30)
                                 (no descarga de SSASUR ni recalcula — usa esto
                                 para republicar rápido tras un cambio manual)

  2 - Gestion Territorial  Lo del ÚLTIMO rango queda al frente; lo anterior, en Historial\\
  3 - Recetas Cheque       Acceso directo a la carpeta LOCAL (no sube datos de pacientes a la nube)
  5 - Pedido Fusionado     Pedido_Fusion_AA.xlsx (Farm_Bod + Bod_Farmacos + Dialisis)
  6 - Centinela             Reportes semanales (json + pdf) por semana epidemiológica
  9 - Clozapina             Accesos a carpetas locales de reportes/hemogramas (RUT pacientes, NO se copian a la nube)
  10 - Servicios Farmaceuticos  Recuento mensual QF x actividad (Agenda Médica), un Excel por mes, SIN RUT
  11 - Centinela Inyectables SM  Stock de antipsicóticos de depósito (salud mental ambulatoria),
                            un reporte por fecha en que hubo cambios, SIN RUT

------------------------------------------------------------------------
Nota: estas son COPIAS para consulta. El programa original sigue en
  {work}
No edites los archivos aquí esperando que cambien los cálculos; los
datos oficiales se generan y publican desde esa carpeta.
========================================================================
"""


def _ps_quote(s):
    """Comilla simple de PowerShell (duplica las comillas simples internas)."""
    return "'" + str(s).replace("'", "''") + "'"


def _crear_lnk(ruta_lnk, target, descripcion, icono=None):
    """Crea un acceso directo .lnk (a archivo o carpeta) con WScript.Shell."""
    if not os.path.exists(target):
        REP.say(f"  [aviso] no existe el destino del acceso: {target}")
        return False
    workdir = target if os.path.isdir(target) else os.path.dirname(target)
    ps = (
        "$ws = New-Object -ComObject WScript.Shell; "
        f"$s = $ws.CreateShortcut({_ps_quote(ruta_lnk)}); "
        f"$s.TargetPath = {_ps_quote(target)}; "
        f"$s.WorkingDirectory = {_ps_quote(workdir)}; "
        f"$s.Description = {_ps_quote(descripcion)}; "
    )
    if icono:
        ps += f"$s.IconLocation = {_ps_quote(icono)}; "
    ps += "$s.WindowStyle = 1; $s.Save()"
    try:
        r = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-Command", ps],
            capture_output=True, text=True,
        )
        if r.returncode != 0:
            REP.say(f"  [aviso] no se pudo crear {os.path.basename(ruta_lnk)}: {r.stderr.strip()}")
            return False
        return True
    except OSError as e:
        REP.say(f"  [aviso] PowerShell no disponible para accesos: {e}")
        return False


_ACCESOS = [
    ("AUTO_SSASUR.lnk",                 "AUTO_SSASUR.bat",    "Descarga de SSASUR, recalcula todo y publica",  _ICON_REFRESH),
    ("Gestion Territorial.lnk",         "GT.bat",             "Descarga y genera las planillas de GT",         _ICON_DOC),
    ("Recetas Cheque ISP.lnk",          "RECETAS_CHEQUE.bat", "Actualiza el registro ISP del mes",             _ICON_RUN),
    ("Clozapina.lnk",                   "CLOZAPINA.bat",      "AUTO_SSASUR + hemogramas HCE -> Excel para ingreso MINSAL", _ICON_RUN),
    ("Sincronizar Todo.lnk",            "SINCRONIZAR_TODO.bat","Publica lo ya generado: Escritorio + GitHub + Drive + Recetas Cheque (sin SSASUR)", _ICON_REFRESH),
]


def crear_estructura(forzar_lnk=False):
    """Crea carpetas, LEEME y (si faltan o forzar_lnk) los accesos directos."""
    for sub in (SUB_GT, SUB_RCH, SUB_PEDIDO, SUB_CENTINELA, SUB_CLOZAPINA, SUB_SERVICIOS, SUB_CENTINELA_SM):
        os.makedirs(os.path.join(BASE, sub), exist_ok=True)
    try:
        with open(os.path.join(BASE, "LEEME.txt"), "w", encoding="utf-8") as fh:
            fh.write(LEEME.format(work=WORK_DIR))
    except OSError as e:
        REP.say(f"  [aviso] no se pudo escribir LEEME.txt: {e}")

    creados = 0
    for nombre_lnk, bat, desc, icono in _ACCESOS:
        ruta = os.path.join(BASE, nombre_lnk)
        if forzar_lnk or not os.path.exists(ruta):
            if _crear_lnk(ruta, os.path.join(WORK_DIR, bat), desc, icono=icono):
                creados += 1
    if creados:
        REP.say(f"[Accesos] {creados} acceso(s) directo(s) actualizados.")


def escribir_log():
    try:
        ts = datetime.now().strftime("%Y-%m-%d %H:%M")
    except Exception:
        ts = "(sin fecha)"
    cuerpo = (
        f"Última sincronización: {ts}\n"
        f"Copiados: {REP.ok} · Sin cambios: {REP.skip} · Con problemas: {REP.fail}\n"
        + "-" * 56 + "\n" + "\n".join(REP.log) + "\n"
    )
    try:
        with open(os.path.join(BASE, "_ultima_sync.txt"), "w", encoding="utf-8") as fh:
            fh.write(cuerpo)
    except OSError:
        pass


# ── CLI ──────────────────────────────────────────────────────────────────────
def main():
    args = set(a.lower() for a in sys.argv[1:])
    print("=" * 60)
    print(f"  Publicando resultados en: {BASE}")
    print("=" * 60)

    crear_estructura(forzar_lnk="--enlaces" in args)
    if args == {"--enlaces"}:
        escribir_log()
        print("\nListo: carpetas y accesos directos actualizados.")
        return

    selectivo = args & {"--gt", "--rch", "--pedido",
                         "--centinela", "--clozapina",
                         "--servicios", "--centinela-sm"}
    todo = not selectivo

    if todo or "--gt" in args:
        sync_gt()
    if todo or "--rch" in args:
        sync_rch()
    if todo or "--pedido" in args:
        sync_pedido()
    if todo or "--centinela" in args:
        sync_centinela()
    if todo or "--clozapina" in args:
        sync_clozapina()
    if todo or "--servicios" in args:
        sync_servicios()
    if todo or "--centinela-sm" in args:
        sync_centinela_sm()

    escribir_log()
    print(f"\nListo ({REP.ok} copiados, {REP.skip} sin cambios). "
          f"Abre «{NOMBRE_CARPETA}» en tu Escritorio.")


if __name__ == "__main__":
    main()
