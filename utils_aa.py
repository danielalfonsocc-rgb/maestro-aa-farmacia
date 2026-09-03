#!/usr/bin/env python3
"""
utils_aa.py — Utilidades compartidas del proyecto Maestro AA.

Importar desde aquí en lugar de duplicar en cada script:
    from utils_aa import norm_erp, HOMOLOGACION, cargar_recetas_csv, setup_stdout
"""
import datetime
import glob
import os
import re
import sys
import unicodedata

import pandas as pd


def setup_stdout() -> None:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")


# ── Rutas configurables por variable de entorno ──────────────────────────────
# Carpeta y plantilla del formulario ISP de Recetas Cheque: viven fuera del
# repo (son específicas de la máquina de la QF), así que estaban hardcodeadas
# en 3 scripts distintos. Se centralizan acá con default = la ruta actual,
# y se pueden sobrescribir sin tocar código en cualquier otra máquina:
#   set MAESTRO_RCH_DIR=D:\OtraCarpeta\Farmacia_AT_Abierta_RCh
#   set MAESTRO_PLANTILLA_RCH=D:\OtraCarpeta\Formulario-Notificacion-...xlsx
RCH_DIR = os.environ.get(
    "MAESTRO_RCH_DIR",
    r"C:\Users\danie\Downloads\Farmacia_AT_Abierta_RCh\Farmacia_AT_Abierta_RCh",
)
PLANTILLA_BLANCO_RCH = os.environ.get(
    "MAESTRO_PLANTILLA_RCH",
    r"C:\Users\danie\Downloads\02_Farmacia_Recetas_e_Informes_CSV\Formulario-Notificacion-Recetas-Cheque_v11.xlsx",
)
# Planilla maestra de Gestión Territorial (histórico por mes, uso de gt_maestro.py):
#   set MAESTRO_GT_MAESTRO_XLSX=D:\OtraCarpeta\GT PITRUFQUEN 2026.xlsx
GT_MAESTRO_XLSX = os.environ.get(
    "MAESTRO_GT_MAESTRO_XLSX",
    r"C:\Users\danie\Downloads\GT PITRUFQUEN 2026 (2).xlsx",
)
# Carpetas con CSV de recetas de ciclos anteriores (fuera del repo, la carpeta
# de descargas del PC de la QF acumula meses de informe_completo_recetas*.csv
# que ya salieron de la rotación del repo). Usadas solo por
# cargar_recetas_historico() para auditorías que necesitan todo el histórico
# disponible, no por el flujo normal (maestro_aa.py etc. solo usan el repo):
#   set MAESTRO_RECETAS_HIST_DIRS=D:\OtraCarpeta\CSV_Viejos;D:\OtraCarpeta2
RECETAS_HIST_DIRS = [
    d.strip() for d in os.environ.get(
        "MAESTRO_RECETAS_HIST_DIRS",
        r"C:\Users\danie\Downloads\02_Farmacia_Recetas_e_Informes_CSV;"
        r"C:\Users\danie\Downloads",
    ).split(";") if d.strip()
]
# Planilla oficial de Actas de Vencimiento (documento legal firmado, vive fuera
# del repo). Usada por crear_acta_vencimiento.py:
#   set MAESTRO_ACTAS_VENCIMIENTO=D:\OtraCarpeta\ACTAS DE VENCIMIENTO 2026.xlsx
ACTAS_VENCIMIENTO_PATH = os.environ.get(
    "MAESTRO_ACTAS_VENCIMIENTO",
    r"C:\Users\danie\Downloads\ACTAS DE VENCIMIENTO 2026.xlsx",
)


# Blindaje contra datos auto-detectados desactualizados (incidente 2026-07-13:
# AUTO_SSASUR puede fallar/omitir la descarga de recetas en silencio —solo
# imprime [AVISO] y sigue— y un script que auto-detecta "el CSV más reciente
# en disco" puede terminar usando uno viejo sin darse cuenta. Usado por
# centinela_reporte.py y recetas_cheque.py.
UMBRAL_DIAS_STALE = 10


def verificar_frescura(fecha_dato: "datetime.date | None", etiqueta: str,
                        hoy: "datetime.date | None" = None) -> None:
    """Aborta con exit(1) si `fecha_dato` (la más reciente encontrada en una
    fuente auto-detectada) está a más de UMBRAL_DIAS_STALE días de hoy.

    Solo debe llamarse cuando la fuente fue auto-detectada (no cuando el
    usuario pasó explícitamente --csv/--xlsx/--form para reprocesar algo
    histórico a propósito).
    """
    hoy = hoy or datetime.date.today()
    if fecha_dato is None:
        print(f"  [ERROR] No se pudo determinar la fecha más reciente de «{etiqueta}» — abortando.")
        sys.exit(1)
    brecha = (hoy - fecha_dato).days
    if brecha > UMBRAL_DIAS_STALE:
        print(f"\n  [ERROR] «{etiqueta}» está desactualizada: el dato más reciente es del "
              f"{fecha_dato.strftime('%d/%m/%Y')} ({brecha} días atrás; hoy {hoy.strftime('%d/%m/%Y')}).")
        print(f"  Se aborta para no procesar/publicar con datos viejos (ver incidente S.52).")
        print(f"  Verifica que AUTO_SSASUR haya descargado los datos de esta semana y vuelve a intentar.")
        sys.exit(1)


def norm_erp(s: str) -> str:
    """Normaliza nombres de medicamentos: NFD + sin diacríticos + espacios únicos."""
    s = unicodedata.normalize("NFD", str(s).upper())
    s = "".join(c for c in s if not unicodedata.combining(c))
    return re.sub(r" {2,}", " ", s).strip()


# Tabla canónica de homologación (fuente: maestro_aa.py, versión más completa).
# Los subsets que tenían auditoria_prescripcion.py y agente_duplicados.py
# son estrictamente un subconjunto de esta tabla.
HOMOLOGACION_RAW: dict[str, str] = {
    "VITAMINA D3 800 UI CAPS":                                 "VITAMINA D3 800 UI CM",
    "BUPROPION 150 MG COMPRIMIDO LIBERACION MODIFICADA":       "BUPROPION (ANFEBUTAMONA) 150 MG CM LIBERACION MODIFICADA",
    "ACIDO ALENDRONICO 70 MG CM.":                             "ACIDO ALENDRONICO  CM 70 MG",
    "ACIDO FOLICO 1 MG COMPRIMIDO":                            "ACIDO FOLICO  CM 1 MG",
    "ACIDO FOLICO 5 MG COMPRIMIDO":                            "ACIDO FOLICO CM 5 MG",
    "ACIDO URSODEOXICOLICO 250 MG COMPRIMIDO":                 "ACIDO URSODEOXICOLICO CM 250 MG",
    "ACETAZOLAMIDA 250 MG COMPRIMIDO":                         "ACETAZOLAMIDA  CM 250 MG",
    "ACETAZOLAMIDA 250 MG CM UD":                               "ACETAZOLAMIDA  CM 250 MG",
    "TRAMADOL 100 MG/ML FC 10 ML":                             "TRAMADOL FRASCO GOTAS 100 MG/ML /10 ML",
    "LAGRIMAS ARTIFICIALES":                                   "LAGRIMAS ARTIFICIALES 0,4% X 10 ML",
    "BUDESONIDA 200 MCG/DO INH FC 200DO":                      "BUDESONIDA 200 MCG/DO INH FC",
    "RISPERIDONA 1 MG/ ML FC 30 ML":                           "RISPERIDONA 1 MG/ ML FC X 30 ML.",
    "FIERRO COMPLEJO HIERRO III POLIMALTOSA 100 MG":           "FIERRO COMPLEJO HIERRO III POLIMALTOSA  COMPRIMIDOS100 MG",
    "LORATADINA 5 MG/5 ML SOLUCION ORAL":                      "LORATADINA 5 MG/5 ML FC 120 ML",
    "BROMURO DE ROCURONIO 50 MG AM (REP)":                     "ROCURONIO BROMURO AM 50 MG/5 ML",
    "BROMURO DE ROCURONIO AMP 100 MG/10 ML":                   "ROCURONIO BROMURO AM 50 MG/5 ML",
    "BICARBONATO DE SODIO 8,4% AM 10 ML":                      "SODIO BICARBONATO 8,4 % AM 10 ML",
    "FERROSO SULFATO 125 MG/ML SOL. ORAL EN GOTAS FC 30 ML":   "FERROSO SULFATO 125 MG/ML SOL. ORAL EN GOTAS FC 30 ML.",
    "LIDOCAINA CLORHIDRATO 2 % AM 5 ML":                       "LIDOCAINA CLORHIDRATO 2 % AM 5ML",
    "ACIDO TRANEXAMICO CM 500 MG":                             "ACIDO TRANEXAMICO 500 MG COMPRIMIDO",
    "ACENOCUMAROL  CM 4 MG":                                   "ACENOCUMAROL 4 MG CM",
    "CINTA PARA DETERMINACION GLICEMIA  USO PACIENTE UD.":     "CINTA PARA DETERMINACION GLUCOSA USO PACIENTE UD.",
    "SALMETEROL Y FLUTICASONA 250 MG/25 MG INH UD":            "SALMETEROL /FLUTICASONA 250 MG/25 MG INH UD",
    # Ferranim = nombre comercial del mismo compuesto (fumarato ferroso 330 mg
    # + cianocobalamina + ac. ascorbico + ac. folico); aparece asi en el
    # reporte de stock de Bodega Farmacos (detectado 21-07-2026)
    "FERRANIM FUMARATO FERROSO 330 MG":                        "FUMARATO FERROSO + CIANOCOBALAMINA + ACIDO ASCORBICO + ACIDO FOLICO 330/1/100/2 MG CP",
    # Conteo 31-08-2026: nombres en planilla impresa difieren del ERP
    "ACIDO ACETILSALICILICO 100 MG COMPRIMIDO":               "ACIDO ACETILSALICILICO CM 100 MG",
    "GEMFIBROZILO CP 300 MG":                                  "GEMFIBROZILO 600 MG COMPRIMIDO",
}
HOMOLOGACION: dict[str, str] = {norm_erp(k): norm_erp(v) for k, v in HOMOLOGACION_RAW.items()}


def cargar_recetas_csv(
    work_dir: str,
    cols: list[str] | None = None,
    solo_ultimo: bool = False,
) -> pd.DataFrame:
    """Carga y deduplica los CSV de recetas SSASUR (informe_completo_recetas*.csv).

    Devuelve DataFrame crudo (strings) con ID Receta Detalle deduplicado.
    Cada archivo consumidor aplica su propio filtrado y tipado encima.

    Args:
        work_dir:     Carpeta donde están los CSVs.
        cols:         Si se indica, solo carga esas columnas (más rápido).
        solo_ultimo:  True → usa solo el CSV más reciente (flag --rapido).
    """
    files = sorted(glob.glob(os.path.join(work_dir, "informe_completo_recetas*.csv")))
    if not files:
        raise FileNotFoundError(
            "No hay archivos informe_completo_recetas*.csv en la carpeta.\n"
            "Ejecuta AUTO_SSASUR.bat primero."
        )
    if solo_ultimo:
        files = files[-1:]
    chunks = []
    for f in files:
        kw: dict = {"encoding": "latin1", "sep": ";", "on_bad_lines": "skip", "dtype": str}
        if cols:
            kw["usecols"] = lambda c, _c=cols: c in _c
        chunks.append(pd.read_csv(f, **kw))
    return (
        pd.concat(chunks, ignore_index=True)
        .drop_duplicates(subset=["ID Receta Detalle"], keep="first")
    )


_RE_FECHA_ARCHIVO_B = re.compile(r"_b(\d{4})(\d{2})(\d{2})")
_RE_FECHA_ARCHIVO_TS = re.compile(r"recetas(\d{2})(\d{2})(\d{4})_(\d{2})(\d{2})(\d{2})")


def _fecha_archivo_recetas(path: str) -> pd.Timestamp:
    """Extrae la fecha de descarga desde el nombre del archivo (para ordenar)."""
    name = os.path.basename(path)
    m = _RE_FECHA_ARCHIVO_B.search(name)
    if m:
        y, mo, d = m.groups()
        return pd.Timestamp(int(y), int(mo), int(d))
    m = _RE_FECHA_ARCHIVO_TS.search(name)
    if m:
        d, mo, y, hh, mm, ss = m.groups()
        return pd.Timestamp(int(y), int(mo), int(d), int(hh), int(mm), int(ss))
    return pd.Timestamp(1970, 1, 1)


def cargar_recetas_historico(work_dir: str, cols: list[str] | None = None) -> pd.DataFrame:
    """Carga TODO el histórico de recetas disponible, combinando el repo con las
    carpetas de respaldo fuera del repo (RECETAS_HIST_DIRS).

    A diferencia de cargar_recetas_csv() (que solo usa work_dir, la rotación
    reciente que usa el flujo normal de maestro_aa.py), esta función es para
    auditorías puntuales que necesitan profundidad histórica (ej. "¿desde
    cuándo retira este paciente tal medicamento?"). Es más lenta (decenas de
    CSV) — pensada para usarse con cache (ej. st.cache_data en una app), no en
    el flujo de consolidación diario.

    Dedup por "ID Receta Detalle" con keep="last" tras ordenar por fecha del
    archivo: una misma línea de receta puede aparecer con "Cantidad Entregada"
    desactualizada en una descarga vieja y completa en una más nueva.
    """
    dirs = [work_dir, os.path.join(work_dir, "_csv_bak"), *RECETAS_HIST_DIRS]
    files: list[str] = []
    vistos: set[str] = set()
    for d in dirs:
        for f in glob.glob(os.path.join(d, "informe_completo_recetas*.csv")):
            base = os.path.basename(f)
            if base in vistos:
                continue
            vistos.add(base)
            files.append(f)
    if not files:
        raise FileNotFoundError(
            "No se encontró ningún informe_completo_recetas*.csv en el repo "
            "ni en RECETAS_HIST_DIRS."
        )
    files.sort(key=_fecha_archivo_recetas)

    chunks = []
    for f in files:
        kw: dict = {"encoding": "latin1", "sep": ";", "on_bad_lines": "skip", "dtype": str}
        if cols:
            kw["usecols"] = lambda c, _c=cols: c in _c
        try:
            chunks.append(pd.read_csv(f, **kw))
        except Exception:
            continue
    return (
        pd.concat(chunks, ignore_index=True)
        .drop_duplicates(subset=["ID Receta Detalle"], keep="last")
    )
