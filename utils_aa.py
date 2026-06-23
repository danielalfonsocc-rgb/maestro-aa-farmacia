#!/usr/bin/env python3
"""
utils_aa.py — Utilidades compartidas del proyecto Maestro AA.

Importar desde aquí en lugar de duplicar en cada script:
    from utils_aa import norm_erp, HOMOLOGACION, cargar_recetas_csv, setup_stdout
"""
import glob
import os
import re
import sys
import unicodedata

import pandas as pd


def setup_stdout() -> None:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")


def norm_erp(s: str) -> str:
    """Normaliza nombres de medicamentos: NFD + sin diacríticos + espacios únicos."""
    s = unicodedata.normalize("NFD", str(s).upper())
    s = "".join(c for c in s if not unicodedata.combining(c))
    return re.sub(r" {2,}", " ", s).strip()


# Tabla canónica de homologación (fuente: maestro_aa.py, versión más completa).
# Los subsets que tenían auditoria_prescripcion.py, agente_duplicados.py y
# recetas_duplicadas.py son estrictamente un subconjunto de esta tabla.
HOMOLOGACION_RAW: dict[str, str] = {
    "TRAZODONA CM 100 MG":                                     "TRAZODONA CM  100 MG",
    "VITAMINA D3 800 UI CAPS":                                 "VITAMINA D3 800 UI CM",
    "BUPROPION 150 MG COMPRIMIDO LIBERACION MODIFICADA":       "BUPROPION (ANFEBUTAMONA) 150 MG CM LIBERACION MODIFICADA",
    "ACIDO ALENDRONICO 70 MG CM.":                             "ACIDO ALENDRONICO  CM 70 MG",
    "ACIDO FOLICO 1 MG COMPRIMIDO":                            "ACIDO FOLICO  CM 1 MG",
    "ACIDO FOLICO 5 MG COMPRIMIDO":                            "ACIDO FOLICO CM 5 MG",
    "ACIDO URSODEOXICOLICO 250 MG COMPRIMIDO":                 "ACIDO URSODEOXICOLICO CM 250 MG",
    "ACETAZOLAMIDA 250 MG COMPRIMIDO":                         "ACETAZOLAMIDA  CM 250 MG",
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
