#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
subir_recetas_cheque_drive.py — Sube el formulario ISP de Recetas Cheque a Drive.

Acción EXCLUIDA por diseño de publicar_drive.py (contiene RUT de pacientes,
Ley 19.628). Este script existe como subida puntual, autorizada explícitamente
por el usuario pese al riesgo conocido. Sube SOLO este archivo, a una carpeta
separada y claramente marcada como confidencial — no toca el sync automático
ni la lógica de exclusión de publicar_drive.py.

Uso:
  py subir_recetas_cheque_drive.py
"""
import sys
import os
import json
from glob import glob

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import publicar_drive as pdv
from utils_aa import RCH_DIR  # configurable por variable de entorno MAESTRO_RCH_DIR

PREFIJO_FORM = "Formulario-Notificacion-Recetas-Cheque"
CARPETA_DESTINO = "7 - Recetas Cheque (CONFIDENCIAL - RUT pacientes)"


def autodescubrir_form(carpeta):
    if not os.path.isdir(carpeta):
        raise FileNotFoundError(f"No existe la carpeta del formulario ISP:\n  {carpeta}")
    cand = [f for f in glob(os.path.join(carpeta, PREFIJO_FORM + "*.xlsx"))
            if not os.path.basename(f).startswith("~$")]
    if not cand:
        raise FileNotFoundError(
            f"No se encontró formulario '{PREFIJO_FORM}*.xlsx' en:\n  {carpeta}")
    return max(cand, key=os.path.getmtime)


def main():
    rch_file = autodescubrir_form(RCH_DIR)
    print(f"  Archivo: {os.path.basename(rch_file)}")

    print("[Drive] Conectando...")
    svc = pdv._get_service()

    with open(pdv._FOLDER_CACHE_FILE, encoding="utf-8") as f:
        known = json.load(f)
    raiz_id = known["Farmacia AA"]

    cache = {}
    fid = pdv._obtener_o_crear_carpeta(svc, CARPETA_DESTINO, raiz_id, cache)

    stats = {"ok": 0, "skip": 0, "fail": 0}
    r = pdv._subir(svc, rch_file, fid, stats=stats)
    print(f"  {os.path.basename(rch_file)}: {r}")
    print(f"\n[Drive] Carpeta: Farmacia AA / {CARPETA_DESTINO}")
    print(f"[Drive] Folder ID: {fid}")


if __name__ == "__main__":
    main()
