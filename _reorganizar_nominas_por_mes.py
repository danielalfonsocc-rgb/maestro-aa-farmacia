#!/usr/bin/env python3
"""Reorganiza 04_Farmacia_Gestion_Territorial/<ESTAB>/Nóminas de Envío/ de
carpetas por fecha individual (con duplicados "_Acumulado") a carpetas por
mes ("JULIO 2026", etc.), mismo estilo que las hojas del maestro GT.
Uso: py _reorganizar_nominas_por_mes.py --dry-run
     py _reorganizar_nominas_por_mes.py
"""
import hashlib
import os
import re
import shutil
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import gt_maestro as GM  # solo para MESES_ES

GT_DIR = r"C:\Users\danie\Downloads\04_Farmacia_Gestion_Territorial"
FECHA_RE = re.compile(r"^(\d{2})-(\d{2})-(\d{4})(_Acumulado)?$")

dry_run = "--dry-run" in sys.argv


def _md5(path):
    h = hashlib.md5()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def _mes_carpeta(dd, mm, yyyy):
    return f"{GM.MESES_ES[int(mm) - 1]} {yyyy}"


total_movidos, total_dedup, total_renombrados = 0, 0, 0

for estab in sorted(os.listdir(GT_DIR)):
    nominas_dir = os.path.join(GT_DIR, estab, "Nóminas de Envío")
    if not os.path.isdir(nominas_dir):
        continue

    ocupados_por_mes = {}
    subcarpetas_fecha = []
    for nombre in sorted(os.listdir(nominas_dir)):
        ruta = os.path.join(nominas_dir, nombre)
        m = FECHA_RE.match(nombre)
        if os.path.isdir(ruta) and m:
            dd, mm, yyyy = m.group(1), m.group(2), m.group(3)
            subcarpetas_fecha.append((ruta, nombre, dd, mm, yyyy))

    if not subcarpetas_fecha:
        continue

    print(f"=== {estab} ===")
    for ruta, nombre, dd, mm, yyyy in subcarpetas_fecha:
        carpeta_mes = _mes_carpeta(dd, mm, yyyy)
        destino_dir = os.path.join(nominas_dir, carpeta_mes)
        if not dry_run:
            os.makedirs(destino_dir, exist_ok=True)
        # Rastrea destinos "ya ocupados" en esta corrida (real o simulados en
        # dry-run) — así el dry-run detecta colisiones entre dos carpetas de
        # fecha que aún no se movieron de verdad, no solo contra el disco.
        ocupados = ocupados_por_mes.setdefault(destino_dir, {})
        for nombre_ya, ruta_ya in [(f, os.path.join(destino_dir, f))
                                    for f in (os.listdir(destino_dir) if os.path.isdir(destino_dir) else [])]:
            ocupados.setdefault(nombre_ya, ruta_ya)

        for archivo in sorted(os.listdir(ruta)):
            origen = os.path.join(ruta, archivo)
            if not os.path.isfile(origen):
                continue
            nombre_final = archivo
            if nombre_final in ocupados:
                origen_previo = ocupados[nombre_final]
                if _md5(origen) == _md5(origen_previo):
                    print(f"  [dedup] {nombre}/{archivo} == {carpeta_mes}/{archivo} (idéntico, se descarta)")
                    total_dedup += 1
                    continue
                nombre_final = f"{dd}-{mm}-{yyyy}_{archivo}"
                print(f"  [renombrado] {nombre}/{archivo} -> {carpeta_mes}/{nombre_final} (distinto contenido)")
                total_renombrados += 1
            else:
                print(f"  {nombre}/{archivo} -> {carpeta_mes}/{archivo}")
            destino = os.path.join(destino_dir, nombre_final)
            total_movidos += 1
            if not dry_run:
                shutil.move(origen, destino)
                ocupados[nombre_final] = destino  # ya se movió — origen ya no existe
            else:
                ocupados[nombre_final] = origen  # dry-run: nada se mueve de verdad, origen sigue válido

        if not dry_run:
            # la carpeta de fecha debería quedar vacía (solo tenía archivos)
            restantes = os.listdir(ruta)
            if not restantes:
                os.rmdir(ruta)
            else:
                print(f"  [AVISO] {nombre} no quedó vacía, no se borra: {restantes}")

print(f"\n{'[DRY-RUN] ' if dry_run else ''}Total: {total_movidos} movidos, {total_dedup} deduplicados, {total_renombrados} renombrados por colisión")
