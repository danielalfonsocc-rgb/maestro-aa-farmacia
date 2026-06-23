#!/usr/bin/env python3
"""
auditoria_medicamento.py — Auditoría clínica para CUALQUIER medicamento del AA.
Generaliza auditoria_empagliflozina.py: CMP dispensado, prescriptores,
diagnósticos y duplicidad de prescripción.

Uso:
    py auditoria_medicamento.py --contiene METFORMINA
    py auditoria_medicamento.py --contiene EMPAGLIFLOZINA --dosis 10
    py auditoria_medicamento.py --contiene ENALAPRIL --dosis "20 MG"
    py auditoria_medicamento.py --contiene INSULINA GLARGINA
"""
import argparse
import os
import sys

import pandas as pd

from utils_aa import cargar_recetas_csv, setup_stdout

setup_stdout()
pd.set_option("display.width", 160)
pd.set_option("display.max_colwidth", 52)

WORK = os.path.dirname(os.path.abspath(__file__))

COLS = [
    "ID Receta Detalle", "Prescripción", "RUN",
    "Nombre Profesional", "Apellido Paterno Profesional",
    "Apellido Materno Profesional", "Especialidad",
    "Cod. Diagnóstico 1", "Diagnóstico 1", "Hipótesis Diagnóstica",
    "Cantidad Recetada", "Cantidad Entregada", "Estado Prescripción",
    "Fecha Atención", "Fecha Entrega Receta", "Número Receta",
]


def main():
    ap = argparse.ArgumentParser(
        description="Auditoría clínica de un medicamento en las recetas históricas AA."
    )
    ap.add_argument(
        "--contiene", required=True, nargs="+",
        help="Palabras que debe contener el nombre de Prescripción (AND implícito). "
             "Ej: --contiene EMPAGLIFLOZINA 10",
    )
    ap.add_argument(
        "--dosis", default=None,
        help="Filtro adicional de dosis (substring en Prescripción). Ej: --dosis '10 MG'",
    )
    args = ap.parse_args()

    terminos = [t.upper() for t in args.contiene]
    if args.dosis:
        terminos.append(args.dosis.upper())

    etiqueta = " ".join(terminos)

    print("=" * 72)
    print(f"  AUDITORÍA: {etiqueta}")
    print("=" * 72)

    try:
        rec = cargar_recetas_csv(WORK, cols=COLS)
    except FileNotFoundError as e:
        print(f"[ERROR] {e}")
        sys.exit(1)

    print(f"Recetas totales (dedup ID Receta Detalle): {len(rec):,}")

    presc = rec["Prescripción"].fillna("").str.upper()
    mask = pd.Series([True] * len(rec), index=rec.index)
    for t in terminos:
        mask &= presc.str.contains(t, regex=False)

    sub = rec[mask].copy()
    print(f"\nFilas para '{etiqueta}': {len(sub):,}")
    if sub.empty:
        print("  [Sin resultados] Revisa los términos de búsqueda.")
        sys.exit(0)

    print("Variantes de nombre incluidas:")
    for n, k in sub["Prescripción"].value_counts().head(10).items():
        print(f"   {k:>6,}  {n}")

    # Tipado
    sub["_ent"]   = pd.to_numeric(sub["Cantidad Entregada"], errors="coerce").fillna(0)
    sub["_rec"]   = pd.to_numeric(sub["Cantidad Recetada"],  errors="coerce").fillna(0)
    sub["_fent"]  = pd.to_datetime(sub["Fecha Entrega Receta"], dayfirst=True, errors="coerce")
    sub["_fat"]   = pd.to_datetime(sub["Fecha Atención"],       dayfirst=True, errors="coerce")
    sub["_fecha"] = sub["_fent"].fillna(sub["_fat"])
    sub["_medico"] = (
        sub["Nombre Profesional"].fillna("") + " " +
        sub["Apellido Paterno Profesional"].fillna("") + " " +
        sub["Apellido Materno Profesional"].fillna("")
    ).str.strip().str.upper()

    # ── 1) CONSUMO MENSUAL ───────────────────────────────────────────────────
    print("\n" + "=" * 72)
    print("1) CONSUMO MENSUAL (unidades entregadas)")
    print("=" * 72)
    con = sub[sub["_ent"] > 0].copy()
    con["_mes"] = con["_fecha"].dt.to_period("M")
    mensual = con.groupby("_mes")["_ent"].sum().sort_index()
    for m, v in mensual.items():
        print(f"   {m}:  {int(v):>10,} ud")
    full = mensual.iloc[:-1] if len(mensual) > 1 else mensual
    print(f"\n   Total entregado periodo  : {int(mensual.sum()):,} ud")
    print(f"   Meses completos          : {len(full)}")
    cmp_disp = full.mean() if not full.empty else 0
    print(f"   CMP DISPENSADO (compl.)  : {cmp_disp:,.0f} ud/mes  ← consumo real")
    print(f"   CMP DISPENSADO (todos)   : {mensual.mean():,.0f} ud/mes")

    sub["_mesp"] = sub["_fecha"].dt.to_period("M")
    rec_mes = sub.groupby("_mesp")["_rec"].sum().sort_index()
    rec_full = rec_mes.iloc[:-1] if len(rec_mes) > 1 else rec_mes
    n_pend = int((sub["_ent"] == 0).sum())
    tot_r, tot_e = sub["_rec"].sum(), sub["_ent"].sum()
    print(f"\n   CMP PRESCRITO (compl.)   : {rec_full.mean():,.0f} ud/mes  ← base del maestro")
    print(f"   Total prescrito periodo  : {int(tot_r):,} ud")
    print(f"   Total dispensado periodo : {int(tot_e):,} ud   ({100*tot_e/tot_r:.0f}% de lo prescrito)" if tot_r else "")
    print(f"   Líneas con entrega = 0   : {n_pend:,} de {len(sub):,} ({100*n_pend/len(sub):.0f}%)")

    # ── 2) MAYORES PRESCRIPTORES ─────────────────────────────────────────────
    print("\n" + "=" * 72)
    print("2) MAYORES PRESCRIPTORES (por unidades entregadas)")
    print("=" * 72)
    g = sub.groupby("_medico").agg(
        recetas=("ID Receta Detalle", "count"),
        pacientes=("RUN", "nunique"),
        ud_entreg=("_ent", "sum"),
        especialidad=("Especialidad", lambda s: s.dropna().mode().iloc[0] if not s.dropna().empty else ""),
    ).sort_values("ud_entreg", ascending=False)
    tot_ud = g["ud_entreg"].sum()
    print(f"   {'Médico':<34}{'Esp.':<18}{'Recetas':>8}{'Pacient.':>9}{'Ud entreg':>11}{'%':>7}")
    for med, r in g.head(15).iterrows():
        pct = 100 * r["ud_entreg"] / tot_ud if tot_ud else 0
        print(f"   {med[:33]:<34}{str(r['especialidad'])[:17]:<18}"
              f"{int(r['recetas']):>8,}{int(r['pacientes']):>9,}{int(r['ud_entreg']):>11,}{pct:>6.1f}%")
    print(f"   Prescriptores distintos: {g.shape[0]}")

    # ── 3) DIAGNÓSTICOS ──────────────────────────────────────────────────────
    print("\n" + "=" * 72)
    print("3) DIAGNÓSTICOS ASOCIADOS (Diagnóstico 1)")
    print("=" * 72)
    sub["_dx"] = (
        sub["Cod. Diagnóstico 1"].fillna("").str.strip() + " · " +
        sub["Diagnóstico 1"].fillna("(sin diagnóstico)").str.strip()
    )
    gd = sub.groupby("_dx").agg(
        recetas=("ID Receta Detalle", "count"),
        pacientes=("RUN", "nunique"),
    ).sort_values("recetas", ascending=False)
    for dx, r in gd.head(15).iterrows():
        print(f"   {int(r['recetas']):>6,} rec · {int(r['pacientes']):>5,} pac   {dx[:62]}")
    sin_dx = (sub["Diagnóstico 1"].fillna("").str.strip() == "").sum()
    print(f"   Recetas sin Diagnóstico 1: {sin_dx:,} ({100*sin_dx/len(sub):.1f}%)")

    # ── 4) DUPLICIDAD POR PACIENTE ───────────────────────────────────────────
    print("\n" + "=" * 72)
    print("4) DUPLICIDAD DE PRESCRIPCIÓN (por paciente)")
    print("=" * 72)
    sub["_mes4"] = sub["_fecha"].dt.to_period("M")
    pac = sub.groupby("RUN").agg(
        recetas=("ID Receta Detalle", "count"),
        n_recetas=("Número Receta", "nunique"),
        medicos=("_medico", "nunique"),
        meses=("_mes4", "nunique"),
    )
    print(f"   Pacientes distintos     : {pac.shape[0]:,}")
    dup_mes = sub.groupby(["RUN", "_mes4"])["Número Receta"].nunique()
    ru_dup = dup_mes[dup_mes >= 2].index.get_level_values(0).unique()
    print(f"   Con ≥2 recetas en 1 mes : {len(ru_dup):,}")
    multi_med = pac[pac["medicos"] >= 2]
    print(f"   Con ≥2 médicos distintos: {multi_med.shape[0]:,}")
    print(f"\n   Ejemplos de mayor duplicidad (RUN · recetas · N° únicos · médicos · meses):")
    ej = pac.sort_values(["n_recetas", "medicos"], ascending=False).head(12)
    for run, r in ej.iterrows():
        print(f"     {str(run):<13} rec={int(r['recetas']):>3}  "
              f"n_rec={int(r['n_recetas']):>3}  médicos={int(r['medicos'])}  meses={int(r['meses'])}")

    print()


if __name__ == "__main__":
    main()
