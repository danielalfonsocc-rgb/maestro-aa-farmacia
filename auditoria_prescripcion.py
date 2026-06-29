#!/usr/bin/env python3
"""
Pre-cálculo de AUDITORÍA DE PRESCRIPCIÓN por medicamento → auditoria_prescripcion.json
(la app lo lee; los CSV de recetas no se publican a la nube, este JSON sí).

Para cada medicamento: consumo mensual (prescrito vs dispensado), CMP, mayores
prescriptores, diagnósticos asociados y duplicidad de prescripción.
"""
import glob, json, os, re, sys
import pandas as pd

from utils_aa import norm_erp, HOMOLOGACION, cargar_recetas_csv, setup_stdout

setup_stdout()
WORK = os.path.dirname(os.path.abspath(__file__))


def main():
    cols = ["ID Receta Detalle", "Prescripción", "RUN",
            "Nombre Profesional", "Apellido Paterno Profesional", "Apellido Materno Profesional",
            "Especialidad", "Cod. Diagnóstico 1", "Diagnóstico 1",
            "Cantidad Recetada", "Cantidad Entregada", "Estado Prescripción",
            "Fecha Atención", "Fecha Entrega Receta", "Número Receta"]
    try:
        rec = cargar_recetas_csv(WORK, cols=cols)
    except FileNotFoundError:
        print("[AVISO] No hay CSV de recetas — no se genera la auditoría.")
        return
    print(f"Recetas dedup: {len(rec):,}")

    rec["_med"] = rec["Prescripción"].fillna("").apply(norm_erp).map(lambda x: HOMOLOGACION.get(x, x))
    rec["_ent"] = pd.to_numeric(rec["Cantidad Entregada"], errors="coerce").fillna(0)
    rec["_rec"] = pd.to_numeric(rec["Cantidad Recetada"], errors="coerce").fillna(0)
    _fe = pd.to_datetime(rec["Fecha Entrega Receta"], dayfirst=True, errors="coerce")
    _fa = pd.to_datetime(rec["Fecha Atención"], dayfirst=True, errors="coerce")
    rec["_fecha"] = _fe.fillna(_fa)
    rec["_mes"] = rec["_fecha"].dt.to_period("M").astype(str)            # mes de ENTREGA (consumo)
    rec["_fpresc"] = _fa.fillna(_fe)                                     # fecha de PRESCRIPCIÓN (atención)
    rec["_mespresc"] = rec["_fpresc"].dt.to_period("M").astype(str)      # mes de prescripción (duplicidad)
    rec["_medico"] = (rec["Nombre Profesional"].fillna("") + " " +
                      rec["Apellido Paterno Profesional"].fillna("") + " " +
                      rec["Apellido Materno Profesional"].fillna("")).str.strip().str.upper()
    rec["_dx"] = (rec["Cod. Diagnóstico 1"].fillna("").str.strip() + " · " +
                  rec["Diagnóstico 1"].fillna("(sin diagnóstico)").str.strip()).str.strip(" ·")

    out = {}
    # Solo medicamentos con consumo dispensado real (evita ruido y limita tamaño)
    meds = [m for m in rec["_med"].unique() if m and m != ""]
    for med in meds:
        sub = rec[rec["_med"] == med]
        if sub["_ent"].sum() <= 0 and len(sub) < 5:
            continue
        # Consumo mensual
        disp = sub[sub["_ent"] > 0].groupby("_mes")["_ent"].sum().sort_index()
        disp = disp[disp.index != "NaT"]
        meses = {m: int(v) for m, v in disp.items()}
        full = list(disp.values[:-1]) if len(disp) > 1 else list(disp.values)
        cmp_disp = round(sum(full) / len(full)) if full else 0
        tot_rec = float(sub["_rec"].sum()); tot_ent = float(sub["_ent"].sum())
        # Prescriptores
        g = sub.groupby("_medico").agg(rec_n=("ID Receta Detalle", "count"),
                                       pac=("RUN", "nunique"), ud=("_ent", "sum"),
                                       esp=("Especialidad", lambda s: s.dropna().mode().iloc[0] if not s.dropna().empty else ""))
        g = g.sort_values("ud", ascending=False)
        tot_ud = g["ud"].sum() or 1
        prescriptores = [{"medico": m[:40], "esp": str(r["esp"])[:18], "recetas": int(r["rec_n"]),
                          "pacientes": int(r["pac"]), "unidades": int(r["ud"]),
                          "pct": round(100 * r["ud"] / tot_ud, 1)}
                         for m, r in g.head(10).iterrows()]
        # Diagnósticos
        gd = sub.groupby("_dx").agg(rec_n=("ID Receta Detalle", "count"),
                                    pac=("RUN", "nunique")).sort_values("rec_n", ascending=False)
        diagnosticos = [{"dx": dx[:60] or "(sin diagnóstico)", "recetas": int(r["rec_n"]),
                         "pacientes": int(r["pac"])} for dx, r in gd.head(10).iterrows()]
        sin_dx = int((sub["Diagnóstico 1"].fillna("").str.strip() == "").sum())
        # Duplicidad — se colapsan las cuotas mensuales de una misma receta anual en
        # EVENTOS de prescripción (RUN+médico+fecha de atención). El ERP SSASUR emite
        # un Número Receta SEPARADO por cada cuota ("X de 12" → ~12 N° el mismo día),
        # así que contar Número Receta inflaba la duplicidad ~4-19× (ver memoria
        # ssasur-cuotas-receta; misma lógica de evento que auditoria_duplicados_profunda.py).
        pac = sub.groupby("RUN").agg(medicos=("_medico", "nunique"))
        ev = sub[sub["_fpresc"].notna()].drop_duplicates(subset=["RUN", "_medico", "_fpresc"])
        ev_mes = ev.groupby(["RUN", "_mespresc"]).size()
        dup_mes_n = ev_mes[ev_mes >= 2].index.get_level_values(0).nunique()
        out[med] = {
            "n_lineas": int(len(sub)), "meses": meses,
            "cmp_dispensado": cmp_disp,
            "total_prescrito": int(tot_rec), "total_dispensado": int(tot_ent),
            "pct_dispensado": round(100 * tot_ent / tot_rec) if tot_rec else 0,
            "pacientes": int(sub["RUN"].nunique()),
            "dup_2mas_medicos": int((pac["medicos"] >= 2).sum()),
            "dup_2mas_recetas_mes": int(dup_mes_n),
            "sin_diagnostico": sin_dx,
            "prescriptores": prescriptores, "diagnosticos": diagnosticos,
        }
    csv_files = sorted(glob.glob(os.path.join(WORK, "informe_completo_recetas*.csv")))
    m = re.search(r"\d{8}", os.path.basename(csv_files[-1])) if csv_files else None
    generado = m.group() if m else ""
    dest = os.path.join(WORK, "auditoria_prescripcion.json")
    with open(dest, "w", encoding="utf-8") as f:
        json.dump({"generado": generado,
                   "n_medicamentos": len(out), "data": out}, f, ensure_ascii=False)
    print(f"[OK] auditoria_prescripcion.json · {len(out)} medicamentos · {os.path.getsize(dest)//1024} KB")


if __name__ == "__main__":
    main()
