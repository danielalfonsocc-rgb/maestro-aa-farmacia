#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
recetas_cheque.py  —  Registro ISP de Recetas Cheque (Farmacia AT Abierta)
═══════════════════════════════════════════════════════════════════════════
Actualiza el formulario oficial ISP de recetas cheque (estupefacientes y
psicotrópicos) agregando SOLO los folios nuevos. Consume la MISMA sábana
(`informe_completo_recetas*.csv`) que AUTO_SSASUR ya descarga del módulo RECETA.

Qué hace:
  1. Toma la sábana más reciente (o la indicada con --csv).
  2. Filtra:  Tipo Receta = CONTROLADA  ·  Estado = ENTREGADA
              ·  Bodega Despacha = "FARMACIA AT ABIERTA"
  3. Mapea el código HHHA → F-código ISP + nombre + presentación.
  4. Agrega a la planilla SOLO los folios que no estén ya (dedup por N° Folio).
     No toca filas existentes. La QF completa a mano DV QF y Nombre QF.

Versus el script original de la QF, esta versión:
  · Normaliza los nombres de columna con unicodedata (mapea tildes a la letra
    base). El original borraba la vocal acentuada y reventaba con la sábana de
    SSASUR (Número Folio → "Nmero Folio" → KeyError).
  · Autodescubre la sábana (carpeta maestro) y el formulario del mes vigente
    (carpeta RCh), o acepta rutas explícitas por CLI.
  · Filtra por el mes/año de dispensación declarado en el formulario (B7/B8),
    para no mezclar meses cuando la sábana cruza el cambio de mes.
  · Hace copia de respaldo (.bak) antes de escribir.
  · Maneja el caso "Excel abierto" (PermissionError) sin reventar.

Uso:
  py recetas_cheque.py                      # autodescubre todo, pausa al final
  py recetas_cheque.py --csv <sabana.csv> --form <formulario.xlsx> --no-pause
  py recetas_cheque.py --no-backup
═══════════════════════════════════════════════════════════════════════════
NUNCA toca credenciales. Solo lee la sábana ya descargada y escribe la planilla.
"""
import sys
import os
import argparse
import shutil
import unicodedata
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, Border, Side
from datetime import datetime, date
from glob import glob

# ── CONFIGURACIÓN ────────────────────────────────────────────────────────────
MAESTRO_DIR = os.path.dirname(os.path.abspath(__file__))
PREFIJO_CSV = "informe_completo_recetas"

# Carpeta donde la QF mantiene el formulario ISP mensual. Se elige el .xlsx más
# reciente que empiece con este prefijo (maneja el rollover v11_junio → v12_julio).
RCH_DIR     = r"C:\Users\danie\Downloads\Farmacia_AT_Abierta_RCh\Farmacia_AT_Abierta_RCh"
PREFIJO_FORM = "Formulario-Notificacion-Recetas-Cheque"
HOJA_RCH    = "Registro de RCh"
FILA_DATOS  = 12   # los registros empiezan en la fila 12

# ── MAPEO HHHA → F-código ISP + Nombre + Presentación ────────────────────────
MAPEO_HHHA = {
    "212-0009": ("F-1373",  "FENOBARBITAL COMPRIMIDOS 100 mg",                                                                    "COMPRIMIDOS"),
    "212-0052": ("F-24748", "FENTADUR PARCHE TRANSDERMICO 25 mcg/hora (FENTANILO)",                                              "PARCHES"),
    "212-0073": ("F-19539", "PALEXIS RETARD COMPRIMIDOS RECUBIERTOS DE LIBERACION PROLONGADA 50 mg (TAPENTADOL CLORHIDRATO)",    "COMPRIMIDOS"),
    "212-0031": ("F-22274", "VENDAL RETARD COMPRIMIDOS RECUBIERTOS DE LIBERACION PROLONGADA 30 mg (MORFINA CLORHIDRATO TRIHIDRATO)", "CAPSULAS"),
    "214-0597": ("F-7553",  "ARADIX RETARD COMPRIMIDOS DE LIBERACION PROLONGADA 10 mg",                                          "COMPRIMIDOS"),
    "212-0059": ("F-17744", "METILFENIDATO CLORHIDRATO COMPRIMIDOS 10 mg",                                                       "COMPRIMIDOS"),
    "212-0034": ("F-19391", "METADONA CLORHIDRATO COMPRIMIDOS 10 mg",                                                            "COMPRIMIDOS"),
    "212-0083": ("F-27066", "NEUROK CAPSULAS 50 mg (LISDEXANFETAMINA DIMESILATO)",                                               "COMPRIMIDOS"),
    "212-0085": ("F-27069", "NEUROK CAPSULAS 70 mg (LISDEXANFETAMINA DIMESILATO)",                                               "CAPSULAS"),
    "212-0068": ("F-19518", "MORFINA SULFATO SOLUCION ORAL 2%",                                                                  "FRASCO x 20 mL"),
}

# Respaldo por nombre de prescripción cuando el código HHHA no esté en el mapeo.
MAPEO_NOMBRE = {
    "FENTANILO  25 MCG/HORA PARCHE": ("F-24748", "FENTADUR PARCHE TRANSDERMICO 25 mcg/hora (FENTANILO)", "PARCHES"),
}

MESES_ES = {
    "ENERO": 1, "FEBRERO": 2, "MARZO": 3, "ABRIL": 4, "MAYO": 5, "JUNIO": 6,
    "JULIO": 7, "AGOSTO": 8, "SEPTIEMBRE": 9, "SETIEMBRE": 9, "OCTUBRE": 10,
    "NOVIEMBRE": 11, "DICIEMBRE": 12,
}

# ── NORMALIZACIÓN / UTILIDADES ───────────────────────────────────────────────
def norm_col(c):
    """Tilde → letra base (Número→Numero), sin borrar la vocal."""
    return unicodedata.normalize("NFKD", str(c).strip()).encode("ascii", "ignore").decode()


def autodescubrir_csv(carpeta):
    archivos = glob(os.path.join(carpeta, PREFIJO_CSV + "*.csv"))
    if not archivos:
        raise FileNotFoundError(
            f"No se encontró sábana '{PREFIJO_CSV}*.csv' en:\n  {carpeta}")
    return max(archivos, key=os.path.getmtime)


def autodescubrir_form(carpeta):
    if not os.path.isdir(carpeta):
        raise FileNotFoundError(f"No existe la carpeta del formulario ISP:\n  {carpeta}")
    cand = [f for f in glob(os.path.join(carpeta, PREFIJO_FORM + "*.xlsx"))
            if not os.path.basename(f).startswith("~$")]
    if not cand:
        raise FileNotFoundError(
            f"No se encontró formulario '{PREFIJO_FORM}*.xlsx' en:\n  {carpeta}")
    return max(cand, key=os.path.getmtime)


def get_producto(row):
    hhha = str(row["Codigo Prescripcion HHHA"]).strip() if pd.notna(row["Codigo Prescripcion HHHA"]) else ""
    if hhha in MAPEO_HHHA:
        return MAPEO_HHHA[hhha]
    nombre = str(row["Prescripcion"]).strip()
    if nombre in MAPEO_NOMBRE:
        return MAPEO_NOMBRE[nombre]
    return (None, nombre, None)


def split_rut(s):
    if pd.isna(s):
        return None, None
    s = str(s).strip().replace(" ", "")
    if "-" in s:
        partes = s.split("-")
        rut = partes[0].replace(".", "")
        dv  = partes[1].upper()
        try:
            rut = int(rut)
        except ValueError:
            pass
        return rut, dv
    return s, None


def get_posologia(row):
    dosis  = str(row["Dosis"]).strip()    if pd.notna(row["Dosis"])    else ""
    interv = str(row["Intervalo"]).strip() if pd.notna(row["Intervalo"]) else ""
    per    = str(row["Periodo.1"]).strip() if pd.notna(row["Periodo.1"]) else ""
    obs    = str(row["Observacion Medica Prescripcion"]).strip() if pd.notna(row["Observacion Medica Prescripcion"]) else ""
    partes = [p for p in [dosis, interv, per] if p and p not in ("nan", "NaN")]
    return " ".join(partes) if partes else (obs if obs and obs not in ("nan", "NaN") else "SEGUN INDICACION")


EXCEL_EPOCH = date(1899, 12, 30)

def parse_fecha(s):
    if pd.isna(s) or str(s).strip() in ("", "nan", "NaT"):
        return None
    try:
        return datetime.strptime(str(s).strip()[:10], "%d/%m/%Y").date()
    except Exception:
        return None


def to_excel_date(s):
    d = parse_fecha(s)
    return (d - EXCEL_EPOCH).days if d else None


def leer_folios_existentes(ws):
    folios = set()
    for row in ws.iter_rows(min_row=FILA_DATOS, values_only=True):
        if row[0] is not None:
            try:
                folios.add(int(row[0]))
            except (ValueError, TypeError):
                pass
    return folios


def leer_periodo_form(ws):
    """Año/mes de dispensación declarados en el formulario (B7/B8). (año, mes)|(None,None)."""
    try:
        anio = ws["B7"].value
        mes  = ws["B8"].value
        anio = int(str(anio).strip()) if anio is not None else None
        mes  = MESES_ES.get(str(mes).strip().upper()) if mes is not None else None
        return anio, mes
    except Exception:
        return None, None


# ── MAIN ─────────────────────────────────────────────────────────────────────
def main():
    ap = argparse.ArgumentParser(description="Registro ISP de recetas cheque — Farmacia AT Abierta")
    ap.add_argument("--csv",  help="Sábana CSV (default: la más reciente en la carpeta maestro)")
    ap.add_argument("--form", help="Formulario ISP .xlsx (default: el más reciente en la carpeta RCh)")
    ap.add_argument("--carpeta-csv",  default=MAESTRO_DIR, help="Carpeta donde buscar la sábana")
    ap.add_argument("--carpeta-form", default=RCH_DIR,     help="Carpeta donde buscar el formulario")
    ap.add_argument("--no-backup", action="store_true", help="No crear copia .bak antes de escribir")
    ap.add_argument("--no-pause",  action="store_true", help="No esperar Enter al final (modo automático)")
    ap.add_argument("--sin-filtro-mes", action="store_true", help="No filtrar por el mes del formulario")
    a = ap.parse_args()

    print("=" * 62)
    print("  Registro ISP Recetas Cheque — Farmacia AT Abierta")
    print("=" * 62)

    csv_path  = a.csv  or autodescubrir_csv(a.carpeta_csv)
    form_path = a.form or autodescubrir_form(a.carpeta_form)
    print("  Sábana     : " + os.path.basename(csv_path))
    print("  Formulario : " + os.path.basename(form_path))

    # 1) Leer y normalizar columnas de la sábana.
    df = pd.read_csv(csv_path, encoding="latin-1", sep=None, engine="python")
    df = df.rename(columns={c: norm_col(c) for c in df.columns})

    requeridas = ["Tipo Receta", "Estado", "Numero Folio", "Bodega Despacha",
                  "Codigo Prescripcion HHHA", "Prescripcion"]
    faltan = [c for c in requeridas if c not in df.columns]
    if faltan:
        raise KeyError("Faltan columnas en la sábana tras normalizar: " + ", ".join(faltan))

    # 2) Filtrar recetas cheque AT Abierta entregadas.
    rch = df[
        (df["Tipo Receta"]     == "CONTROLADA")       &
        (df["Estado"]          == "ENTREGADA")         &
        (df["Numero Folio"].notna())                   &
        (df["Bodega Despacha"] == "FARMACIA AT ABIERTA")
    ].copy()
    print("\n  Recetas cheque AT Abierta en la sábana: " + str(len(rch)))
    if rch.empty:
        print("\n  No hay recetas cheque AT Abierta en esta sábana.")
        if not a.no_pause:
            input("\nPresiona Enter para cerrar...")
        return

    # 3) Derivar campos.
    rch[["F_COD", "NOMBRE_PROD", "PRESENTACION"]] = rch.apply(lambda r: pd.Series(get_producto(r)), axis=1)
    rch[["RUT_MED", "DV_MED"]] = rch["RUN Profesional"].apply(lambda x: pd.Series(split_rut(x)))
    rch[["RUT_PAC", "DV_PAC"]] = rch["RUN"].apply(lambda x: pd.Series(split_rut(x)))
    rch[["RUT_ADQ", "DV_ADQ"]] = rch["Run Persona Retira"].apply(lambda x: pd.Series(split_rut(x)))
    rch["RUT_QF"]        = rch["Usuario Creacion Registro"].apply(lambda x: int(str(x).strip()) if pd.notna(x) else None)
    rch["POSOLOGIA"]     = rch.apply(get_posologia, axis=1)
    rch["FECHA_PRESC_N"] = rch["Fecha Atencion"].apply(to_excel_date)
    rch["FECHA_DISP_N"]  = rch["Fecha Entrega Receta"].apply(to_excel_date)
    rch["FECHA_DISP_D"]  = rch["Fecha Entrega Receta"].apply(parse_fecha)
    rch["FOLIO_INT"]     = rch["Numero Folio"].apply(lambda x: int(x) if pd.notna(x) else None)

    # 4) Abrir formulario.
    if not os.path.exists(form_path):
        raise FileNotFoundError("No se encontró el formulario:\n  " + form_path)
    wb = load_workbook(form_path)
    if HOJA_RCH not in wb.sheetnames:
        raise KeyError(f"El formulario no tiene la hoja '{HOJA_RCH}'. Hojas: {wb.sheetnames}")
    ws = wb[HOJA_RCH]

    folios_existentes = leer_folios_existentes(ws)
    print("  Registros ya en la planilla: " + str(len(folios_existentes)))

    # 5) Filtrar por mes/año del formulario (evita mezclar meses).
    anio_f, mes_f = leer_periodo_form(ws)
    if not a.sin_filtro_mes and anio_f and mes_f:
        antes = len(rch)
        rch = rch[rch["FECHA_DISP_D"].apply(
            lambda d: bool(d) and d.year == anio_f and d.month == mes_f)].copy()
        omit = antes - len(rch)
        nombre_mes = next((k for k, v in MESES_ES.items() if v == mes_f), str(mes_f))
        print(f"  Filtro periodo formulario: {nombre_mes} {anio_f}  (omitidas {omit} de otro mes)")
    elif not a.sin_filtro_mes:
        print("  [aviso] No pude leer el periodo del formulario (B7/B8) — agrego sin filtrar por mes.")

    # 6) Quedarnos con los folios nuevos.
    nuevos = rch[~rch["FOLIO_INT"].isin(folios_existentes)].copy()
    nuevos = nuevos.drop_duplicates(subset=["FOLIO_INT"])
    print("  Registros NUEVOS a agregar: " + str(len(nuevos)))
    if nuevos.empty:
        print("\n  La planilla ya está al día.")
        if not a.no_pause:
            input("\nPresiona Enter para cerrar...")
        return

    # 7) Respaldo antes de escribir.
    if not a.no_backup:
        bak = form_path + ".bak"
        try:
            shutil.copy2(form_path, bak)
            print("  Respaldo: " + os.path.basename(bak))
        except OSError as e:
            print(f"  [aviso] No pude crear respaldo: {e}")

    # 8) Escribir filas nuevas.
    sample_font = Font(name="Arial", size=10)
    thin   = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    primera = ws.max_row + 1
    sin_fcod = 0

    for idx, (_, row) in enumerate(nuevos.iterrows()):
        er = primera + idx
        cant = int(row["Cantidad Entregada"]) if pd.notna(row["Cantidad Entregada"]) else None
        if not row["F_COD"]:
            sin_fcod += 1
        valores = [
            row["FOLIO_INT"],
            row["RUT_MED"], row["DV_MED"],
            row["RUT_PAC"], row["DV_PAC"],
            str(row["Direccion"]) if pd.notna(row["Direccion"]) else None,
            str(row["Comuna"])    if pd.notna(row["Comuna"])    else None,
            row["F_COD"],
            "=VLOOKUP(H" + str(er) + ",OCULTA!A:B,2,0)",
            row["PRESENTACION"],
            cant,
            row["POSOLOGIA"],
            row["FECHA_PRESC_N"],
            row["RUT_ADQ"], row["DV_ADQ"],
            row["FECHA_DISP_N"],
            row["RUT_QF"], None, None,
            cant,
            row["PRESENTACION"],
            1,
        ]
        for ci, val in enumerate(valores):
            cell = ws.cell(row=er, column=ci + 1)
            cell.value  = val
            cell.font   = sample_font
            cell.border = border
            if ci in (12, 15) and val is not None:
                cell.number_format = "DD/MM/YYYY"

    # 9) Guardar (manejar Excel abierto).
    try:
        wb.save(form_path)
    except PermissionError:
        print("\n  [ERROR] No pude guardar: el formulario está ABIERTO en Excel.")
        print("          Ciérralo y vuelve a ejecutar.")
        if not a.no_pause:
            input("\nPresiona Enter para cerrar...")
        return

    print("\n  ✓ Se agregaron " + str(len(nuevos)) + " registros nuevos.")
    print("  Total en planilla ahora: " + str(len(folios_existentes) + len(nuevos)))
    if sin_fcod:
        print(f"  [revisar] {sin_fcod} sin F-código (HHHA no mapeado) — completar a mano en la planilla.")
    print("\n  Registros agregados por producto:")
    for prod, cnt in nuevos["Prescripcion"].value_counts().items():
        print("    " + str(prod) + ": " + str(cnt))
    print("\n  Recuerda completar a mano: DV QF y Nombre QF.")
    print("  Planilla: " + os.path.abspath(form_path))

    if not a.no_pause:
        input("\nPresiona Enter para cerrar...")


if __name__ == "__main__":
    try:
        main()
    except (FileNotFoundError, KeyError) as e:
        print("\nERROR: " + str(e))
        if "--no-pause" not in sys.argv:
            input("\nPresiona Enter para cerrar...")
    except Exception as e:
        print("\nERROR inesperado: " + str(e))
        import traceback
        traceback.print_exc()
        if "--no-pause" not in sys.argv:
            input("\nPresiona Enter para cerrar...")
