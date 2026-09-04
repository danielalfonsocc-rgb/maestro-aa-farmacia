#!/usr/bin/env python3
"""
gt_pendientes_maestro_pacientes.py — Cruza la hoja Despachos_<Mes> del Sheet
paciente-céntrico (migrar_gestion_territorial.py / Codigo.gs) contra el
histórico informe_completo_recetas*.csv, y auto-completa las recetas que
tienen algún fármaco con Cantidad Pendiente > 0 (mismo criterio que
cruce_gt.py: la columna real del CSV, no el campo "Estado", que es menos
confiable — ver memoria gt-nomina-logica-correcta).

Rellena SOLO celdas vacías o previamente auto-escritas por este mismo script:
  - Columna "Fármaco Pendiente / Stock": "[SSASUR-AUTO] Producto xCantidad; ..."
  - Columna "Estado Despacho": "pendiente por falta stock" (solo si estaba vacía)

Blindajes aplicados (mismos patrones de bug ya vistos en gt_maestro.py/
cruce_gt.py — ver memorias gt-inyectables-lai-deposito-bug,
stock-huerfano-homologacion-forzar, gt-nomina-logica-correcta):

  1. Columnas ubicadas por NOMBRE de encabezado (fila 3), no por índice fijo.
     Un cambio futuro en el orden de columnas de migrar_gestion_territorial.py
     rompe esto con un error claro, no con un cruce silenciosamente mal hecho
     contra la columna equivocada (el bug real de _TIPO_GT_RE con los
     inyectables fue exactamente este patrón: un regex/posición fija dejó de
     matchear un formato real y la data se perdió sin aviso).
  2. Nº de receta normalizado en ambos lados (recorta espacios, separadores de
     miles, ".0" de floats) antes de cruzar — evita falsos "no encontrado"
     por diferencia de formato.
  3. Marcador "[SSASUR-AUTO]" en lo que este script escribe en Fármaco
     Pendiente: en la próxima corrida, una celda que empieza con ese marcador
     se puede refrescar (evita que quede mostrando información obsoleta para
     siempre, el mismo patrón del bug de "stock fantasma" — dato reportado
     que ya no es cierto pero nadie lo actualiza); una celda con texto de un
     QF (sin el marcador) nunca se toca.
  4. Alerta explícita si la tasa de cruce contra el histórico es sospechosamente
     baja (<5% de las recetas con datos encontradas) — señal de que el CSV
     cambió de formato o el glob no encontró los archivos correctos, en vez de
     reportar silenciosamente "0 pendientes" como si todo estuviera bien.
  5. Nunca sobreescribe Estado Despacho si ya tiene algún valor — es un estado
     de trabajo asignado por el QF, no un dato derivable con certeza del CSV.
  6. Escritura en Sheets en lotes (chunks) con manejo de error explícito: si
     falla a medio camino, se informa exactamente qué se alcanzó a escribir,
     no un fallo silencioso a medias.

Uso:
  py gt_pendientes_maestro_pacientes.py --sheet-id <ID>              # dry-run
  py gt_pendientes_maestro_pacientes.py --sheet-id <ID> --aplicar    # escribe
  py gt_pendientes_maestro_pacientes.py --sheet-id <ID> --hoja Despachos_Octubre --aplicar
"""
import argparse
import glob
import os
import re
import sys

MAESTRO_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, MAESTRO_DIR)

for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

SHEET_ID_PRUEBA = "1p7eFcKO1xS2WJRD5KlGKRPo5Wb0Cm-lBT4lX5dwPt8c"
TOKEN_FILE = os.path.join(MAESTRO_DIR, "token_drive.json")
SCOPES = ["https://www.googleapis.com/auth/drive"]
MARCADOR_AUTO = "[SSASUR-AUTO] "
ESTADO_PENDIENTE = "pendiente por falta stock"

# Encabezados esperados en la fila 3 de Despachos_<Mes>/Plantilla_Mes_Nuevo
# (ver migrar_gestion_territorial.py). Coincidencia tolerante (sin tildes/
# mayúsculas) por palabras clave, no el string exacto completo — para no
# romperse por un cambio menor de redacción del encabezado.
COL_RECETA_KW = ["receta"]
COL_ESTADO_KW = ["estado", "despacho"]
COL_FARMACO_KW = ["farmaco", "pendiente"]
CHUNK = 400


def _key(s):
    import unicodedata
    s = unicodedata.normalize("NFKD", str(s or "")).encode("ascii", "ignore").decode()
    return re.sub(r"[^a-z0-9]", "", s.lower())


def _normalizar_receta(v):
    """Recorta espacios, separadores de miles y ".0" de floats — para que un
    Nº de receta represente lo mismo venga del CSV (texto) o del Sheet
    (a veces número). Sin esto, un formato distinto entre ambos lados produce
    falsos "no encontrado en histórico" sin que se note por qué."""
    s = str(v or "").strip().replace(",", "")
    if s.endswith(".0"):
        s = s[:-2]
    return s


def _localizar_columna(header_row, keywords, nombre_humano):
    claves = [_key(h) for h in header_row]
    for i, k in enumerate(claves):
        if all(kw in k for kw in keywords):
            return i
    raise SystemExit(
        f"No se encontró la columna '{nombre_humano}' en el encabezado (fila 3) de la hoja. "
        f"Encabezado leído: {header_row}\n"
        f"Esto probablemente significa que cambió el formato de migrar_gestion_territorial.py "
        f"— revisar antes de seguir, no asumir una posición fija."
    )


def _conectar():
    from google.oauth2.credentials import Credentials
    from google.auth.transport.requests import Request
    import gspread

    creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    if not creds.valid:
        if creds.expired and creds.refresh_token:
            creds.refresh(Request())
            with open(TOKEN_FILE, "w", encoding="utf-8") as fh:
                fh.write(creds.to_json())
        else:
            raise SystemExit("Token invalido - re-autorizar con publicar_drive.py --setup")
    return gspread.authorize(creds)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--sheet-id", default=SHEET_ID_PRUEBA,
                     help="ID del Google Sheet (default: Sheet de PRUEBA)")
    ap.add_argument("--hoja", default=None,
                     help="Hoja a procesar (default: la que diga Parametros!G2, 'Mes Activo')")
    ap.add_argument("--hist-glob", default=os.path.join(MAESTRO_DIR, "informe_completo_recetas*.csv"))
    ap.add_argument("--aplicar", action="store_true",
                     help="Escribe los cambios en el Sheet. Sin este flag solo se muestra qué haría (dry-run).")
    a = ap.parse_args()

    import cruce_gt as CG

    gc = _conectar()
    sh = gc.open_by_key(a.sheet_id)

    hoja = a.hoja
    if not hoja:
        ws_param = sh.worksheet("Parametros")
        hoja = (ws_param.acell("G2").value or "").strip()
        if not hoja:
            raise SystemExit("Parametros!G2 (Mes Activo) esta vacio y no se paso --hoja explicito.")
    print(f"Hoja objetivo: {hoja}")

    ws = sh.worksheet(hoja)

    header_row = ws.get("A3:R3", value_render_option="FORMATTED_VALUE")
    if not header_row or not header_row[0]:
        raise SystemExit(f"No se pudo leer el encabezado (fila 3) de '{hoja}'.")
    header_row = header_row[0]
    idx_receta = _localizar_columna(header_row, COL_RECETA_KW, "N° Receta")
    idx_estado = _localizar_columna(header_row, COL_ESTADO_KW, "Estado Despacho")
    idx_farmaco = _localizar_columna(header_row, COL_FARMACO_KW, "Fármaco Pendiente / Stock")
    print(f"Columnas localizadas por encabezado: Receta={chr(65+idx_receta)}, "
          f"Estado={chr(65+idx_estado)}, FarmacoPendiente={chr(65+idx_farmaco)}")

    last_row = ws.row_count
    ancho = max(idx_receta, idx_estado, idx_farmaco) + 1
    ultima_col_letra = chr(65 + ancho - 1) if ancho <= 26 else "Z"
    filas = ws.get(f"A4:{ultima_col_letra}{last_row}", value_render_option="FORMATTED_VALUE")
    print(f"Filas leidas: {len(filas)} (A4:{ultima_col_letra}{last_row})")

    def celda(fila, idx):
        return (fila[idx] if len(fila) > idx else "").strip()

    recetas_set = set()
    for fila in filas:
        rec = _normalizar_receta(celda(fila, idx_receta))
        if rec:
            recetas_set.add(rec)

    archivos = sorted(glob.glob(a.hist_glob))
    if not archivos:
        raise SystemExit(f"No se encontraron archivos con el patron {a.hist_glob}")
    print(f"Historico: {len(archivos)} archivo(s) — cruzando por N° receta...")

    recetas_set_norm = {_normalizar_receta(r) for r in recetas_set}
    det_raw = CG.cruzar_historico(recetas_set_norm, archivos)
    det = {_normalizar_receta(k): v for k, v in det_raw.items()}

    cambios_n = []  # [(row_num, texto)]
    cambios_k = []
    recetas_pendientes = []
    recetas_sin_historico = 0
    filas_con_receta = 0

    for i, fila in enumerate(filas):
        row_num = 4 + i
        rec = _normalizar_receta(celda(fila, idx_receta))
        if not rec:
            continue
        filas_con_receta += 1
        estado_actual = celda(fila, idx_estado)
        farmaco_actual = celda(fila, idx_farmaco)

        d = det.get(rec)
        if not d or not d["lineas"]:
            recetas_sin_historico += 1
            continue

        pend = []
        for ln in d["lineas"].values():
            if ln["pendiente"] > 0 and ln["prod"]:
                pend.append(f"{ln['prod'].title()} x{ln['pendiente']}")
        pend = list(dict.fromkeys(pend))  # dedup preservando orden
        if not pend:
            continue

        pend_str = MARCADOR_AUTO + "; ".join(pend)
        recetas_pendientes.append(rec)

        # Solo se toca N si está vacía O si es una escritura previa de este
        # mismo script (empieza con el marcador) — así se puede refrescar sin
        # riesgo de pisar una nota manual del QF.
        if not farmaco_actual or farmaco_actual.startswith(MARCADOR_AUTO.strip()):
            if farmaco_actual != pend_str:
                cambios_n.append((row_num, pend_str))
        if not estado_actual:
            cambios_k.append((row_num, ESTADO_PENDIENTE))

    print(f"\nFilas con N° de receta: {filas_con_receta}")
    print(f"Recetas con algun farmaco pendiente (Cantidad Pendiente > 0): {len(recetas_pendientes)}")
    print(f"Recetas no halladas en el historico (sin cruce, no tocadas): {recetas_sin_historico}")
    print(f"  -> Columna Farmaco Pendiente/Stock a escribir/refrescar: {len(cambios_n)}")
    print(f"  -> Columna Estado Despacho a rellenar (estaba vacia): {len(cambios_k)}")

    omitidas_k = len(recetas_pendientes) - len(cambios_k)
    if omitidas_k:
        print(f"  (Filas con pendiente detectado pero Estado Despacho YA tenia algo escrito, no se toca: {omitidas_k})")

    # Blindaje: tasa de cruce sospechosamente baja -> probable rotura de
    # formato (CSV cambió de columnas, o el glob no encontró los archivos
    # correctos), no "no hay pendientes de verdad".
    if filas_con_receta >= 20:
        tasa_cruce = (filas_con_receta - recetas_sin_historico) / filas_con_receta
        if tasa_cruce < 0.05:
            print(f"\n[ALERTA] Solo {tasa_cruce:.0%} de las recetas con dato en la hoja se encontraron en el "
                  f"histórico. Con {filas_con_receta} recetas y {len(archivos)} archivo(s) CSV, eso es "
                  f"sospechosamente bajo — revisar antes de confiar en el resultado (posible cambio de "
                  f"formato en informe_completo_recetas*.csv, o los archivos no cubren el período de esta hoja).")

    if not a.aplicar:
        print("\n[DRY-RUN] No se escribio nada. Corre de nuevo con --aplicar para confirmar los cambios de arriba.")
        return

    if not cambios_n and not cambios_k:
        print("\nNada que aplicar.")
        return

    col_farmaco_letra = chr(65 + idx_farmaco)
    col_estado_letra = chr(65 + idx_estado)
    requests = []
    for row_num, texto in cambios_n:
        requests.append({"range": f"{col_farmaco_letra}{row_num}", "values": [[texto]]})
    for row_num, texto in cambios_k:
        requests.append({"range": f"{col_estado_letra}{row_num}", "values": [[texto]]})

    escritas = 0
    try:
        for i in range(0, len(requests), CHUNK):
            lote = requests[i:i + CHUNK]
            ws.batch_update(lote, value_input_option="USER_ENTERED")
            escritas += len(lote)
    except Exception as e:
        print(f"\n[ERROR] Falló la escritura después de {escritas}/{len(requests)} celdas: {e}")
        print("        Las celdas ya escritas quedaron aplicadas; corre de nuevo para completar el resto "
              "(es seguro — no vuelve a tocar lo que ya quedó bien, salvo refresco de Fármaco Pendiente).")
        raise

    print(f"\n[APLICADO] {len(cambios_n)} celda(s) de Fármaco Pendiente y {len(cambios_k)} celda(s) de "
          f"Estado Despacho escritas en '{hoja}' ({escritas} celdas en total).")


if __name__ == "__main__":
    main()
