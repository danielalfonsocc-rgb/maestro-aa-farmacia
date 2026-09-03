# -*- coding: utf-8 -*-
"""
Convierte el texto pegado de una solicitud de Gestión Territorial (el
listado suelto de pacientes que manda un establecimiento por correo, tipo
"nombre RUT establecimiento especialidad [nota]" una línea por paciente) en
la planilla FECHA/RUT/NOMBRE/ESPECIALIDAD/N° RECETA que espera
revision_solicitudes.py --xlsx.

Por qué existe: transcribir a mano cada fila de la solicitud (RUT, nombre,
especialidad, deduplicar RUT repetidos) gasta muchísimos tokens de salida
sin aportar nada — es trabajo 100% mecánico. Este script lo hace
determinísticamente a partir del RUT como ancla (patrón muy confiable), sin
necesidad de que el modelo retipee cada fila.

No intenta separar perfectamente "especialidad" de "nota/observación" en el
texto que sigue al RUT — cada establecimiento redacta distinto y una
heurística de mayúsculas/minúsculas falla tanto como acierta. En su lugar
deja el resto de la línea completo en la columna ESPECIALIDAD; el cruce en
revision_solicitudes.py compara por substring en ambos sentidos
(_elegir_receta), así que igual matchea aunque venga texto extra pegado.

Uso:
    py parsear_solicitud.py --texto solicitud.txt --estab tolten --out Solicitud.xlsx
    (o por stdin: cat solicitud.txt | py parsear_solicitud.py --estab freire --out Solicitud.xlsx)
"""
import argparse
import datetime
import re
import sys

import openpyxl

RUT_RE = re.compile(r"(\d{1,2}\.?\d{3}\.?\d{3}\s*-\s*[\dkK])")
RECETA_RE = re.compile(r"(?:receta\s*(?:completa)?\s*)?[Nn][°ºo]\.?\s*(\d{6,9})")

# Algunos establecimientos digitan el RUT sin el guion del dígito
# verificador (ej. "83444649" en vez de "8344464-9"). No es ambiguo -- el
# último carácter siempre es el DV -- pero tampoco se auto-acepta como un
# RUT normal: se calcula el DV esperado y, si calza, se deja como sugerencia
# en el aviso de "sin RUT reconocible" para que el QF solo tenga que
# confirmar, no hacer el cálculo a mano.
RUT_SIN_GUION_RE = re.compile(r"(?<!\d)(\d{7,8}[\dkK])(?!\d)")


def _normalizar_rut(rut):
    return re.sub(r"[.\s]", "", rut).upper()


def _calcular_dv(cuerpo):
    """Módulo 11 estándar del RUT chileno. Devuelve '0'-'9' o 'K'."""
    suma = 0
    multiplicador = 2
    for digito in reversed(cuerpo):
        suma += int(digito) * multiplicador
        multiplicador = multiplicador + 1 if multiplicador < 7 else 2
    resto = 11 - (suma % 11)
    if resto == 11:
        return "0"
    if resto == 10:
        return "K"
    return str(resto)


def _sugerir_rut_sin_guion(linea):
    """Si la línea trae un RUT pegado sin guion (todo dígitos, DV incluido
    como último carácter), calcula el DV esperado y arma la sugerencia
    'cuerpo-dv' solo si el cálculo confirma el DV que trae la línea --
    así no se sugiere nada cuando el número no es en realidad un RUT."""
    m = RUT_SIN_GUION_RE.search(linea)
    if not m:
        return None
    crudo = m.group(1)
    cuerpo, dv_dado = crudo[:-1], crudo[-1].upper()
    if _calcular_dv(cuerpo) != dv_dado:
        return None
    return f"{cuerpo}-{dv_dado}"


def parsear_lineas(lineas, alias_estab):
    """alias_estab: lista de palabras/frases (minúsculas) que identifican al
    establecimiento en el texto, para quitarlas del inicio del resto de la
    línea (ej. ["tolten", "hospital tolten"])."""
    filas = []
    sin_rut = []
    for cruda in lineas:
        linea = cruda.strip()
        if not linea:
            continue
        m = RUT_RE.search(linea)
        if not m:
            sin_rut.append(cruda)
            continue
        nombre = linea[: m.start()].strip(" \t-—")
        resto = linea[m.end():].strip()
        rut = re.sub(r"\s", "", m.group(1))

        # saca el alias del establecimiento si arranca con él
        resto_lower = resto.lower()
        for alias in sorted(alias_estab, key=len, reverse=True):
            if resto_lower.startswith(alias):
                resto = resto[len(alias):].strip(" \t-—")
                break

        # extrae un N° de receta sugerido si aparece en cualquier parte
        n_receta = ""
        rm = RECETA_RE.search(resto)
        if rm:
            n_receta = rm.group(1)
            resto = (resto[: rm.start()] + resto[rm.end():]).strip(" \t-—.")

        especialidad = re.sub(r"\s+", " ", resto).strip()
        filas.append({
            "rut": rut, "rut_norm": _normalizar_rut(rut),
            "nombre": nombre, "especialidad": especialidad, "n_receta": n_receta,
        })
    return filas, sin_rut


def deduplicar(filas):
    """RUT repetido en la solicitud (pasa seguido: el establecimiento manda
    la misma fila dos veces, o la repite agregando el N° de receta que le
    faltaba la primera vez) -> se fusiona en una sola fila en vez de mandar
    duplicados a revision_solicitudes.py."""
    por_rut = {}
    orden = []
    for f in filas:
        clave = f["rut_norm"]
        if clave not in por_rut:
            por_rut[clave] = dict(f)
            orden.append(clave)
            continue
        existente = por_rut[clave]
        if f["n_receta"] and not existente["n_receta"]:
            existente["n_receta"] = f["n_receta"]
        if f["especialidad"] and f["especialidad"] not in existente["especialidad"]:
            extra = f["especialidad"]
            existente["especialidad"] = (existente["especialidad"] + " | " + extra).strip(" |")
    return [por_rut[c] for c in orden]


def escribir_xlsx(filas, out_path, fecha):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Solicitud"
    ws.append(["FECHA", "RUT", "NOMBRE", "ESPECIALIDAD", "N° RECETA"])
    for f in filas:
        ws.append([fecha, f["rut"], f["nombre"], f["especialidad"], f["n_receta"]])
    anchos = [12, 14, 32, 40, 12]
    for col, w in zip("ABCDE", anchos):
        ws.column_dimensions[col].width = w
    wb.save(out_path)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--texto", help="Archivo .txt con el listado pegado (una línea por paciente). Si se omite, lee stdin.")
    ap.add_argument("--estab", required=True,
                     help="Alias del establecimiento a quitar del inicio de cada línea (ej. 'tolten' o 'tolten,hospital tolten'). Separar varios con coma.")
    ap.add_argument("--out", required=True, help="Ruta del xlsx de salida")
    ap.add_argument("--fecha", help="dd-mm-aaaa (default: hoy)")
    args = ap.parse_args()

    texto = open(args.texto, encoding="utf-8").read() if args.texto else sys.stdin.read()
    alias_estab = [a.strip().lower() for a in args.estab.split(",") if a.strip()]
    fecha = args.fecha or datetime.date.today().strftime("%d-%m-%Y")

    filas, sin_rut = parsear_lineas(texto.splitlines(), alias_estab)
    if not filas:
        print("[ERROR] No encontré ninguna línea con formato RUT reconocible.")
        sys.exit(1)

    unicas = deduplicar(filas)
    escribir_xlsx(unicas, args.out, fecha)

    print(f"{len(filas)} línea(s) con RUT -> {len(unicas)} paciente(s) único(s). Guardado: {args.out}")
    if sin_rut:
        print(f"[AVISO] {len(sin_rut)} línea(s) NO tenían un RUT reconocible y se omitieron — revísalas a mano:")
        for l in sin_rut:
            sugerencia = _sugerir_rut_sin_guion(l)
            if sugerencia:
                print(f"    {l.strip()}  [posible RUT sin guion, DV verificado: {sugerencia}]")
            else:
                print(f"    {l.strip()}")


if __name__ == "__main__":
    main()
