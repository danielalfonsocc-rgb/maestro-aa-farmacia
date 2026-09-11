#!/usr/bin/env python3
"""
cruce_gt.py — Cruza el reporte de Gestión Territorial (modalidad de despacho) con
el histórico de recetas (informe_completo_recetas*.csv) para clasificar, por
receta, sus medicamentos en:

  · REFRIGERADOS  (insulinas / cadena de frío)        + cantidad recetada
  · CONTROLADOS   (Tipo Receta == CONTROLADA en SSASUR) + cantidad recetada
  · PENDIENTES    (Cantidad Pendiente > 0)             + cantidad pendiente

Join por "Nº Receta" (GT) == "Número Receta" (histórico). Deduplica las líneas de
prescripción por "ID Receta Detalle". Produce:
  · resumen por consola
  · <salida>/gt_enriquecido.json   (formato registros del skill gestion-territorial)
  · <salida>/Cruce_GT_Clasificacion.xlsx  (3 categorías con cantidad, por destino)

Uso:
  py cruce_gt.py <reporteGT.xlsx> --salida ./out_gt [--hist-glob "informe_completo_recetas*.csv"]
"""
import argparse, csv, glob, json, os, re, sys, tempfile, unicodedata
from collections import OrderedDict, defaultdict

for _s in (sys.stdout, sys.stderr):
    try: _s.reconfigure(encoding="utf-8", errors="replace")
    except Exception: pass

MAESTRO_DIR = os.path.dirname(os.path.abspath(__file__))

# ── Listas de clasificación ────────────────────────────────────────────────────
TIPOS_INSULINA = [("GLARGINA","Glargina"),("ASPART","Asparta"),("GLULISINA","Glulisina"),
                  ("CRISTALINA","Cristalina"),("LISPRO","Lispro"),("DEGLUDEC","Degludec"),
                  ("DETEMIR","Detemir"),("NPH","NPH")]
# Otros refrigerados frecuentes (cadena de frío) además de insulinas:
REFRIG_OTROS = ["ANALOGO GLP", "DULAGLUTIDA", "SEMAGLUTIDA", "LIRAGLUTIDA",
                "ENOXAPARINA", "VACUNA", "TOXINA BOTULINICA", "OCTREOTIDA",
                "FILGRASTIM", "ERITROPOYETINA", "EPOETINA", "SOMATROPINA",
                "TERIPARATIDA", "INTERFERON"]
# Controlados — INN extraídos de reporte_de_stock_20260618093729.xlsx (columna Descripción).
# Fuente: stock filtrado a psicotrópicos/opioides de Farmacia Hospital de Pitrufquén (18-06-2026).
# SENSIDISCO ERTAPENEM excluido (disco antibiótico, no psicotrópico).
CONTROL_INN = {
    "ALPRAZOLAM","BUPRENORFINA","CLOBAZAM","CLONAZEPAM","CODEINA","DIAZEPAM",
    "FENOBARBITAL","FENTANILO","KETAMINA","LISDEXANFETAMINA","LORAZEPAM",
    "METADONA","METILFENIDATO","MIDAZOLAM","MORFINA","OXICODONA",
    "PETIDINA","REMIFENTANILO","TAPENTADOL","ZOLPIDEM",
}

# Alias de "Estab. Destino" — SSASUR a veces exporta el mismo establecimiento
# con un string distinto al canónico (confirmado 25-08-2026: la receta
# 44515366 de María Ernestina Huircan Pichun vino con "SUR TEODORO SCHMIDT"
# en el reporte de modalidad de despacho, no "CESFAM TEODORO SCHMIDT" — no
# fue un typo de digitación manual, es lo que SSASUR mismo exportó). Sin
# esto, generar.py agrupa por el string crudo y genera una Nómina de Envío
# aparte de 1 paciente para el alias — mismo bug que agregar_gt_manual.py
# tiene con _CARPETA_LOCAL, pero acá en el pipeline automático. Se normaliza
# ACÁ (leer_reporte_gt), antes de que el destino se use para agrupar,
# clasificar o generar planillas — así corrige TODO el pipeline de una vez.
ALIAS_DESTINO = {
    "SUR TEODORO SCHMIDT": "CESFAM TEODORO SCHMIDT",
}

# Estados de "Estado Prescripción" en los que la línea ya NO se va a despachar:
# no están ENTREGADO, pero tampoco son un pendiente que la farmacia deba
# preparar. Ver clasificar().
_ESTADOS_TERMINALES = {"ANULADO", "REEMPLAZADO", "DEVUELTO", "RECHAZADO"}


def _normalizar_destino(d):
    d = str(d or "").strip()
    return ALIAS_DESTINO.get(d.upper(), d)


def _key(h):
    h = unicodedata.normalize("NFKD", str(h or "")).encode("ascii","ignore").decode().lower()
    return re.sub(r"[^a-z0-9]", "", h)


def insulina_label(prod):
    u = str(prod or "").upper()
    for kw, lab in TIPOS_INSULINA:
        if kw in u: return f"Insulina {lab}"
    if "INSULINA" in u and "JERINGA" not in u and "AGUJA" not in u:
        return "Insulina"
    return None


def es_refrigerado(prod):
    lab = insulina_label(prod)
    if lab: return lab
    u = str(prod or "").upper()
    for kw in REFRIG_OTROS:
        if kw in u:
            return prod.strip().title()
    return None


def es_controlado_oficial(prod):
    u = str(prod or "").upper()
    for inn in CONTROL_INN:
        if inn in u: return True
    return False


def es_inyectable_lai(prod):
    """Antipsicóticos de depósito (LAI) de salud mental ambulatoria — mismo
    catálogo que centinela_inyectables_sm.py. Requiere el calificador de forma
    de depósito (DECANOATO/JER/FA) para no confundir con la forma oral del
    mismo principio activo (ej. 'HALOPERIDOL 5 MG COMPRIMIDO' o
    'RISPERIDONA 1 MG COMPRIMIDO' NO deben quedar marcados)."""
    u = str(prod or "").upper()
    if "HALOPERIDOL" in u and "DECANOATO" in u: return True
    if "ZUCLOPENTIXOL" in u and "DECANOATO" in u: return True
    if "PALIPERIDONA" in u and ("PALMITATO" in u or "JER" in u): return True
    if "RISPERIDONA" in u and (re.search(r"\bFA\b", u) or "MICROESFERAS" in u or "CONSTA" in u): return True
    return False


def _num(v):
    try: return int(float(str(v).replace(",", ".").strip() or 0))
    except Exception: return 0


# ── Lectura del reporte GT (.xlsx) ─────────────────────────────────────────────
def leer_reporte_gt(path):
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = [r for r in ws.iter_rows(values_only=True)]
    wb.close()
    # fila de encabezado = primera con >=5 celdas no vacías
    hi = next((i for i, r in enumerate(rows) if sum(1 for c in r if c not in (None, "")) >= 5), None)
    if hi is None:
        raise ValueError(f"No se encontró fila de encabezado válida en {path} (se esperaban >=5 columnas con datos)")
    hdr = [str(c).strip() if c is not None else "" for c in rows[hi]]
    K = {_key(h): i for i, h in enumerate(hdr)}

    def col(*keys, default=None):
        for k in keys:
            if k in K: return K[k]
        return default

    c_rec  = col("nreceta", "noreceta", "numeroreceta")
    c_pac  = col("paciente")
    c_run  = col("runpaciente", "run", "rut")
    c_edad = col("edad")
    c_dir  = col("direccion")
    c_com  = col("comuna")
    c_tel  = col("telefono", "fono")
    c_org  = col("estaborigen")
    c_dst  = col("estabdestino")
    c_fen  = col("fechaentrega")
    c_per  = col("periodoreceta", "periodo")
    c_esp  = col("especialidad")
    c_np   = col("numeroprescripciones", "nprescripciones")
    c_tret = col("tiporetiro")

    requeridas = {"numeroreceta": c_rec, "paciente": c_pac, "run": c_run, "estabdestino": c_dst, "periodo": c_per}
    faltantes = [k for k, v in requeridas.items() if v is None]
    if faltantes:
        print(f"  [aviso] columnas no encontradas en el reporte GT: {faltantes} — revisar encabezados de {path}")
    if c_rec is None:
        raise ValueError(f"No se encontró columna de Nº Receta en {path} (se esperaba nreceta/noreceta/numeroreceta)")

    def g(r, i):
        return str(r[i]).strip() if (i is not None and i < len(r) and r[i] is not None) else ""

    regs = OrderedDict()
    for r in rows[hi+1:]:
        if not r:
            continue
        rec = g(r, c_rec)
        if not rec or rec.upper() in ("TOTAL", "NONE"):
            continue
        if rec not in regs:
            edad = re.sub(r"\D", "", g(r, c_edad))
            regs[rec] = {
                "receta": rec, "paciente": g(r, c_pac), "run": g(r, c_run),
                "edad": int(edad) if edad else None, "direccion": g(r, c_dir),
                "comuna": g(r, c_com), "telefono": g(r, c_tel),
                "estab_origen": g(r, c_org) or "Pitrufquén Hosp.",
                "estab_destino": _normalizar_destino(g(r, c_dst)), "fecha_entrega_rep": g(r, c_fen),
                "periodo": g(r, c_per), "especialidad": g(r, c_esp),
                "n_presc": _num(g(r, c_np)),
                "ventanilla": g(r, c_tret).upper() == "PACIENTE",
            }
    return regs, hdr


# ── Cruce con el histórico ─────────────────────────────────────────────────────
def cruzar_historico(recetas_set, archivos):
    """Devuelve {n_receta: {"tipo_receta","lineas":[{id,prod,recetada,pendiente,estado_presc}]}}.
    Deduplica líneas por ID Receta Detalle."""
    det = defaultdict(lambda: {"tipo_receta": "", "lineas": OrderedDict()})
    vistos = set()
    for fp in archivos:
        try:
            f = open(fp, encoding="latin-1", newline="")
        except Exception as e:
            print(f"  [aviso] no pude abrir {os.path.basename(fp)}: {e}")
            continue
        with f:
            rd = csv.reader(f, delimiter=";")
            try:
                hdr = next(rd)
            except StopIteration:
                continue
            K = {_key(h): i for i, h in enumerate(hdr)}
            ix_rec = K.get("numeroreceta"); ix_tr = K.get("tiporeceta")
            ix_id = K.get("idrecetadetalle")
            ix_pre = K.get("prescripcion")
            ix_cr = K.get("cantidadrecetada"); ix_cp = K.get("cantidadpendiente")
            ix_ep = K.get("estadoprescripcion")
            if ix_rec is None or ix_pre is None:
                print(f"  [aviso] {os.path.basename(fp)} sin columnas esperadas — omito")
                continue
            n = len(hdr)
            for row in rd:
                if len(row) < n:
                    continue
                rec = (row[ix_rec] or "").strip()
                if rec not in recetas_set:
                    continue
                idd = (row[ix_id] or "").strip() if ix_id is not None else ""
                vk = idd or f"{rec}|{row[ix_pre]}"
                if vk in vistos:
                    continue
                vistos.add(vk)
                d = det[rec]
                if not d["tipo_receta"] and ix_tr is not None: d["tipo_receta"] = (row[ix_tr] or "").strip()
                d["lineas"][vk] = {
                    "prod": (row[ix_pre] or "").strip(),
                    "recetada": _num(row[ix_cr]) if ix_cr is not None else 0,
                    "pendiente": _num(row[ix_cp]) if ix_cp is not None else 0,
                    "estado_presc": (row[ix_ep] or "").strip().upper() if ix_ep is not None else "",
                }
    return det


def clasificar(reg, d):
    """Rellena refrigerado/controlado/pendiente/inyectable_lai (texto con cantidad) en reg."""
    refri, control, pend, iny = [], [], [], []
    tipo_controlada = (d["tipo_receta"] or "").upper() == "CONTROLADA"
    for ln in d["lineas"].values():
        prod = ln["prod"]
        if not prod:
            continue
        rlab = es_refrigerado(prod)
        if rlab:
            refri.append(f"{rlab} x{ln['recetada']}")
        if es_inyectable_lai(prod):
            iny.append(f"{prod.title()} x{ln['recetada']}")
        # controlado: receta CONTROLADA y producto que matchea lista, o (si ninguno
        # matchea) todos los fármacos de una receta CONTROLADA
        if tipo_controlada and es_controlado_oficial(prod):
            control.append(f"{prod.title()} x{ln['recetada']}")
        # pendiente: CUALQUIERA de las 2 señales (no solo Cantidad Pendiente).
        # Bug real 02-09-2026: una receta recién digitada trae Estado
        # Prescripción='SOLICITADO' (correcto, ya está pendiente) con
        # Cantidad Pendiente todavía en 0 porque ese campo lo actualiza un
        # proceso de SSASUR con retraso frente al estado real — la Planilla
        # de ese día salía con "Pendiente" en blanco pese a que el medicamento
        # genuinamente no se había entregado (caso real: Aripiprazol 15mg,
        # receta 48212191, corregido a mano — ver [[gt-manual-vs-pipeline-auto]]).
        # El resto del proyecto (agente_duplicados.py, auditoria_prescripcion.py,
        # centinela_reporte.py, etc.) ya usa Estado Prescripción=='ENTREGADO'
        # como criterio de "ya se despachó" — cruce_gt.py era el único que no
        # lo miraba. PERO Estado Prescripción tampoco es 100% confiable solo:
        # auditoría 07-09-2026 sobre 560 líneas GT encontró 5 con Estado
        # Prescripción='ENTREGADO' y Cantidad Entregada=0/Cantidad
        # Pendiente=recetada completa (inconsistencia real de SSASUR, ej.
        # Metformina 1000mg LM x3 recetas) — ahí Cantidad Pendiente es la
        # señal correcta y Estado Prescripción miente. Por eso se usa OR:
        # cualquiera de las 2 en "no entregado" marca la línea como
        # pendiente — prioriza no perder un medicamento realmente faltante
        # sobre el riesgo menor de marcar de más. Cantidad mostrada: la
        # pendiente si ya está > 0, si no la recetada completa.
        # ...salvo los estados TERMINALES: una línea ANULADO/REEMPLAZADO/
        # DEVUELTO/RECHAZADO tampoco está "ENTREGADO", pero ya no se va a
        # despachar nunca — la prescripción murió (reemplazada por otra,
        # anulada por el prescriptor, devuelta o rechazada). Contarlas como
        # pendiente le pone a la nómina medicamentos que el establecimiento
        # no debe esperar: son 1.428 líneas del histórico (512 ANULADO, 505
        # REEMPLAZADO, 371 DEVUELTO, 40 RECHAZADO al 11-09-2026). Caso real:
        # la receta 46853557 de CESFAM TEODORO SCHMIDT tenía todo entregado
        # salvo una Atorvastatina 80 REEMPLAZADO, y salía marcada como
        # pendiente completa. Cantidad Pendiente > 0 tampoco los rescata: en
        # una línea anulada ese campo queda con la cantidad recetada.
        estado_presc = ln.get("estado_presc", "")
        if estado_presc in _ESTADOS_TERMINALES:
            continue
        no_entregado_por_estado = bool(estado_presc) and estado_presc != "ENTREGADO"
        if no_entregado_por_estado or ln["pendiente"] > 0:
            cant = ln["pendiente"] if ln["pendiente"] > 0 else ln["recetada"]
            pend.append(f"{prod.title()} x{cant}")
    if tipo_controlada and not control:
        # receta marcada CONTROLADA pero sin match de palabra clave → listar fármacos no insulina
        for ln in d["lineas"].values():
            if ln["prod"] and not es_refrigerado(ln["prod"]):
                control.append(f"{ln['prod'].title()} x{ln['recetada']}")
    reg["refrigerado"] = "; ".join(dict.fromkeys(refri))
    reg["controlado"]  = "; ".join(dict.fromkeys(control))
    reg["pendiente"]   = "; ".join(dict.fromkeys(pend))
    reg["inyectable_lai"] = "; ".join(dict.fromkeys(iny))
    reg["tipo_receta"] = d["tipo_receta"]
    reg["_en_historico"] = bool(d["lineas"])
    return reg


# ── Salidas ────────────────────────────────────────────────────────────────────
def escribir_excel(regs, ruta):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    thin = Side(style="thin", color="BFBFBF"); bd = Border(thin, thin, thin, thin)
    wb = Workbook(); ws = wb.active; ws.title = "Cruce GT"
    cols = ["Estab. Destino","Nº Receta","Paciente","RUN","Especialidad","Período",
            "Tipo Receta","Refrigerados (cant.)","Controlados (cant.)","Pendientes (cant.)",
            "Inyectables LAI (cant.)","En histórico"]
    widths = [22,11,30,14,24,8,14,34,34,34,34,11]
    for c, w in enumerate(widths, 1): ws.column_dimensions[get_column_letter(c)].width = w
    for c, h in enumerate(cols, 1):
        cell = ws.cell(1, c, h); cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="2E5496"); cell.border = bd
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 30
    r = 2
    for g in sorted(regs.values(), key=lambda x: (x["estab_destino"], x["paciente"])):
        vals = [g["estab_destino"], g["receta"], g["paciente"], g["run"], g["especialidad"],
                g["periodo"], g.get("tipo_receta",""), g.get("refrigerado",""),
                g.get("controlado",""), g.get("pendiente",""), g.get("inyectable_lai",""),
                "Sí" if g.get("_en_historico") else "NO"]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(r, c, v); cell.border = bd
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            cell.font = Font(size=9)
            if c == 8 and v: cell.font = Font(size=9, bold=True, color="1F6F3D")   # refrig verde
            if c == 9 and v: cell.font = Font(size=9, bold=True, color="C00000")   # control rojo
            if c == 10 and v: cell.font = Font(size=9, bold=True, color="C55A11")  # pend naranjo
            if c == 11 and v: cell.font = Font(size=9, bold=True, color="7030A0")  # inyectable LAI morado
        if not g.get("_en_historico"):
            for c in range(1, len(cols)+1): ws.cell(r, c).fill = PatternFill("solid", fgColor="FFF2CC")
        r += 1
    ws.freeze_panes = "A2"; ws.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{r-1}"
    wb.save(ruta)


def _buscar_generar_py():
    """Busca generar.py del skill gestion-territorial-pitrufquen.
    Prioridad: copia local skill_gt/ > AppData (solo accesible fuera del sandbox)."""
    # 1) Copia local incluida en el repositorio maestro (siempre accesible)
    local = os.path.join(MAESTRO_DIR, "skill_gt", "scripts", "generar.py")
    if os.path.exists(local):
        return local
    return None


# Carpeta hermana del repo con el árbol de Nóminas de Envío ya depositadas
# (ver CLAUDE.md → "GT raw downloads").
GT_DIR = os.path.join(os.path.dirname(MAESTRO_DIR), "04_Farmacia_Gestion_Territorial")
NOMINAS_CACHE = os.path.join(MAESTRO_DIR, "_gt_nominas_generadas.json")

# Estados de gt_maestro.xlsx que SÍ prueban que la receta ya salió en una
# nómina. "EN REVISIÓN" y "EN PREPARACIÓN" NO prueban nada: son los estados
# con que se registra una receta recién vista en el reporte de despacho
# (_sincronizar_maestro escribe "EN PREPARACIÓN" para TODAS las recetas del
# reporte crudo, incluidas las que ese día no generaron planilla).
_ESTADOS_YA_DESPACHADA = {"listapararetiro", "retiroenventanilla", "enviada", "entregada"}


def _recetas_de_planilla(ruta):
    """Nº de receta de una Nómina de Envío ya generada (hoja "Funcionarios",
    encabezado en la fila 5, Nº Receta en la columna B — ver
    generar.hoja_funcionarios y agregar_gt_manual._leer_nomina_existente)."""
    from openpyxl import load_workbook
    try:
        wb = load_workbook(ruta, read_only=True, data_only=True)
    except Exception:
        return []
    try:
        if "Funcionarios" not in wb.sheetnames:
            return []
        ws = wb["Funcionarios"]
        out = []
        for row in ws.iter_rows(min_row=6, min_col=2, max_col=2, values_only=True):
            v = row[0]
            if v is None:
                continue
            v = str(v).strip()
            if v and v.upper() != "TOTAL":
                out.append(v)
        return out
    finally:
        wb.close()


def _recetas_con_nomina():
    """Nº de receta que YA aparecen en alguna Nómina de Envío realmente
    generada — la única prueba dura de "a esta receta ya se le hizo planilla".
    Busca en out_gt/<rango>/ y en el árbol depositado
    04_Farmacia_Gestion_Territorial/<ESTAB>/Nóminas de Envío/.

    Cachea por archivo (ruta → mtime/tamaño/recetas) en NOMINAS_CACHE para no
    releer ~330 xlsx en cada corrida; solo se vuelve a leer el archivo nuevo o
    modificado."""
    # "Nomina_Manual_*" es el nombre legado de agregar_gt_manual.py (ya no lo
    # produce — ver CLAUDE.md, cambio del 07-09-2026), pero quedan decenas en
    # el árbol y son nóminas reales: si no se cuentan, el pipeline volvería a
    # generar planilla para recetas que ya se enviaron a mano.
    patrones = [os.path.join(MAESTRO_DIR, "out_gt", "**", "*Planilla*.xlsx"),
                os.path.join(GT_DIR, "*", "**", "*Planilla*.xlsx"),
                os.path.join(GT_DIR, "*", "**", "Nomina_Manual*.xlsx")]
    try:
        with open(NOMINAS_CACHE, encoding="utf-8") as f:
            cache = json.load(f)
    except Exception:
        cache = {}

    nuevo, recetas = {}, set()
    for patron in patrones:
        for ruta in glob.glob(patron, recursive=True):
            if os.path.basename(ruta).startswith("~$"):
                continue
            try:
                st = os.stat(ruta)
            except OSError:
                continue
            sello = [int(st.st_mtime), st.st_size]
            hit = cache.get(ruta)
            lista = hit[2] if (hit and hit[:2] == sello) else _recetas_de_planilla(ruta)
            nuevo[ruta] = [sello[0], sello[1], lista]
            recetas.update(lista)

    if nuevo != cache:
        try:
            with open(NOMINAS_CACHE, "w", encoding="utf-8") as f:
                json.dump(nuevo, f, ensure_ascii=False)
        except Exception as e:
            print(f"  [aviso] no pude escribir {os.path.basename(NOMINAS_CACHE)}: {e}")
    return recetas


def _recetas_en_gt_maestro():
    """Nº de receta que gt_maestro.xlsx da por YA DESPACHADAS — solo las que
    tienen un Estado avanzado (_ESTADOS_YA_DESPACHADA), no la mera presencia
    de la fila.

    Bug real detectado 11-09-2026: esta función devolvía TODAS las recetas de
    TODAS las hojas del maestro y cruce_gt.py las trataba como "ya tiene
    nómina generada". Pero gt_maestro.xlsx es el registro anual de Gestión
    Territorial — lo alimentan _sincronizar_maestro() con el reporte crudo
    COMPLETO (estado "EN PREPARACIÓN"), agregar_gt_manual.py, la revisión de
    solicitudes y la edición manual de la QF. Así, cualquier receta que
    apareciera en un reporte un día en que no se le generó planilla quedaba
    marcada como "ya procesada" para siempre y el pipeline no la volvía a
    mirar nunca. Caso comprobado:
    reporteGestionTerritorial_05-08-2026_24-08-2026.xlsx traía 126 recetas de
    6 destinos (CESFAM QUEPE, CESFAM TEODORO SCHMIDT, TOLTEN HOSP., PSR
    QUEULE, GORBEA HOSP., GORBEA DSM) y out_gt/ solo recibió la planilla de
    GORBEA HOSP.: las otras 121 se saltaron por "ya estar en gt_maestro" —
    en el reporte del día siguiente y en todos los posteriores también. El
    equipo lo compensó sacando esas nóminas a mano (de ahí las decenas de
    Nomina_Manual_* del árbol). La auditoría del 11-09-2026 dejó 26 recetas
    que ni siquiera alcanzó el rescate manual, 5 de ellas todavía con
    prescripciones sin entregar (CESFAM TEODORO SCHMIDT 2, CESFAM QUEPE 1,
    LONCOCHE HOSP. 1, PSR QUEULE 1).

    La prueba dura de "ya tiene nómina" es la planilla misma
    (_recetas_con_nomina); esta función queda como refuerzo para las recetas
    cuyo archivo de nómina ya no esté en disco pero que la QF marcó como
    enviada/entregada en el maestro."""
    import gt_maestro as GM
    wb, _ = GM.cargar_maestro()
    recetas = set()
    for ws in wb.worksheets:
        if ws.title == GM.HOJA_HISTORIAL:
            continue
        try:
            fila_header, headers = GM._headers_de_hoja(ws)
        except Exception:
            continue
        idx_receta = GM._col_index(headers, "receta")
        idx_estado = GM._col_index(headers, "estado")
        if idx_receta is None or idx_estado is None:
            continue
        for r in range(fila_header + 1, ws.max_row + 1):
            v = ws.cell(row=r, column=idx_receta + 1).value
            if v is None or not str(v).strip():
                continue
            if GM._norm(ws.cell(row=r, column=idx_estado + 1).value) in _ESTADOS_YA_DESPACHADA:
                recetas.add(str(v).strip())
    return recetas


def _recetas_ya_con_nomina():
    """Set de dedup del pipeline GT: recetas que ya salieron en una nómina.
    Une la prueba dura (planillas en disco) con el refuerzo del maestro."""
    recetas = _recetas_con_nomina()
    try:
        recetas |= _recetas_en_gt_maestro()
    except Exception as e:
        print(f"  [aviso] no pude leer gt_maestro.xlsx para el dedup: {e} — "
              f"sigo solo con las planillas en disco")
    return recetas


def _sincronizar_maestro(reporte, estado_destino="EN PREPARACIÓN"):
    """Registra en gt_maestro.xlsx todas las recetas del reporte recién
    procesado (nuevas y ya existentes — upsert idempotente): el maestro es el
    REGISTRO anual de Gestión Territorial, así que debe reflejar todo lo que
    SSASUR despachó, se le haya generado planilla o no.

    OJO: este registro NO es el dedup del pipeline. Lo fue hasta el
    11-09-2026 y por eso se perdían nóminas — una receta registrada acá en
    estado "EN PREPARACIÓN" quedaba marcada como ya procesada aunque nunca
    hubiera salido en una planilla. El dedup ahora lo decide
    _recetas_ya_con_nomina() a partir de las planillas realmente generadas."""
    import gt_maestro as GM
    wb, path = GM.cargar_maestro()
    resumen, hojas = GM.sincronizar_gt_report(wb, reporte, estado_destino=estado_destino)
    GM.guardar(wb, path)
    return resumen, hojas


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("reporte", help="reporteGestionTerritorial*.xlsx")
    ap.add_argument("--salida", default="./out_gt")
    ap.add_argument("--hist-glob", default=os.path.join(MAESTRO_DIR, "informe_completo_recetas*.csv"))
    ap.add_argument("--generar", action="store_true", help="Invocar generar.py del skill al terminar el cruce")
    ap.add_argument("--no-pdf", action="store_true", help="No generar PDFs (pasa --no-pdf a generar.py)")
    ap.add_argument("--no-dedup", action="store_true", help="No filtrar recetas que ya salieron en una nómina")
    ap.add_argument("--no-sync-maestro", action="store_true",
                     help="No registrar en gt_maestro.xlsx después de generar (debug/pruebas)")
    a = ap.parse_args()

    os.makedirs(a.salida, exist_ok=True)
    regs, hdr = leer_reporte_gt(a.reporte)
    print(f"Reporte GT: {len(regs)} recetas únicas | columnas detectadas OK")

    # Dedup: excluir recetas que YA salieron en una Nómina de Envío. La fuente
    # es la planilla misma en disco (_recetas_con_nomina), más las que el
    # maestro marca como enviada/entregada — NO la mera presencia de la fila
    # en gt_maestro.xlsx, que daba por procesadas recetas que nunca llegaron a
    # tener nómina (ver _recetas_en_gt_maestro).
    if not a.no_dedup:
        try:
            ya_procesadas = _recetas_ya_con_nomina()
        except Exception as e:
            print(f"  [aviso] no pude calcular el dedup GT: {e} — sigo sin dedup")
            ya_procesadas = set()
        if ya_procesadas:
            antes = len(regs)
            regs = {k: v for k, v in regs.items() if k not in ya_procesadas}
            omitidas = antes - len(regs)
            if omitidas:
                print(f"  [dedup GT] {omitidas} receta(s) omitidas por tener ya una nómina generada")
    recetas_set = set(regs.keys())

    archivos = sorted(glob.glob(a.hist_glob))
    print(f"Histórico: {len(archivos)} archivo(s) — cruzando por Nº receta (puede tardar)...")
    det = cruzar_historico(recetas_set, archivos)
    encontradas = sum(1 for r in recetas_set if r in det and det[r]["lineas"])
    print(f"  Cruce: {encontradas}/{len(recetas_set)} recetas encontradas en el histórico")

    for rec, reg in regs.items():
        clasificar(reg, det.get(rec, {"tipo_receta":"","estado":"","gt":"","lineas":OrderedDict()}))

    # Resumen
    n_ref = sum(1 for g in regs.values() if g["refrigerado"])
    n_con = sum(1 for g in regs.values() if g["controlado"])
    n_pen = sum(1 for g in regs.values() if g["pendiente"])
    n_iny = sum(1 for g in regs.values() if g["inyectable_lai"])
    n_nohist = sum(1 for g in regs.values() if not g["_en_historico"])
    print(f"\n  Clasificación (recetas con al menos uno):")
    print(f"    ❄  Refrigerados     : {n_ref}")
    print(f"    ⚠  Controlados      : {n_con}")
    print(f"    ⏳ Pendientes       : {n_pen}")
    print(f"    💉 Inyectables LAI  : {n_iny}")
    if n_nohist:
        print(f"    • No halladas en histórico: {n_nohist} (sin clasificar)")
    print(f"\n  Por establecimiento de destino:")
    por_dest = defaultdict(lambda: [0,0,0,0,0])
    for g in regs.values():
        d = por_dest[g["estab_destino"]]
        d[0]+=1; d[1]+=bool(g["refrigerado"]); d[2]+=bool(g["controlado"]); d[3]+=bool(g["pendiente"]); d[4]+=bool(g["inyectable_lai"])
    for dest, (t,rf,co,pe,iy) in sorted(por_dest.items()):
        print(f"    {dest:<28} {t:>3} recetas | ❄{rf}  ⚠{co}  ⏳{pe}  💉{iy}")

    # JSON enriquecido (formato registros del skill)
    CAMPOS_PUBLICOS = ("receta","paciente","edad","direccion","comuna","telefono",
                        "estab_origen","estab_destino","periodo","especialidad","n_presc",
                        "ventanilla","refrigerado","pendiente","controlado","inyectable_lai")
    out_json = os.path.join(a.salida, "gt_enriquecido.json")
    data = {
        "fecha_entrega": next((g["fecha_entrega_rep"] for g in regs.values() if g.get("fecha_entrega_rep")), ""),
        "origen": "Farmacia Hospital de Pitrufquén",
        "registros": [{k: g[k] for k in CAMPOS_PUBLICOS} for g in regs.values()],
    }
    with open(out_json, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"\n  → {out_json}")

    out_xlsx = os.path.join(a.salida, "Cruce_GT_Clasificacion.xlsx")
    escribir_excel(regs, out_xlsx)
    print(f"  → {out_xlsx}")

    if a.generar:
        generar = _buscar_generar_py()
        if not generar:
            print("\n  [aviso] No se encontró generar.py del skill — genera manualmente con:")
            print(f"    py <ruta>/generar.py {out_json} --salida {a.salida}")
        else:
            print(f"\n[GT] Generando planillas → {a.salida} ...")
            import subprocess as _sp
            # generar.py necesita "run" para la columna RUN de la planilla impresa,
            # pero el JSON persistido (gt_enriquecido.json) no debe llevar RUT (Ley 19.628).
            # Se pasa un JSON temporal con "run" que se borra apenas termina el subproceso.
            data_con_run = {**data, "registros": [
                {**{k: g[k] for k in CAMPOS_PUBLICOS}, "run": g.get("run", "")}
                for g in regs.values()
            ]}
            tmp_fd, tmp_path = tempfile.mkstemp(suffix=".json", prefix="gt_tmp_")
            try:
                with os.fdopen(tmp_fd, "w", encoding="utf-8") as f:
                    json.dump(data_con_run, f, ensure_ascii=False, indent=2)
                cmd = [sys.executable, generar, os.path.abspath(tmp_path),
                       "--salida", os.path.abspath(a.salida)]
                if a.no_pdf:
                    cmd.append("--no-pdf")
                env = os.environ.copy()
                env["PYTHONUTF8"] = "1"; env["PYTHONIOENCODING"] = "utf-8"
                ret = _sp.run(cmd, env=env)
            finally:
                try: os.remove(tmp_path)
                except OSError: pass

            # Registrar en gt_maestro.xlsx SOLO si generar.py terminó bien — si
            # falló a medio camino, mejor que el día siguiente lo vuelva a
            # intentar (dedup no las va a saltar) a que queden marcadas como
            # 'ya enviadas' sin haber salido nunca. Esto cierra el ciclo: sin
            # esto, cruce_gt.py generaba planillas pero gt_maestro nunca se
            # enteraba, así que el dedup del día siguiente no tenía cómo
            # saber que ya se habían procesado (bug real 10-08-2026).
            if ret.returncode == 0 and regs and not a.no_sync_maestro:
                try:
                    resumen, hojas = _sincronizar_maestro(a.reporte)
                    print(f"  [maestro GT] sincronizado -> {resumen} | hojas: {', '.join(hojas)}")
                except Exception as e:
                    print(f"  [aviso] no pude sincronizar con gt_maestro.xlsx: {e}")
            elif ret.returncode != 0:
                print(f"  [aviso] generar.py terminó con código {ret.returncode} — "
                      f"NO se sincronizó con gt_maestro.xlsx (se reintentará mañana).")


if __name__ == "__main__":
    main()
