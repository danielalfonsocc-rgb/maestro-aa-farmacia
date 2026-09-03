#!/usr/bin/env python3
"""
AUTO_SSASUR.py  —  Maestro AA Farmacia
═══════════════════════════════════════════════════════════════
Descarga automáticamente RECETA y ABASTECIMIENTO desde SSASUR
y actualiza los Excel del Maestro AA.

Flujo:
  1. Se abre Chromium en la página de login de SSASUR
  2. Tú te logeas manualmente (el script no toca tus credenciales)
  3. El script trabaja UN módulo a la vez en UNA sola pestaña
     (SSASUR es una SPA que exige usar una única pestaña):
       a) Clic en la tarjeta RECETA → informe de recetas → descarga
       b) Vuelve al dashboard → clic en ABASTECIMIENTO → stock → descarga
       c) Mismo módulo ABASTECIMIENTO → reporte Programación AA (mes en curso) → descarga
  4. Ejecuta maestro_aa.py para actualizar el Consolidado
  5. Publica los datos en GitHub (si está configurado)

NOTA TÉCNICA — cómo navega SSASUR (verificado por inspección del DOM):
  · Las tarjetas del dashboard son <button> de Vue (no <a>): hay que
    clicarlas por texto, p.ej. button:has-text("ABASTECIMIENTO").
  · El clic NO abre pestaña nueva: navega ESTA misma pestaña al módulo
    (www.ssasur.cl/<modulo>) y acuña la sesión del módulo vía SSO.
  · Ir directo por URL a un módulo sin pasar por el clic da 403 / rebote
    al login, porque la cookie de sesión solo cubre login.ssasur.cl.
  · RECETA: el informe sabana sale VACÍO (0 filas) si la sesión no se acuña
    para el PROYECTO (629). Hay que entrar por la TARJETA RECETA (dispara el
    puente SSO /sesion/obtener) — NO por RCE ni por deep-link. No es la firma
    ni el calendario: era el proyecto. (Verificado por captura HTTP del flujo
    manual: 30.582 filas con la sesión correcta, 0 sin ella.)
  · El reporte de stock NO requiere firma electrónica; basta elegir
    bodega = TODAS (value 0) y pulsar "Generar XLS".

Modos (CLI):
  · (sin flags)        corrida completa: recetas + GT + stock + maestro + planillas GT + publicar
  · --solo-recetas     solo baja recetas y termina
  · --gt / --solo-gt   GESTIÓN TERRITORIAL: entra por la tarjeta RECETA → Reporte
                       → Gestión Territorial; define fechas (default: ayer→hoy),
                       marca ORIGEN y baja el listado de recetas (Excel).
                       Luego corre cruce_gt.py --generar automáticamente: cruza con
                       el histórico, clasifica refrigerados/controlados/pendientes y
                       genera planillas + letreros por destino en
                       out_gt/<desde>_<hasta>/ (una carpeta por rango, no se pisan).
                       Opcional: --fecha dd/mm/yyyy (un día) o
                       --desde dd/mm/yyyy --hasta dd/mm/yyyy (rango); --debug-gt
                       vuelca el formulario [DESCUBRIR …] y guarda screenshots.
  · --no-rch           no actualiza el registro ISP de recetas cheque
  · --no-programacion  no baja el reporte de Programación AA (PASO 4b)
  · --no-controlados   no baja el informe de medicamentos controlados (PASO 3b)
  · --no-publicar      no publica en GitHub (debug). OJO: SOLO afecta GitHub —
                       el PASO 7-9 (Escritorio + Google Drive) corre igual, con
                       datos reales. No existe un flag que frene Escritorio/Drive;
                       para una corrida de prueba realmente acotada hay que
                       combinarlo con --no-gt --no-controlados --no-programacion
                       --no-servicios-farmaceuticos --no-rch,
                       y aun así el Escritorio y Drive se actualizan de verdad.
                       Confundido en vivo 07-08-2026 probando el fix de clozapina:
                       se asumió que --no-publicar frenaba todo, y terminó
                       corriendo la sincronización completa a Escritorio+Drive.
  · --servicios-farmaceuticos  fuerza el PASO 4e (recuento mensual QF, Agenda
                       Médica) ignorando el gatillo de "primeros días hábiles
                       del mes" — úsalo para la 1ª prueba supervisada o para
                       reprocesar un mes puntual junto con --mes-servicios.
  · --no-servicios-farmaceuticos  omite el PASO 4e por completo
  · --mes-servicios MM-AAAA  mes objetivo del PASO 4e (default: mes calendario
                       anterior a hoy)
  · --debug-agenda     vuelca el formulario [DESCUBRIR …] y guarda screenshots
                       del PASO 4e (Agenda Médica)

El registro ISP de recetas cheque (recetas_cheque.py) corre como PASO 5d: usa
la MISMA sábana ya descargada para agregar los folios cheque nuevos de Farmacia
AT Abierta al formulario ISP del mes vigente (carpeta de la QF, fuera del repo).

PASO 4b — Programación AA: dentro del mismo módulo ABASTECIMIENTO (después del
stock), baja el reporte "Consumos por centro de costo" del MES EN CURSO
(FARMACOS/LOCAL/FARMACIA) — lo consume programacion_aa.py para Cantidad
Programada / Cantidad Solicitada del ciclo Bodega AA. No se adelanta a meses
futuros: la Cantidad Solicitada de un mes que no ha empezado siempre viene en 0.

PASO 3b — Informe Medicamentos Controlados: dentro del módulo RECETA
(Receta → Reportes → Informe Medicamentos Controlados, misma sesión proyecto
629 que la sábana), baja el listado de despacho de controlados de Farmacia AT
Abierta del DÍA HÁBIL ANTERIOR (sin filtrar por medicamento → trae todos los
productos controlados). Solo se guarda localmente en Informes_Controlados_AA/
(no se publica a Drive/GitHub por defecto — no se confirmó si trae RUT de
pacientes; ver regla de privacidad de CLAUDE.md antes de agregarlo al publish).

(sgli_historico.py → SGLI_Historico_<fecha>.xlsx NO se agrega aquí: maestro_aa.py
YA lo corre solo, al final de su propio main() — agregarlo lo duplicaba.)

NOTA: agente_duplicados.py y auditoria_duplicados_profunda.py (llaman a la API
de Claude) NO corren aquí — se mantienen en sus .bat propios para ejecutarse
a demanda y no gastar tokens en cada corrida de AUTO_SSASUR.

PASO 4e — Servicios Farmacéuticos (recuento mensual QF): módulo NUEVO y
separado (Agenda Médica, araucaniasur/agenda_medica) — no comparte sesión con
RECETA/ABASTECIMIENTO. Dispara solo, con marcador + catch-up (mismo patrón que
Centinela), en los primeros ~5 días corridos de cada mes: baja el listado de
actividades del mes ANTERIOR completo (Proceso_nuevo → Hoja Diaria de
Profesional → Tipo Reporte "Principalmente Actividades"), corre
servicios_farmaceuticos.py (filtra ESPECIALIDAD='QUIMICO FARMACEUTICO', arma
el recuento QF × actividad) y sube el agregado a Drive (SIN RUT — el reporte
crudo con RUT de paciente NUNCA se publica, igual que la sábana de recetas).

IMPORTANTE — 1ª corrida en vivo pendiente: los selectores del PASO 4e
(entrar_agenda_medica, _fijar_establecimiento_hosp, _fijar_tipo_profesional_aps,
_marcar_tipo_reporte_actividades) son heurísticos por TEXTO, nunca verificados
contra el DOM real de Agenda Médica. Antes de dejarlo corriendo desatendido,
corre una vez con `py AUTO_SSASUR.py --servicios-farmaceuticos --debug-agenda`
delante de una persona y ajusta lo que no calce leyendo los volcados
[DESCUBRIR] — mismo protocolo de bootstrap que ya se usó para GT y Controlados.
═══════════════════════════════════════════════════════════════
"""
import asyncio
import ctypes
import csv as _csv
import glob as _glob
import os
import re
import subprocess
import sys
from pathlib import Path
from datetime import date, timedelta

# Forzar UTF-8 en la salida — si no, al redirigir a un archivo/pipe Windows usa
# cp1252 y los caracteres ═ ✓ → ✗ revientan el script antes de empezar.
# line_buffering=True: sin esto, al redirigir a un archivo/pipe Python usa buffer
# de bloque para el propio stdout del script, mientras los subprocess.run()
# (maestro_aa.py, etc.) heredan el mismo handle y escriben en tiempo real. El
# resultado es un log con el orden de eventos mezclado — ej. el texto de
# maestro_aa.py aparece ANTES que el "[5/9] Actualizando Maestro AA..." que en
# realidad lo dispara, dando la falsa impresión de pasos duplicados o de que
# un paso posterior (GT, controlados) no corrió. Bug real visto 27-07-2026.
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
    except Exception:
        pass


# ── Configuración ─────────────────────────────────────────────────────────────
MAESTRO_DIR      = Path(__file__).parent
SESSION_FILE     = MAESTRO_DIR / ".ssasur_session.json"
DASHBOARD_URL    = "https://login.ssasur.cl/dashboard"
RECETA_INFORME   = "https://www.ssasur.cl/receta/informes/sabana"
STOCK_REPORTE    = "https://www.ssasur.cl/abastecimiento/reportes/stock_en_momento_bodega"
BODEGA_TODAS     = "0"    # bodega → "TODAS" (todas las bodegas en un archivo)

# ── Programación AA (reporte "Consumos por centro de costo") ───────────────────
# Consumido por programacion_aa.py. Ir directo por URL SOLO funciona después de
# entrar_modulo(page, "ABASTECIMIENTO") (acuña la sesión del módulo) — igual que
# STOCK_REPORTE; probado en vivo 15-07-2026 (403 sin el módulo acuñado antes).
PROGRAMACION_REPORTE = "https://www.ssasur.cl/abastecimiento/reportes/consumo_por_cc_desplegado"
DISTRIB_FARMACOS = "5"    # <select id="distribucion"> → FARMACOS
PROYECTO_LOCAL   = "3"    # <select id="select_proyecto"> → LOCAL (no "DAC LOCAL")
CC_FARMACIA      = "75"   # <select id="cc"> → FARMACIA
TIMEOUT_LOGIN    = 300_000   # 5 minutos para logarse
TIMEOUT_DESCARGA = 600_000   # 10 minutos por descarga

# ── GT · Gestión territorial (informe de modalidad de despacho) ────────────────
# Reporte que alimenta el skill de gestión territorial: el LISTADO de recetas a
# despachar. Vive DENTRO del módulo RECETA (misma sesión proyecto 629 que la
# sabana). Flujo en la página: definir fechas → marcar ORIGEN en Origen/Destino →
# aparece el listado → botón Excel. El .xlsx cae en la carpeta de gestión
# territorial (NO entra al maestro_aa); el skill lo cruza con el histórico de
# recetas por día de despacho y cuenta refrigerados/pendientes/controlados.
GT_DIR = MAESTRO_DIR.parent / "04_Farmacia_Gestion_Territorial"
# URL real del reporte (confirmada por el usuario). Si no carga el formulario, el
# script navega por el menú Reporte → Gestión Territorial.
GESTION_TERRITORIAL_URL = "https://www.ssasur.cl/receta/informes/gestionTerritorial"
# Campos de fecha del reporte (ids reales confirmados en vivo 18-06-2026, dd/mm/aaaa).
SEL_FECHA_INI = "fechaInicio"
SEL_FECHA_FIN = "fechaTermino"
# Botones (se prueba el primero que exista). AJUSTAR tras la 1ª corrida en vivo
# leyendo la salida [DESCUBRIR …] que imprime el script.
SELS_BUSCAR = ('button:has-text("Buscar")', 'a:has-text("Buscar")',
               'input[type="submit"][value*="Buscar" i]', '#buscar', '#btnBuscar')
SELS_EXCEL  = ('button:has-text("Excel")', 'a:has-text("Excel")',
               'button:has-text("XLS")', 'a:has-text("XLS")',
               '[id*="xls" i]', '[id*="excel" i]', '[class*="excel" i]',
               'img[src*="excel" i]', 'i.fa-file-excel')

# ── RECETA · Informe Medicamentos Controlados (Farmacia AT Abierta) ────────────
# Vive DENTRO del módulo RECETA (Reportes → Informe Medicamentos Controlados),
# misma sesión proyecto 629 que la sábana. Formulario: Medicamento (se deja
# vacío a propósito — el propio formulario indica que así trae TODOS los
# productos controlados que pasaron por la bodega), Bodega (elegir "FARMACIA AT
# ABIERTA"), Fecha Inicio/Término (mismo id fechaInicio/fechaTermino que GT,
# confirmado por inspección visual del formulario 23-07-2026) y botón Buscar.
CONTROLADOS_URL = "https://www.ssasur.cl/receta/informes/controlados"
CONTROLADOS_DIR = MAESTRO_DIR / "Informes_Controlados_AA"


# ── SERVICIOS FARMACÉUTICOS · Agenda Médica (recuento mensual de actividades QF) ─
# Módulo NUEVO y separado de RECETA/ABASTECIMIENTO: vive en otra app dentro de
# SSASUR (araucaniasur/agenda_medica), con su propio menú "Proceso_nuevo" →
# "Hoja Diaria de Profesional". A diferencia de los demás pasos, este NUNCA se
# probó en vivo — los selectores de abajo son heurísticos (por TEXTO, mismo
# patrón que _marcar_origen/_seleccionar_bodega_at_abierta) para maximizar la
# chance de acertar a la primera, pero la 1ª corrida real DEBE hacerse con
# --debug-agenda delante de una persona para leer los volcados [DESCUBRIR] y
# ajustar lo que no calce — exactamente el mismo protocolo de bootstrap que ya
# se usó para GT y Controlados (ver sus comentarios arriba).
AGENDA_MEDICA_URL = "https://www.ssasur.cl/araucaniasur/agenda_medica/index_menu.php"
SERVICIOS_MARCADOR = MAESTRO_DIR / "Servicios_Farmaceuticos" / "_ultimo_mes.json"
# Última fecha en que AUTO_SSASUR completó el PASO 5 (maestro_aa.py) sin
# error. Permite detectar huecos REALES (PC apagado, sesión SSASUR caída,
# feriado no listado, etc.) en vez de asumir que el hueco es siempre
# "fin de semana/feriado" como hacía dia_habil_anterior() a solas — ver
# ultima_corrida_ok()/rango_backlog_gt().
ULTIMA_CORRIDA_JSON = MAESTRO_DIR / "_ultima_corrida_ok.json"
SELS_PROCESO_NUEVO = ('a:has-text("Proceso_nuevo")', 'button:has-text("Proceso_nuevo")',
                       ':text("Proceso_nuevo")', 'a:has-text("Proceso Nuevo")')
SELS_HOJA_DIARIA = ('a:has-text("Hoja Diaria de Profesional")', ':text("Hoja Diaria de Profesional")',
                     'a:has-text("Hoja Diaria")', ':text("Hoja Diaria")')
# "Hoja Diaria de Profesional" es en realidad un submenú (hover) con dos hijos:
# "Reporte Hoja Diaria" y "Reporte Procedimientos" — confirmado por screenshot
# de la 1ª corrida real (03-09-2026, debug_agenda.png): el clic anterior solo
# abría el submenú y se quedaba ahí, sin navegar nunca al formulario.
SELS_REPORTE_HOJA_DIARIA = ('a:has-text("Reporte Hoja Diaria")', ':text("Reporte Hoja Diaria")')
# El formulario en sí NO vive en el documento principal de agenda_medica: es un
# <iframe> separado (.../hoja_diaria_profesional/reportes/CriterioSabana.php).
# Confirmado en vivo 03-09-2026 con un probe aparte — por eso el chequeo viejo
# de "ok" (basado en document.body.innerText del documento top) siempre daba
# falso aunque el formulario se viera perfecto en la captura. IDs reales del
# formulario (mismo probe): FechaInicio / FechaTermino (inputs texto,
# dd/mm/aaaa), t_profesional (<select>, value "1" = APS), establecimiento
# (<select>, value "59" = PITRUFQUEN HOSP. — SOLO se puebla en cascada DESPUÉS
# de fijar t_profesional, antes solo trae "SELECCIONE ..."), t_reporte_A
# (<input type=radio>, "Principalmente Actividades").
SEL_AGENDA_FECHA_INI = "FechaInicio"
SEL_AGENDA_FECHA_FIN = "FechaTermino"


def _frame_hoja_diaria(page):
    """Devuelve el Frame del formulario Hoja Diaria (CriterioSabana.php) si ya
    está montado en la página, o None."""
    for fr in page.frames:
        if "hoja_diaria_profesional" in fr.url:
            return fr
    return None


async def entrar_agenda_medica(page):
    """Dashboard → tarjeta AGENDA MÉDICA. Si no aparece ninguna tarjeta con ese
    texto (el dashboard de SSASUR agrupa módulos por texto de botón, ver
    entrar_modulo), cae a navegar directo por URL — con el mismo riesgo de
    rebote/403 que ya se documentó para RCE si la sesión no queda acuñada al
    proyecto correcto (ver notas de entrar_receta)."""
    try:
        await entrar_modulo(page, "AGENDA")
        return
    except Exception:
        pass
    print("  [info] No encontré tarjeta 'AGENDA' en el dashboard — navego directo por URL.")
    await page.goto(AGENDA_MEDICA_URL)
    await _networkidle(page)
    await page.wait_for_timeout(1_500)
    print(f"  ✓ En Agenda Médica: {page.url}")


async def _abrir_hoja_diaria_profesional(page, debug=False):
    """Proceso_nuevo → Hoja Diaria de Profesional → Reporte Hoja Diaria.
    Devuelve el Frame del formulario (Fecha Inicio/Termino, Tipo Profesional,
    Establecimiento, Tipo Reporte) si quedó cargado — vive en un <iframe>
    propio, no en el documento principal — o None si no se pudo abrir."""
    try:
        await _click_primero(page, SELS_PROCESO_NUEVO, "Proceso_nuevo")
        await page.wait_for_timeout(800)
    except Exception:
        if debug:
            await _dump_formulario(page, "agenda-sin-proceso_nuevo")
    try:
        await _click_primero(page, SELS_HOJA_DIARIA, "Hoja Diaria de Profesional")
        await page.wait_for_timeout(500)
    except Exception:
        if debug:
            await _dump_formulario(page, "agenda-sin-hoja_diaria")
        return None
    try:
        await _click_primero(page, SELS_REPORTE_HOJA_DIARIA, "Reporte Hoja Diaria")
    except Exception:
        if debug:
            await _dump_formulario(page, "agenda-sin-reporte_hoja_diaria")
        return None
    await _networkidle(page)
    await page.wait_for_timeout(1_500)
    frm = _frame_hoja_diaria(page)
    if debug:
        await _dump_formulario(frm or page, "agenda-hoja_diaria")
    return frm


async def _fijar_tipo_profesional_aps(frame):
    """<select id="t_profesional"> → 'APS' (value "1"). IDs y valor
    confirmados en vivo 03-09-2026 con un probe aparte contra el frame real
    del formulario. Debe llamarse ANTES que _fijar_establecimiento_hosp: fijar
    esto dispara el cascade que puebla el <select id="establecimiento">
    (que si no, solo trae 'SELECCIONE ...')."""
    try:
        await frame.select_option("#t_profesional", "1")
        await frame.wait_for_timeout(800)
        return {"found": True, "cambiado": True, "actual": "APS"}
    except Exception:
        return None


async def _fijar_establecimiento_hosp(frame):
    """<select id="establecimiento"> → 'PITRUFQUEN HOSP.' (value "59" en vivo
    03-09-2026, pero se busca por texto por si SSASUR cambia el value). Este
    select solo tiene opciones DESPUÉS de fijar Tipo Profesional — ver
    _fijar_tipo_profesional_aps. Devuelve una etiqueta de lo que hizo, o None
    si no encontró el select o la opción."""
    try:
        opts = await frame.evaluate(r"""() => {
          const s = document.querySelector('#establecimiento');
          if (!s) return null;
          return {
            actual: s.options[s.selectedIndex] ? s.options[s.selectedIndex].textContent.trim() : '',
            opts: [...s.options].map(o => ({value: o.value, text: (o.textContent || '').trim()})),
          };
        }""")
    except Exception:
        return None
    if not opts:
        return None
    if re.search(r"hosp", opts["actual"], re.I):
        return f"ya en HOSP. ({opts['actual']})"
    destino = next(
        (o for o in opts["opts"] if re.search(r"pitrufquen", o["text"], re.I) and re.search(r"hosp", o["text"], re.I)),
        None,
    )
    if not destino:
        return None
    try:
        await frame.select_option("#establecimiento", destino["value"])
        return f"cambiado → {destino['text']}"
    except Exception as e:
        return f"[AVISO] encontré la opción pero no pude seleccionarla: {e}"


async def _marcar_tipo_reporte_actividades(frame):
    """<input type=radio id="t_reporte_A" name="t_reporte" value="A"> →
    'Principalmente Actividades'. ID confirmado en vivo 03-09-2026."""
    try:
        await frame.check("#t_reporte_A")
        return True
    except Exception:
        return False


async def paso_agenda_actividades(page, desde: str, hasta: str, debug=False):
    """Agenda Médica → Proceso_nuevo → Hoja Diaria de Profesional. Fija fechas
    (mes objetivo), Tipo Profesional=APS, Establecimiento=PITRUFQUEN HOSP.,
    Tipo Reporte=Principalmente Actividades, Buscar, descarga Excel.
    Devuelve (archivo|None, n_filas): n=0 sin datos; n=-1 error."""
    print(f"\n[Servicios Farmacéuticos] Agenda Médica — Hoja Diaria de Profesional ({desde} → {hasta})")
    await entrar_agenda_medica(page)
    frm = await _abrir_hoja_diaria_profesional(page, debug)
    if frm is None:
        print("  [ERROR] No cargó el formulario Hoja Diaria de Profesional.")
        await page.screenshot(path=str(MAESTRO_DIR / "debug_agenda.png"))
        return (None, -1)

    await _set_fechas(frm, desde, hasta, id_ini=SEL_AGENDA_FECHA_INI, id_fin=SEL_AGENDA_FECHA_FIN)
    print(f"  Fechas: {desde} → {hasta}")
    await frm.wait_for_timeout(500)

    # Tipo Profesional PRIMERO: dispara el cascade que puebla Establecimiento.
    tipo_prof = await _fijar_tipo_profesional_aps(frm)
    if tipo_prof and tipo_prof.get("cambiado"):
        print(f"  Tipo Profesional → {tipo_prof['actual']}")
    elif not tipo_prof:
        print("  [AVISO] No encontré el select de Tipo Profesional — sigo con el default.")

    estab = await _fijar_establecimiento_hosp(frm)
    print(f"  Establecimiento → {estab}" if estab else "  [AVISO] No encontré el select de Establecimiento.")

    marcado = await _marcar_tipo_reporte_actividades(frm)
    print("  Tipo Reporte → Principalmente Actividades" if marcado
          else "  [AVISO] No encontré el radio 'Principalmente Actividades'.")
    await frm.wait_for_timeout(500)

    # Escribir en FechaInicio/FechaTermino abre un date-picker emergente que
    # se queda tapando el botón Buscar (Playwright lo ve "no accionable" y el
    # clic falla aunque el texto exista en el DOM) — visto en vivo 03-09-2026.
    # Cerrar el picker (botón "Cerrar" si sigue abierto) y sacar el foco con
    # Escape antes de buscar.
    try:
        await frm.click(':text("Cerrar")', timeout=1_000)
    except Exception:
        pass
    try:
        await page.keyboard.press("Escape")
    except Exception:
        pass
    await frm.wait_for_timeout(300)

    if debug:
        await _dump_formulario(frm, "agenda-pre-buscar")

    try:
        await _click_primero(frm, SELS_BUSCAR, "Buscar")
    except Exception as e:
        print(f"  [ERROR] No encontré el botón Buscar: {e}")
        await _dump_formulario(frm, "agenda-sin-buscar")
        await page.screenshot(path=str(MAESTRO_DIR / "debug_agenda.png"))
        return (None, -1)
    await _networkidle(page)
    await frm.wait_for_timeout(2_500)
    if debug:
        await _dump_formulario(frm, "agenda-post-buscar")
        await page.screenshot(path=str(MAESTRO_DIR / "debug_agenda.png"))

    n = await _filas_resultado(frm)
    if n == 0:
        print("  (sin actividades en el rango — no hay Excel que bajar)")
        return (None, 0)

    d, h = desde.replace("/", "-"), hasta.replace("/", "-")
    dest = MAESTRO_DIR / f"reporte_listado_actividades_{d}_{h}.xlsx"
    try:
        await descargar_como(page, dest,
                             lambda: _click_primero(frm, SELS_EXCEL, "Excel", force=True))
    except Exception as e:
        print(f"  [ERROR] No se pudo bajar el Excel: {e}")
        await page.screenshot(path=str(MAESTRO_DIR / "debug_agenda.png"))
        return (None, -1)
    return (dest, _contar_filas_xlsx(dest))


def fmt(d: date) -> str:
    return d.strftime("%d/%m/%Y")


def _feriados_chile() -> set:
    """Lee feriados_chile.csv (mismo archivo/formato 'fecha;nombre;confianza'
    que ya usa pedido_fusion.py — una sola fuente de feriados para todo el
    proyecto)."""
    fer = set()
    ruta = MAESTRO_DIR / "feriados_chile.csv"
    try:
        with open(ruta, encoding="utf-8") as fh:
            next(fh)
            for ln in fh:
                p = ln.rstrip().split(";")
                if p[0].strip():
                    fer.add(date.fromisoformat(p[0].strip()))
    except FileNotFoundError:
        pass
    return fer


def dia_habil_anterior(d: date) -> date:
    """Retrocede desde d hasta el primer día hábil anterior (lunes-viernes,
    sin feriado chileno)."""
    fer = _feriados_chile()
    ant = d - timedelta(days=1)
    while ant.weekday() >= 5 or ant in fer:
        ant -= timedelta(days=1)
    return ant


def primer_dia_habil(anio: int, mes: int) -> date:
    """Primer día hábil (lunes-viernes, sin feriado chileno) de un mes."""
    fer = _feriados_chile()
    d = date(anio, mes, 1)
    while d.weekday() >= 5 or d in fer:
        d += timedelta(days=1)
    return d


def _mes_anterior(hoy: date) -> tuple[int, int]:
    """(año, mes) del mes calendario anterior a `hoy`."""
    if hoy.month == 1:
        return hoy.year - 1, 12
    return hoy.year, hoy.month - 1


def _ultimo_dia_mes(anio: int, mes: int) -> date:
    if mes == 12:
        return date(anio, 12, 31)
    return date(anio, mes + 1, 1) - timedelta(days=1)


def ultima_corrida_ok() -> date | None:
    """Fecha de la última vez que AUTO_SSASUR completó el PASO 5 (maestro_aa.py)
    sin error, según ULTIMA_CORRIDA_JSON. None si no hay marcador (1ª corrida,
    o archivo borrado)."""
    if not ULTIMA_CORRIDA_JSON.exists():
        return None
    try:
        import json as _json
        data = _json.loads(ULTIMA_CORRIDA_JSON.read_text(encoding="utf-8"))
        return date.fromisoformat(data["fecha"])
    except Exception:
        return None


def marcar_corrida_ok(d: date) -> None:
    """Registra `d` como la última corrida completada sin error — ver
    ultima_corrida_ok()."""
    try:
        import json as _json
        ULTIMA_CORRIDA_JSON.write_text(
            _json.dumps({"fecha": d.isoformat()}, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
    except Exception as e:
        print(f"  [aviso] no se pudo escribir {ULTIMA_CORRIDA_JSON.name}: {e}")


def rango_backlog_gt(hoy: date) -> tuple[date, date] | None:
    """Bug real detectado 30-07-2026 (caso Teodoro Schmidt, 23 recetas GT
    despachadas el viernes 24 y sábado 25 de julio sin nómina): el informe
    'Modalidad de Despacho' que usa paso_gt() solo muestra despachos
    PENDIENTES — una vez que SSASUR procesa/entrega la receta, desaparece del
    listado para siempre. Si no corre AUTO_SSASUR un día (fin de semana,
    feriado, PC apagado, corrida fallida), cualquier receta que se digitó Y
    despachó ese mismo día queda invisible para SIEMPRE a paso_gt(), sin
    importar qué tan amplio sea el rango de fechas que se consulte (verificado
    en vivo: ni ampliando la ventana hacia atrás aparecen).

    El punto de partida NO asume que el único hueco posible es fin de
    semana/feriado (dia_habil_anterior): si ULTIMA_CORRIDA_JSON dice que la
    última corrida exitosa fue anterior a eso — porque se saltó un día hábil
    normal — el rango se extiende hacia atrás hasta cubrir ese hueco real.

    Devuelve el rango de fechas [desde, hasta] a revisar contra el HISTÓRICO
    real (informe_completo_recetas*.csv, no el reporte de pendientes) para
    detectar recetas GT que se despacharon en ese hueco. None si no hay hueco
    que revisar."""
    desde = dia_habil_anterior(hoy)   # inclusive: el propio último día hábil puede
                                       # tener despachos posteriores a la corrida de esa mañana
    ultima_ok = ultima_corrida_ok()
    if ultima_ok is not None and ultima_ok + timedelta(days=1) < desde:
        desde = ultima_ok + timedelta(days=1)   # hueco real más amplio que el calendario
    hasta = hoy - timedelta(days=1)   # ayer
    if desde >= hoy:
        return None
    return (desde, hasta)


def verificar_backlog_gt(hoy: date) -> None:
    """Cruza informe_completo_recetas*.csv (histórico real, no el reporte de
    pendientes) contra gt_maestro.xlsx para encontrar recetas con Gestión
    Territorial='S' y Fecha Entrega Receta dentro del hueco fin de
    semana/feriado que paso_gt() no pudo ver (ver rango_backlog_gt). Antes
    esto solo IMPRIMÍA el hallazgo y esperaba que un QF corriera
    GT_NOMINA_PARTICULAR.bat a mano (en la práctica, no pasaba de forma
    confiable — ver _generar_backlog_gt). Ahora, además de avisar, GENERA
    y REGISTRA la(s) Nómina(s) de Envío que faltan automáticamente."""
    rango = rango_backlog_gt(hoy)
    if rango is None:
        return
    desde, hasta = rango
    fechas_ok = {(desde + timedelta(days=i)).strftime("%d/%m/%Y")
                 for i in range((hasta - desde).days + 1)}

    # Campos que interesan de informe_completo_recetas*.csv. OJO: "Periodo"
    # aparece DOS VECES en ese CSV — la columna de la receta (tras "Estado",
    # valores tipo "1 de 1") y la de la dosis/prescripción (tras "Intervalo",
    # valores "DIAS"/"HORAS"). csv.DictReader se queda con la ÚLTIMA columna
    # de nombre repetido (la equivocada) — por eso acá se lee por posición y
    # se usa la PRIMERA ocurrencia de cada nombre, igual que hace
    # maestro_aa.py con pandas al excluir el sufijo ".1" de read_csv.
    # Bug real 10-08-2026: "Fecha Entrega Receta" en este CSV NO significa
    # "ya se entregó" — puede ser una fecha objetivo/agendada que después se
    # reprograma. La prueba real: la receta 45206882 traía Fecha Entrega
    # Receta=06/08/2026 con Estado='PENDIENTE', y seguía viva y "EN TRÁNSITO"
    # en el reporte de Modalidad de Despacho días después con OTRA fecha
    # (10/08). Sin filtrar por Estado='ENTREGADA', este chequeo generaba
    # nóminas prematuras/duplicadas para recetas que la corrida normal de
    # cruce_gt.py iba a procesar de todos modos (6 de 50 casos ese día). Solo
    # cuenta como "backlog perdido" una receta genuinamente ya despachada por
    # la farmacia (Estado='ENTREGADA') que además ya no aparece en ningún
    # reporte de pendientes futuro.
    _CAMPOS_BACKLOG = ("Nombre", "Apellido Paterno", "Apellido Materno", "RUN",
                        "Establecimiento Retira G. Territorial", "Comuna", "Periodo",
                        "Especialidad", "Fono 1", "Fecha Entrega Receta", "Estado")

    filas_por_receta = {}
    conteo_por_receta = {}   # receta -> nº de líneas de prescripción (usado como n_presc del backfill)
    lineas_por_receta = {}   # receta -> [(Prescripción, Cantidad Recetada), ...] (para detectar refrigerados)
    for path in sorted(_glob.glob(str(MAESTRO_DIR / "informe_completo_recetas*.csv"))):
        try:
            with open(path, encoding="latin-1", newline="") as fh:
                rd = _csv.reader(fh, delimiter=";")
                try:
                    hdr = [h.strip() for h in next(rd)]
                except StopIteration:
                    continue
                idx = {}
                for i, h in enumerate(hdr):
                    if h not in idx:   # primera ocurrencia manda (ver nota arriba)
                        idx[h] = i

                def _get(r, campo, _idx=idx):
                    i = _idx.get(campo)
                    return r[i].strip() if (i is not None and i < len(r)) else ""

                i_gt = idx.get("Gestión Territorial")
                i_fecha = idx.get("Fecha Entrega Receta")
                i_receta = idx.get("Número Receta")
                i_presc = idx.get("Prescripción")
                i_cant = idx.get("Cantidad Recetada")
                i_estado = idx.get("Estado")
                if i_gt is None or i_fecha is None or i_receta is None:
                    continue
                for r in rd:
                    if i_gt >= len(r) or (r[i_gt] or "").strip().upper() != "S":
                        continue
                    if i_fecha >= len(r) or (r[i_fecha] or "").strip() not in fechas_ok:
                        continue
                    # Solo Estado=ENTREGADA — PENDIENTE/otros siguen su curso normal
                    # por cruce_gt.py, que ahora dedupe correctamente contra
                    # gt_maestro.xlsx (ver cruce_gt._recetas_en_gt_maestro).
                    if i_estado is None or i_estado >= len(r) or (r[i_estado] or "").strip().upper() != "ENTREGADA":
                        continue
                    n = (r[i_receta] or "").strip() if i_receta < len(r) else ""
                    if not n:
                        continue
                    if n not in filas_por_receta:
                        filas_por_receta[n] = {campo: _get(r, campo) for campo in _CAMPOS_BACKLOG}
                    conteo_por_receta[n] = conteo_por_receta.get(n, 0) + 1
                    prod = r[i_presc].strip() if (i_presc is not None and i_presc < len(r)) else ""
                    if prod:
                        cant = r[i_cant].strip() if (i_cant is not None and i_cant < len(r)) else ""
                        lineas_por_receta.setdefault(n, []).append((prod, cant))
        except Exception:
            continue
    if not filas_por_receta:
        return

    try:
        import gt_maestro as _GM
        wb, _ = _GM.cargar_maestro()
        faltantes = [(n, r) for n, r in filas_por_receta.items()
                     if _GM.buscar_receta_en_maestro(wb, n) is None]
    except Exception as e:
        print(f"\n[AVISO] No pude cruzar contra gt_maestro.xlsx para el chequeo de backlog GT: {e}")
        faltantes = list(filas_por_receta.items())

    if not faltantes:
        return

    print("\n" + "═" * 62)
    print(f"  [ALERTA] Backlog GT sin nómina — {desde.strftime('%d/%m/%Y')} → {hasta.strftime('%d/%m/%Y')}")
    print("═" * 62)
    print(f"  {len(faltantes)} receta(s) con Gestión Territorial='S' despachada(s) en ese rango")
    print("  y SIN fila en el maestro GT — no salieron con AUTO_SSASUR (el reporte de")
    print("  pendientes ya no las muestra, quedaron invisibles).")
    for n, r in sorted(faltantes, key=lambda t: t[0]):
        nombre = f"{r.get('Nombre','').strip()} {r.get('Apellido Paterno','').strip()} {r.get('Apellido Materno','').strip()}"
        destino = r.get("Establecimiento Retira G. Territorial") or f"(sin destino — comuna: {r.get('Comuna','?')})"
        print(f"    {n} | {nombre} | {destino} | entregada {r.get('Fecha Entrega Receta')}")
    print("═" * 62)

    _generar_backlog_gt(faltantes, conteo_por_receta, lineas_por_receta)


def _generar_backlog_gt(faltantes, conteo_por_receta, lineas_por_receta=None) -> None:
    """Genera y registra AUTOMÁTICAMENTE la(s) Nómina(s) de Envío para las
    recetas que verificar_backlog_gt() detectó sin fila en el maestro GT.

    Antes esto solo se imprimía en el log (ver git blame) y dependía de que
    alguien lo viera y corriera GT_NOMINA_PARTICULAR.bat a mano — en la
    práctica no pasaba: caso real detectado 07-08-2026, 50 recetas
    entregadas el 06-08-2026 (23 CESFAM QUEPE, 15 TOLTEN HOSP., 8 PSR
    QUEULE, 4 CESFAM TEODORO SCHMIDT) llevaban acumuladas SIN nómina.

    OJO — la opción 2 de GT_NOMINA_PARTICULAR.bat (agregar_gt_manual.py
    --gt-excel) NO sirve para este caso: ese flujo lee un Informe Modalidad
    de Despacho recién bajado, pero estas recetas ya desaparecieron de ese
    reporte (SSASUR solo lista despachos PENDIENTES — ver rango_backlog_gt).
    Por eso acá se arman las filas directo desde el histórico real
    (informe_completo_recetas*.csv, ya leído por verificar_backlog_gt) y se
    reusa la MISMA lógica de registro + generación de agregar_gt_manual.py
    (equivalente a su opción 1, "--csv" pero sin el paso manual de Excel)."""
    try:
        import agregar_gt_manual as _AGM
        import gt_maestro as _GM
    except Exception as e:
        print(f"  [AVISO] No pude generar automáticamente las nóminas del backlog: {e}")
        print("  → Corre GT_NOMINA_PARTICULAR.bat (opción 1) a mano con estos datos.")
        return
    try:
        import cruce_gt as _CG
    except Exception:
        _CG = None   # sin detección de refrigerados — la nómina/letrero se generan igual, sin la lista

    lineas_por_receta = lineas_por_receta or {}

    def _refrigerado_texto(n):
        """Mismo formato que cruce_gt.clasificar(): 'Nombre xCantidad; ...'."""
        if _CG is None:
            return ""
        vistos = []
        for prod, cant in lineas_por_receta.get(n, []):
            lab = _CG.es_refrigerado(prod)
            if lab:
                vistos.append(f"{lab} x{cant or '1'}")
        return "; ".join(dict.fromkeys(vistos))

    filas, sin_destino = [], []
    for n, row in faltantes:
        # "TOLTEN HOSP." trae punto final en el histórico crudo; se saca para calzar
        # con _CARPETA_LOCAL de agregar_gt_manual.py (mismo tratamiento que --gt-excel).
        destino = (row.get("Establecimiento Retira G. Territorial") or "").strip().upper().rstrip(".")
        if not destino:
            sin_destino.append(n)
            continue
        nombre = f"{row.get('Nombre','').strip()} {row.get('Apellido Paterno','').strip()} {row.get('Apellido Materno','').strip()}".strip()
        filas.append({
            "receta": n, "paciente": nombre,
            "rut": (row.get("RUN") or "").strip(),
            "destino": destino,
            "periodo": (row.get("Periodo") or "").strip(),
            "especialidad": (row.get("Especialidad") or "").strip(),
            "n_presc": str(conteo_por_receta.get(n, 1)),
            "telefono": (row.get("Fono 1") or "").strip(),
            "pendiente": "",
            "refrigerado": _refrigerado_texto(n),
        })
    if sin_destino:
        print(f"  [AVISO] {len(sin_destino)} receta(s) del backlog sin 'Establecimiento Retira G. "
              f"Territorial' informado — quedan SIN generar, revisar a mano: {sin_destino}")
    if not filas:
        return

    wb, path = _GM.cargar_maestro()
    hoy = date.today()
    hojas_tocadas = {}
    for f in filas:
        receta_dict = {k: v for k, v in f.items() if k != "receta" and v} | {"receta": f["receta"]}
        ws, _resultado, _ = _GM.upsert_receta_maestro(wb, receta_dict, estado="EN PREPARACIÓN", fecha_fallback=hoy)
        hojas_tocadas[ws.title] = ws
    for ws in hojas_tocadas.values():
        _GM.aplicar_formato_maestro(ws)
    _GM.formatear_historial(wb)
    _GM.guardar(wb, path)

    por_destino = {}
    for f in filas:
        por_destino.setdefault(f["destino"], []).append(f)
    print(f"\n  [BACKLOG GT] Generando {len(filas)} receta(s) que faltaban en {len(por_destino)} nómina(s):")
    for destino, filas_destino in por_destino.items():
        ruta, ruta_letrero = _AGM._generar_nomina(destino, filas_destino, hoy)
        if ruta_letrero:
            print(f"    {destino}: {len(filas_destino)} receta(s) -> {ruta}, {ruta_letrero}")
        else:
            print(f"    {destino}: {len(filas_destino)} receta(s) -> {ruta}")
    print("  (quedaron registradas en el maestro GT y en Nóminas de Envío — despachar cuanto antes)")
    print("═" * 62)


def gt_salida(dest: Path) -> Path:
    """Carpeta de salida del cruce GT, SEPARADA por rango de fechas del reporte.
    Deriva el slug del nombre del Excel descargado
    (reporteGestionTerritorial_<desde>_<hasta> → out_gt/<desde>_<hasta>/),
    para que cada captura tenga su propia carpeta y no pise a las anteriores."""
    pref = "reporteGestionTerritorial_"
    slug = dest.stem[len(pref):] if dest.stem.startswith(pref) and len(dest.stem) > len(pref) else dest.stem
    return MAESTRO_DIR / "out_gt" / slug


def _arg_val(flag, default=None):
    """Devuelve el valor que sigue a un flag CLI (p.ej. `--fecha 17/06/2026`)."""
    if flag in sys.argv:
        i = sys.argv.index(flag)
        if i + 1 < len(sys.argv) and not sys.argv[i + 1].startswith("--"):
            return sys.argv[i + 1]
    return default


async def entrar_modulo(page, nombre: str):
    """
    Entra a un módulo desde el dashboard. SSASUR es una SPA de UNA pestaña:
    la tarjeta es un <button>; al clicarlo, ESTA misma pestaña navega al
    módulo (www.ssasur.cl/<modulo>) y se acuña su sesión. No abre tab nueva.
    """
    btn = page.locator(f'button:has-text("{nombre}")').first
    await btn.wait_for(state="visible", timeout=30_000)
    await btn.click()
    # Esperar a que la SPA navegue fuera del dashboard (login.ssasur.cl → www.ssasur.cl)
    try:
        await page.wait_for_url(lambda u: "login.ssasur.cl" not in u, timeout=30_000)
    except Exception:
        pass
    await _networkidle(page)
    await page.wait_for_timeout(2_000)
    print(f"  ✓ En módulo {nombre}: {page.url}")


async def descargar(page, dest_dir: Path, accion, etiqueta: str, timeout=TIMEOUT_DESCARGA):
    """Envuelve un clic que dispara una descarga y guarda el archivo."""
    async with page.expect_download(timeout=timeout) as dl_info:
        await accion()
    dl   = await dl_info.value
    dest = dest_dir / dl.suggested_filename
    await dl.save_as(dest)
    size_kb = dest.stat().st_size // 1024
    print(f"  ✓ {dl.suggested_filename}  ({size_kb:,} KB)")
    return dest


async def descargar_como(page, dest_path: Path, accion, timeout=TIMEOUT_DESCARGA):
    """Como descargar(), pero guarda con un nombre FIJO (dest_path)."""
    dest_path.parent.mkdir(parents=True, exist_ok=True)
    async with page.expect_download(timeout=timeout) as dl_info:
        await accion()
    dl = await dl_info.value
    await dl.save_as(dest_path)
    print(f"  ✓ {dest_path.name}  ({dest_path.stat().st_size // 1024:,} KB)")
    return dest_path


def _limpiar_versiones_previas(dest_dir: Path, patron: str, mantener: Path):
    """Borra copias anteriores de un reporte con nombre fechado por SSASUR
    (mismo patrón glob), conservando solo `mantener` (la recién descargada).
    maestro_aa.py y programacion_aa.py SIEMPRE usan la copia más reciente por
    mtime (glob + max) — las versiones viejas nunca se vuelven a leer y solo
    se acumulan como basura en la carpeta del proyecto, run tras run."""
    borrados = 0
    for f in dest_dir.glob(patron):
        if f.resolve() == mantener.resolve():
            continue
        try:
            f.unlink()
            borrados += 1
        except Exception:
            pass
    if borrados:
        print(f"  🧹 {borrados} copia(s) anterior(es) de '{patron}' eliminada(s)")


# ── Entrada CORRECTA al módulo RECETA + descarga del informe sabana ─────────────
# CLAVE (descubierto por captura HTTP del flujo manual): el informe sabana sale
# VACÍO (0 filas) si la sesión NO está acuñada para el PROYECTO (629), aunque el
# formulario cargue y jsonValidaSabana responda {status:true}. Ni el deep-link
# directo a /receta/informes/sabana ni entrar por RCE (araucaniasur/rce/index.php)
# acuñan el proyecto. SOLO clicar la tarjeta RECETA del dashboard lo hace, porque
# dispara el puente SSO (/sesion/obtener con data_proyecto=629) que deja la
# pestaña en www.ssasur.cl/receta con la sesión correcta. Con esa sesión, el
# informe se baja por POST directo (ni calendario ni firma hacían falta nunca).
PROYECTO_RECETA  = 629          # informativo: proyecto que acuña la sesión receta
ANCLA_HISTORICO  = date(2026, 6, 1)   # primer día del histórico de recetas
BLOQUE_DIAS      = 30           # tope del servidor: máx. 30 días por consulta


async def _networkidle(page, timeout=15_000):
    """Espera networkidle con tolerancia: varios modulos de SSASUR mantienen
    alguna conexion de fondo que nunca llega a estar realmente idle, asi que
    un timeout aqui no es un error real de navegacion (el wait_for_timeout()
    de resguardo que sigue a cada llamado ya cubre el asentamiento visual)."""
    try:
        await page.wait_for_load_state("networkidle", timeout=timeout)
    except Exception:
        pass


async def entrar_receta(page):
    """Dashboard → tarjeta RECETA → (modal proyecto) → www.ssasur.cl/receta.
    Si ya estamos dentro del módulo receta, vuelve al inicio del módulo sin
    pasar por el dashboard (evita el timeout del botón RECETA)."""
    if "www.ssasur.cl/receta" in page.url:
        try:
            await page.goto("https://www.ssasur.cl/receta")
            await _networkidle(page)
        except Exception:
            pass
        await page.wait_for_timeout(1_000)
        print(f"  ✓ En módulo receta: {page.url}")
        return

    # Clic en la tarjeta RECETA (texto EXACTO para no pegarle a "RCE - REGISTRO ...").
    try:
        await page.get_by_text("RECETA", exact=True).first.click(timeout=15_000)
    except Exception:
        await page.click('button:has-text("RECETA")', timeout=15_000)
    await page.wait_for_timeout(2_000)

    # Si aparece el modal de selección de proyecto, aceptarlo.
    for sel in ('button:has-text("Aceptar e ingresar")',
                'a:has-text("Aceptar e ingresar")',
                'button:has-text("Aceptar")'):
        try:
            await page.click(sel, timeout=3_000)
            break
        except Exception:
            continue

    # Esperar a estar en www.ssasur.cl/receta (no en araucaniasur/rce).
    try:
        await page.wait_for_function(
            "location.href.includes('www.ssasur.cl/receta')", timeout=30_000
        )
    except Exception:
        pass
    try:
        await _networkidle(page)
    except Exception:
        pass
    await page.wait_for_timeout(1_500)
    print(f"  ✓ En módulo receta: {page.url}")


async def descargar_sabana(page, fi: str, ft: str) -> str:
    """
    Descarga el informe completo de recetas por POST directo (con la sesión ya
    acuñada al proyecto). Replica los dos POST exactos del flujo manual:
      1) /receta/informes/jsonValidaSabana   (valida; fija contexto)
      2) .../informeCompletoReceta.php        (devuelve el CSV ';' latin-1)
    fi/ft en formato dd/mm/yyyy. Devuelve el CSV como texto latin-1.
    Timeout de 150 s por bloque (AbortSignal en JS + asyncio.wait_for de respaldo).
    """
    coro = page.evaluate(
        r"""async ({fi, ft}) => {
          const enc = encodeURIComponent;
          const csrf = (document.querySelector('input[name=csrf_token]')||{}).value || '';
          try {
            await fetch('https://www.ssasur.cl/receta/informes/jsonValidaSabana', {
              method:'POST',
              headers:{'content-type':'application/x-www-form-urlencoded; charset=UTF-8','x-requested-with':'XMLHttpRequest'},
              body:`csrf_token=${csrf}&fechaInicio=${enc(fi)}&fechaTermino=${enc(ft)}&estadoSolicitadoHidden=F`,
              credentials:'include',
              signal: AbortSignal.timeout(120_000),
            });
          } catch(e){}
          const r = await fetch('https://www.ssasur.cl/application/sistemas/receta/reportes/informeCompletoReceta.php', {
            method:'POST', headers:{'content-type':'application/x-www-form-urlencoded'},
            body:`fechaInicio=${enc(fi)}&fechaTermino=${enc(ft)}&estadoSolicitadoHidden=F&entregaDomicilio=F&estadoAnulada=N`,
            credentials:'include',
            signal: AbortSignal.timeout(120_000),
          });
          const buf = await r.arrayBuffer();
          const b = new Uint8Array(buf); let s='';
          for (let i=0;i<b.length;i++) s += String.fromCharCode(b[i]);
          return s;
        }""",
        {"fi": fi, "ft": ft},
    )
    return await asyncio.wait_for(coro, timeout=150)


# ── GT · Modalidad de despacho (entra por la tarjeta RECETA, igual que la sabana) ─
async def _dump_formulario(page, etiqueta="form"):
    """Imprime los <select> (id/name + opciones) y los botones/enlaces visibles.
    Sirve para DESCUBRIR los selectores/controles reales la 1ª vez en vivo: revisa
    la salida y ajusta SELS_BUSCAR / SELS_EXCEL (y confirma el control Origen) arriba."""
    try:
        info = await page.evaluate(r"""() => {
          const sels = [...document.querySelectorAll('select')].map(s => ({
            id: s.id, name: s.name,
            opts: [...s.options].slice(0, 30).map(o => `${o.value}=${(o.textContent||'').trim()}`)
          }));
          const inps = [...document.querySelectorAll('input')]
            .filter(i => !['button','submit','image','hidden'].includes(i.type))
            .map(i => ({id: i.id, name: i.name, type: i.type,
                        value: i.value, ph: i.placeholder || '', checked: i.checked}));
          const btns = [...document.querySelectorAll('button,a,input[type=button],input[type=submit]')]
            .map(b => (b.textContent || b.value || '').trim()).filter(Boolean).slice(0, 60);
          return {sels, inps, btns, url: location.href};
        }""")
    except Exception as e:
        print(f"  [dump] no se pudo inspeccionar: {e}")
        return
    print(f"  [DESCUBRIR · {etiqueta}] {info['url']}")
    for s in info["sels"]:
        print(f"    <select id='{s['id']}' name='{s['name']}'>  {s['opts']}")
    for i in info["inps"]:
        chk = " checked" if i.get("checked") else ""
        print(f"    <input id='{i['id']}' name='{i['name']}' type='{i['type']}'{chk}> "
              f"value='{i['value']}' ph='{i['ph']}'")
    print(f"    botones/enlaces: {info['btns']}")


async def _hay_form_gt(page):
    """Heurística de 'el reporte de gestión territorial cargó': URL del reporte,
    o un input de fecha, o aparecen las palabras origen/destino en la página."""
    return await page.evaluate(r"""() => {
      if (/gestionterritorial/i.test(location.href)) return true;
      if (document.querySelector('input[type=date]')) return true;
      const txt = (document.body.innerText || '').toLowerCase();
      return /origen/.test(txt) && /destino/.test(txt);
    }""")


async def _abrir_reporte_gt(page):
    """Abre el reporte de gestión territorial dentro de RECETA: prueba el deep-link
    y, si no aparece el formulario, navega por el menú Reporte → Gestión Territorial
    (o 'Modalidad de despacho'). Devuelve True si el formulario cargó."""
    try:
        await page.goto(GESTION_TERRITORIAL_URL)
        await _networkidle(page)
        await page.wait_for_timeout(1_500)
    except Exception:
        pass
    if await _hay_form_gt(page):
        return True
    # El menú real (verificado en DOM 18-06-2026) es "Reportes" → "Informe Modalidad Despacho"
    print("  [info] Deep-link sin formulario — navego por el menú Reportes → Informe Modalidad Despacho")
    for sel in ('a:has-text("Reportes")', 'button:has-text("Reportes")',
                'a:has-text("Reporte")', 'button:has-text("Reporte")',
                'li:has-text("Reportes")', 'span:has-text("Reportes")'):
        try:
            await page.click(sel, timeout=3_000)
            await page.wait_for_timeout(700)
            break
        except Exception:
            continue
    for sel in ('a:has-text("Informe Modalidad Despacho")',
                'a:has-text("Modalidad Despacho")',
                'a:has-text("Gestión Territorial")', 'a:has-text("Gestion Territorial")',
                'a:has-text("Modalidad de despacho")', ':text("Informe Modalidad Despacho")',
                ':text("Gestión Territorial")', ':text("Modalidad de despacho")'):
        try:
            await page.click(sel, timeout=3_000)
            break
        except Exception:
            continue
    await _networkidle(page)
    await page.wait_for_timeout(1_500)
    return await _hay_form_gt(page)


async def _marcar_origen(page):
    """Marca SIEMPRE 'Origen' en el control Origen/Destino (radio, checkbox, select
    o pestaña). Devuelve una etiqueta de lo que marcó, o None si no lo encontró."""
    return await page.evaluate(r"""() => {
      const isOrigen = s => /origen/i.test(s || '');
      // 1) radios / checkboxes (por label, value, id o name)
      for (const r of document.querySelectorAll('input[type=radio], input[type=checkbox]')) {
        const lbl  = (r.labels && r.labels[0]) ? r.labels[0].textContent : '';
        const meta = `${lbl} ${r.value || ''} ${r.id || ''} ${r.name || ''}`;
        if (isOrigen(meta)) {
          if (!r.checked) r.click();
          r.dispatchEvent(new Event('change', {bubbles: true}));
          return 'radio:' + (r.id || r.name || r.value || 'origen');
        }
      }
      // 2) select con opción Origen
      for (const s of document.querySelectorAll('select')) {
        const o = [...s.options].find(o => isOrigen(o.textContent));
        if (o) {
          s.value = o.value;
          s.dispatchEvent(new Event('change', {bubbles: true}));
          return 'select:' + (s.id || s.name || 'origen');
        }
      }
      // 3) botón / pestaña / label con texto Origen
      for (const b of document.querySelectorAll('button, a, label, [role=tab], .tab, .nav-link')) {
        if (isOrigen(b.textContent) && b.textContent.trim().length < 30) {
          b.click();
          return 'tab:' + b.textContent.trim().slice(0, 20);
        }
      }
      return null;
    }""")


async def _click_primero(page, selectores, etiqueta, force=False):
    """Clic en el primer selector que exista. Lanza si ninguno aparece."""
    for sel in selectores:
        try:
            await page.click(sel, timeout=4_000, force=force)
            return
        except Exception:
            continue
    raise RuntimeError(f"No encontré el botón '{etiqueta}' (probé: {selectores})")


async def _set_fechas(page, desde, hasta, id_ini=SEL_FECHA_INI, id_fin=SEL_FECHA_FIN):
    """Rellena fechaInicio/fechaTermino simulando tipeo real (click + seleccionar
    todo + escribir carácter por carácter + Tab). Necesario porque asignar
    .value vía JS y disparar eventos sintéticos (el método anterior) dejaba
    fechaInicio VACÍO de forma intermitente — visto en vivo 14-07-2026: el
    campo mostraba value='' pese al assignment, mientras fechaTermino sí
    tomaba el valor. El campo tiene máscara/validación propia que solo
    reacciona a eventos de teclado reales, no a Event() sintéticos en bloque.
    Cae a la heurística JS anterior si el campo por id no existe. Fechas en
    dd/mm/yyyy. `id_ini`/`id_fin` permiten reutilizar esto con formularios que
    usan otros ids (p.ej. Agenda Médica: FechaInicio/FechaTermino, distintos
    de fechaInicio/fechaTermino de GT/Controlados). El Tab se manda vía
    loc.press() (no page.keyboard) para que funcione igual si `page` es en
    realidad un Frame (sin atributo .keyboard)."""
    async def _escribir(field_id, valor):
        try:
            loc = page.locator(f"#{field_id}")
            if await loc.count() == 0:
                return False
            await loc.click()
            await loc.press("Control+A")
            await loc.press("Delete")
            await loc.type(valor, delay=30)
            await loc.press("Tab")
            await page.wait_for_timeout(250)
            real = await loc.input_value()
            return real == valor
        except Exception:
            return False

    ok_ini = await _escribir(id_ini, desde) if desde else True
    ok_fin = await _escribir(id_fin, hasta) if hasta else True
    if ok_ini and ok_fin:
        return

    try:
        await page.evaluate(r"""({ini, fin, desde, hasta}) => {
          const fire = el => ['input','change','blur','keyup']
              .forEach(ev => el.dispatchEvent(new Event(ev, {bubbles: true})));
          const setById = (id, v) => {
            const el = document.getElementById(id);
            if (el && v) { el.value = v; fire(el); return true; }
            return false;
          };
          let ok = setById(ini, desde);
          ok = setById(fin, hasta) || ok;
          if (!ok) {   // fallback: por nombre/placeholder
            const ins = [...document.querySelectorAll('input')].filter(i => {
              const k = ((i.name||'') + (i.id||'') + (i.placeholder||'')).toLowerCase();
              return i.type !== 'hidden' && (i.type === 'date' || /fecha|dia/.test(k));
            });
            if (ins[0] && desde) { ins[0].value = desde; fire(ins[0]); }
            if (ins[1] && hasta) { ins[1].value = hasta; fire(ins[1]); }
          }
        }""", {"ini": SEL_FECHA_INI, "fin": SEL_FECHA_FIN, "desde": desde, "hasta": hasta})
    except Exception:
        pass


async def _filas_resultado(page):
    """Heurística de cuántos pacientes trae el resultado: filas de la tabla más
    grande. 0 si hay mensaje de 'sin datos'; None si no hay tabla (desconocido)."""
    return await page.evaluate(r"""() => {
      const body = (document.body.innerText || '').toLowerCase();
      if (/no se encontraron|sin datos|no hay (registros|datos|resultados)|0 registros/.test(body)) return 0;
      const tablas = [...document.querySelectorAll('table')];
      if (!tablas.length) return null;
      let max = 0;
      for (const t of tablas) max = Math.max(max, t.querySelectorAll('tbody tr').length);
      return max;
    }""")


def _contar_filas_xlsx(path):
    """Nº de filas de datos del Excel descargado (descuenta título + encabezado).
    Best-effort; -1 si no se puede leer."""
    try:
        import openpyxl
        wb = openpyxl.load_workbook(path, read_only=True)
        ws = wb.worksheets[0]
        n = sum(1 for r in ws.iter_rows(values_only=True) if any(c is not None for c in r))
        wb.close()
        return max(n - 2, 0)   # fila de título + fila de encabezado
    except Exception:
        return -1


async def paso_gt(page, desde=None, hasta=None, debug=False):
    """RECETA → Reporte → Gestión Territorial (informe de modalidad de despacho).
    Define las fechas, marca SIEMPRE 'Origen' en Origen/Destino, espera el listado
    de recetas y baja el Excel a GT_DIR. Devuelve (archivo|None, n_filas):
    n = 0 → sin recetas (no descarga); n = -1 → error. `debug` imprime los volcados
    [DESCUBRIR …] y guarda screenshots para inspeccionar el formulario."""
    GT_DIR.mkdir(parents=True, exist_ok=True)
    print("\n[GT] Módulo RECETA — Reporte de gestión territorial (modalidad de despacho)")
    await entrar_receta(page)
    if not await _abrir_reporte_gt(page):
        # Justo tras un login recién restaurado el SPA puede no haber terminado
        # de hidratar — un solo reintento con más espera basta en ese caso
        # (mismo patrón que el retry de sesión expirada del paso 1).
        print("  [AVISO] Formulario GT no cargó al primer intento — reintento en 3s...")
        await page.wait_for_timeout(3_000)
    if not await _hay_form_gt(page) and not await _abrir_reporte_gt(page):
        print("  [ERROR] No cargó el formulario de gestión territorial.")
        await _dump_formulario(page, "gt-fallo")
        await page.screenshot(path=str(MAESTRO_DIR / "debug_gt.png"))
        return (None, -1)
    if debug:
        await _dump_formulario(page, "gestion-territorial")

    # 1) Fechas (solo si se pasaron; si no, se respeta el default del formulario).
    if desde or hasta:
        await _set_fechas(page, desde or hasta, hasta or desde)
        print(f"  Fechas: {desde or hasta} → {hasta or desde}")
        await page.wait_for_timeout(800)

    # 2) Origen/Destino → SIEMPRE Origen (radio tipoEstablecimiento=0).
    marca = await _marcar_origen(page)
    print(f"  Origen/Destino → ORIGEN ({marca})" if marca
          else "  [AVISO] No encontré el control Origen/Destino.")
    await page.wait_for_timeout(800)

    # 3) Buscar y esperar el listado (DataTable tablaGestionTerritorial).
    try:
        await _click_primero(page, SELS_BUSCAR, "Buscar")
    except Exception:
        pass   # por si el listado cargara solo al marcar origen
    await _networkidle(page)
    await page.wait_for_timeout(2_500)
    if debug:
        await _dump_formulario(page, "post-buscar")
        await page.screenshot(path=str(MAESTRO_DIR / "debug_gt.png"))

    n = await _filas_resultado(page)
    if n == 0:
        # Reintento: puede ser el mismo tipo de condición de carrera del
        # llenado de fechas ya visto en vivo (14-07-2026) — a veces el filtro
        # no alcanza a aplicarse antes de leer la tabla y devuelve "sin datos"
        # aunque SSASUR sí tenga despachos para el rango (detectado 30-07-2026:
        # 29/07 quedó vacío una corrida y con 64 filas la siguiente, mismo rango).
        print("  (0 filas — reintentando por si el filtro no aplicó a tiempo...)")
        await page.wait_for_timeout(2_000)
        try:
            await _click_primero(page, SELS_BUSCAR, "Buscar")
        except Exception:
            pass
        await _networkidle(page)
        await page.wait_for_timeout(2_500)
        n = await _filas_resultado(page)
    if n == 0:
        print("  (sin recetas en el listado — no hay despacho para esas fechas)")
        return (None, 0)

    # 4) Excel (la firma electrónica NO es necesaria para este reporte).
    d = (desde or hasta or "").replace("/", "-")
    h = (hasta or "").replace("/", "-")
    suf = (f"_{d}" + (f"_{h}" if h and h != d else "")) if d else ""
    dest = GT_DIR / f"reporteGestionTerritorial{suf}.xlsx"
    try:
        await descargar_como(page, dest,
                             lambda: _click_primero(page, SELS_EXCEL, "Excel", force=True))
    except Exception as e:
        print(f"  [ERROR] No se pudo bajar el Excel: {e}")
        await page.screenshot(path=str(MAESTRO_DIR / "debug_gt.png"))
        return (None, -1)
    return (dest, _contar_filas_xlsx(dest))


async def _seleccionar_bodega_at_abierta(page):
    """Selecciona 'FARMACIA AT ABIERTA' en el <select> Bodega del informe de
    controlados. Busca por TEXTO de la opción (no por id/value fijo, porque no
    se conoce de antemano) — mismo patrón que _marcar_origen()."""
    return await page.evaluate(r"""() => {
      for (const s of document.querySelectorAll('select')) {
        const o = [...s.options].find(o => /farmacia\s*at\s*abierta/i.test(o.textContent || ''));
        if (o) {
          if (s.value !== o.value) {
            s.value = o.value;
            s.dispatchEvent(new Event('change', {bubbles: true}));
          }
          return {sel: s.id || s.name || 'bodega', val: o.value, label: (o.textContent || '').trim()};
        }
      }
      return null;
    }""")


async def paso_controlados(page, fecha: str, debug=False):
    """RECETA → Reportes → Informe Medicamentos Controlados. Bodega = FARMACIA
    AT ABIERTA, Medicamento vacío (trae todos los controlados), misma fecha en
    Inicio/Término (día hábil anterior). Devuelve (archivo|None, n_filas):
    n = 0 → sin despachos ese día (no descarga); n = -1 → error."""
    CONTROLADOS_DIR.mkdir(parents=True, exist_ok=True)
    print(f"\n[Controlados] Informe Medicamentos Controlados — Farmacia AT Abierta ({fecha})")
    await entrar_receta(page)
    await page.goto(CONTROLADOS_URL)
    await _networkidle(page)
    await page.wait_for_timeout(1_500)

    if not await page.evaluate("() => !!document.querySelector('select')"):
        print("  [ERROR] No cargó el formulario de controlados.")
        await _dump_formulario(page, "controlados-fallo")
        await page.screenshot(path=str(MAESTRO_DIR / "debug_controlados.png"))
        return (None, -1)
    if debug:
        await _dump_formulario(page, "controlados")

    bodega = await _seleccionar_bodega_at_abierta(page)
    print(f"  Bodega → FARMACIA AT ABIERTA ({bodega})" if bodega
          else "  [AVISO] No encontré el select de Bodega — sigo con el default del formulario.")
    await page.wait_for_timeout(500)

    await _set_fechas(page, fecha, fecha)
    print(f"  Fecha: {fecha} (día hábil anterior)")
    await page.wait_for_timeout(500)

    try:
        await _click_primero(page, SELS_BUSCAR, "Buscar")
    except Exception as e:
        print(f"  [ERROR] No encontré el botón Buscar: {e}")
        await _dump_formulario(page, "controlados-sin-buscar")
        await page.screenshot(path=str(MAESTRO_DIR / "debug_controlados.png"))
        return (None, -1)
    await _networkidle(page)
    await page.wait_for_timeout(2_000)
    if debug:
        await _dump_formulario(page, "controlados-post-buscar")
        await page.screenshot(path=str(MAESTRO_DIR / "debug_controlados.png"))

    n = await _filas_resultado(page)
    if n == 0:
        print("  (sin despachos de controlados ese día — no hay Excel que bajar)")
        return (None, 0)

    dest = CONTROLADOS_DIR / f"Informe_Controlados_AT_Abierta_{fecha.replace('/', '-')}.xlsx"
    try:
        await descargar_como(page, dest,
                             lambda: _click_primero(page, SELS_EXCEL, "Excel", force=True))
    except Exception as e:
        print(f"  [ERROR] No se pudo bajar el Excel: {e}")
        await page.screenshot(path=str(MAESTRO_DIR / "debug_controlados.png"))
        return (None, -1)
    return (dest, _contar_filas_xlsx(dest))


async def main():
    no_pause = "--no-pause" in sys.argv
    # ── Verificar Playwright ───────────────────────────────────────────────────
    try:
        from playwright.async_api import async_playwright
    except ImportError:
        print("\n[ERROR] Playwright no está instalado.")
        print("  Ejecuta AUTO_SSASUR.bat para instalarlo automáticamente.\n")
        if not no_pause:
            try:
                input("Presiona Enter para cerrar...")
            except EOFError:
                pass
        sys.exit(1)

    # Modo prueba: descarga SOLO recetas y termina (sin stock, maestro ni publicar).
    solo_recetas = "--solo-recetas" in sys.argv
    # No publicar en GitHub al final (útil para corridas de prueba/debug). SOLO
    # frena GitHub — PASO 7-9 sigue copiando al Escritorio y subiendo a Drive
    # con datos reales pase lo que pase (ver nota en SINCRONIZAR_TODO.bat más
    # abajo). No asumir que "no_publicar" == "no toca nada externo".
    no_publicar  = "--no-publicar" in sys.argv
    # Modo GT exclusivo: solo gestión territorial y termina (sin sábana/stock/maestro).
    gt_mode  = ("--gt" in sys.argv) or ("--solo-gt" in sys.argv)
    # Solo stock: salta recetas y GT, baja solo stock y corre maestro.
    solo_stock = "--solo-stock" in sys.argv
    # En corrida completa, saltar GT con --no-gt (si ya se corrió hoy por separado).
    no_gt    = "--no-gt" in sys.argv or solo_stock
    # Saltar el registro ISP de recetas cheque con --no-rch.
    no_rch   = "--no-rch" in sys.argv
    # Consolidado de hemogramas Clozapina: OPT-IN (--clozapina) hasta hacer un
    # ensayo supervisado de clozapina_hce_hemogramas.py en vivo — no corre por
    # defecto todavía.
    con_clozapina = "--clozapina" in sys.argv
    # Saltar la descarga del reporte de Programación AA con --no-programacion.
    no_programacion = "--no-programacion" in sys.argv
    # Saltar el informe de medicamentos controlados con --no-controlados.
    no_controlados = "--no-controlados" in sys.argv or solo_stock
    debug_gt = "--debug-gt" in sys.argv        # volcados [DESCUBRIR …] + screenshots
    # Servicios Farmacéuticos (Agenda Médica): dispara solo en los primeros días
    # hábiles del mes (catch-up si el 1er día hábil no corrió), o forzado con
    # --servicios-farmaceuticos (ignora el gatillo — sirve para la 1ª prueba
    # supervisada y para reprocesar un mes puntual con --mes-servicios).
    forzar_servicios = "--servicios-farmaceuticos" in sys.argv
    no_servicios = "--no-servicios-farmaceuticos" in sys.argv
    debug_agenda = "--debug-agenda" in sys.argv or debug_gt
    _mes_servicios_arg = _arg_val("--mes-servicios")   # "MM-AAAA", opcional
    _fecha   = _arg_val("--fecha")             # atajo: mismo día en desde/hasta
    today = date.today()
    ayer  = today - timedelta(days=1)
    # Informe Medicamentos Controlados: día hábil anterior a HOY (no a "ayer"
    # calendario) — si hoy es lunes, cae en viernes; si hubo feriado, lo salta.
    # Override para reintentos/backfill: --fecha-controlados dd/mm/yyyy.
    fecha_controlados = _arg_val("--fecha-controlados", fmt(dia_habil_anterior(today)))
    # Default GT: último día hábil (no "ayer" calendario) → +13 días. Captura
    # despachos próximos programados (el reporte muestra pendientes, no
    # histórico entregado) — los despachos pasados desaparecen del listado al
    # ser procesados, por eso se mira al FUTURO en lugar de hacia atrás.
    # Bug real 30-07-2026 (caso Teodoro Schmidt): con "ayer" a secas, un lunes
    # arranca en domingo y nunca llega a ver viernes/sábado — cualquier cosa
    # que se digitó Y despachó ese fin de semana (sin corrida de por medio)
    # queda fuera de la ventana consultada. dia_habil_anterior() ya salta
    # fines de semana Y feriados, así que un lunes normal arranca en viernes
    # solo; si el viernes fue feriado, arranca más atrás. Sin cambio en el
    # resto de la semana (dia_habil_anterior(martes) = lunes = ayer de todos
    # modos). Complementa a verificar_backlog_gt(), que cubre el caso — más
    # grave — de lo que ya se despachó y desapareció del reporte antes de
    # que corriera cualquier ventana.
    # Si la última corrida exitosa (ULTIMA_CORRIDA_JSON) es más vieja que el
    # último día hábil — porque se saltó un día hábil normal, no solo fin de
    # semana/feriado — se extiende el inicio hasta ahí para no dejar un hueco
    # real sin consultar (ver ultima_corrida_ok()/rango_backlog_gt()).
    _gt_inicio_base = dia_habil_anterior(today)
    _ultima_ok = ultima_corrida_ok()
    if _ultima_ok is not None and _ultima_ok + timedelta(days=1) < _gt_inicio_base:
        _gt_inicio_base = _ultima_ok + timedelta(days=1)
    # Se puede acotar con --desde dd/mm/yyyy o --fecha dd/mm/yyyy.
    _gt_inicio = _fecha or fmt(_gt_inicio_base)
    desde_gt = _arg_val("--desde", _gt_inicio)
    hasta_gt = _arg_val("--hasta", _fecha or fmt(today + timedelta(days=13)))

    print()
    print("═" * 62)
    print("  AUTO SSASUR  ·  Maestro AA Farmacia")
    print(f"  Recetas: bloques de 30 días desde {fmt(ANCLA_HISTORICO)} hasta {fmt(ayer)}")
    print("═" * 62)

    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=False,
            args=["--start-maximized", "--disable-blink-features=AutomationControlled"],
        )

        if SESSION_FILE.exists():
            context = await browser.new_context(
                accept_downloads=True,
                storage_state=str(SESSION_FILE),
            )
            print("\n  (Sesión guardada encontrada — puede que no necesites logarte)")
        else:
            context = await browser.new_context(accept_downloads=True)

        page = await context.new_page()

        # ── PASO 1 — DETECTAR LOGIN ────────────────────────────────────────────
        print("\n[1/9] Logéate en SSASUR (tienes 5 minutos)...")
        await page.goto(DASHBOARD_URL)

        async def _esperar_abastecimiento(pg):
            # Antes esto era una sola espera bloqueante de hasta 5 min sin
            # imprimir nada — si el dashboard tardaba en cargar era
            # indistinguible de un freeze real. Se poll-ea en tramos cortos
            # con heartbeat para que la consola muestre que sigue vivo.
            LATIDO = 20_000  # ms entre heartbeats
            transcurrido = 0
            while transcurrido < TIMEOUT_LOGIN:
                try:
                    await pg.wait_for_selector(
                        'button:has-text("ABASTECIMIENTO"), div:has-text("ABASTECIMIENTO")',
                        timeout=LATIDO,
                    )
                    return
                except Exception:
                    pass
                try:
                    if await pg.evaluate("document.body.innerText.includes('ABASTECIMIENTO')"):
                        return
                except Exception:
                    pass  # página navegando/cerrada momentáneamente — reintenta
                transcurrido += LATIDO
                restante = max(TIMEOUT_LOGIN - transcurrido, 0) // 1000
                print(f"  … esperando login/dashboard ({restante}s restantes)")
            raise TimeoutError(f"No se detectó ABASTECIMIENTO tras {TIMEOUT_LOGIN // 1000}s")

        try:
            await _esperar_abastecimiento(page)
        except Exception as _login_err:
            from playwright._impl._errors import TargetClosedError as _TCE
            if isinstance(_login_err, _TCE) and SESSION_FILE.exists():
                # Sesión expirada: las cookies vencidas provocaron un redirect
                # que cerró la página — reiniciar sin sesión guardada
                print("  [aviso] Sesión guardada expirada — abre el browser y logéate.")
                SESSION_FILE.unlink(missing_ok=True)
                await context.close()
                context = await browser.new_context(accept_downloads=True)
                page = await context.new_page()
                await page.goto(DASHBOARD_URL)
                await _esperar_abastecimiento(page)
            else:
                raise

        await context.storage_state(path=str(SESSION_FILE))
        print("  ✓ Sesión detectada")

        # ── MODO GT — modalidad de despacho (gestión territorial) ───────────────
        # Independiente del maestro: baja el/los Excel de despacho por
        # establecimiento de destino y termina. Uso:
        #   py AUTO_SSASUR.py --gt                  (todos los destinos)
        #   py AUTO_SSASUR.py --gt --estab FREIRE   (solo ese destino)
        #   py AUTO_SSASUR.py --gt --fecha 17/06/2026
        if gt_mode:
            dest, n = await paso_gt(page, desde_gt, hasta_gt, debug_gt)
            await browser.close()
            print("\n" + "═" * 62)
            print("  Informe Modalidad Despacho (GT) — resumen")
            if dest:
                cnt = f"{n} recetas" if isinstance(n, int) and n >= 0 else "recetas ?"
                print(f"    ✓ {dest.name}  ({cnt})")
                print(f"  Archivo: {dest}")
                print(f"  Carpeta: {GT_DIR}")
                print("═" * 62)
                # ── Cruce + planillas automático ──────────────────────────────
                cruce = MAESTRO_DIR / "cruce_gt.py"
                out_gt = gt_salida(dest)   # carpeta propia por rango de fechas
                print(f"\n[GT] Cruzando con histórico y generando planillas...")
                print(f"     Salida: out_gt/{out_gt.name}/")
                if cruce.exists():
                    ret = subprocess.run(
                        [sys.executable, str(cruce), str(dest),
                         "--salida", str(out_gt), "--generar"],
                        env={**os.environ, "PYTHONUTF8": "1", "PYTHONIOENCODING": "utf-8"},
                    )
                    if ret.returncode != 0:
                        print(f"  [aviso] cruce_gt.py terminó con código {ret.returncode}")
                else:
                    print(f"  [aviso] cruce_gt.py no encontrado — omitiendo cruce.")
            elif n == 0:
                print("    · Sin recetas en el listado — no hay despacho para esas fechas.")
                print(f"  Carpeta: {GT_DIR}")
                print("═" * 62)
            else:
                print("    [ERROR] No se generó el Excel — revisa debug_gt.png y el dump [DESCUBRIR].")
                print(f"  Carpeta: {GT_DIR}")
                print("═" * 62)
            return

        # ════════════════════════════════════════════════════════════════════
        #  PASO 2 — RECETA  (entrada correcta por la tarjeta → informe sabana)
        # ════════════════════════════════════════════════════════════════════
        # El server limita cada consulta a 30 días → se baja por BLOQUES de 30
        # días desde ANCLA_HISTORICO (01/06/2026) hasta ayer. Cada bloque:
        #   · mientras no cubra los 30 días → archivo PARCIAL `..._b<inicio>.csv`,
        #     que se REESCRIBE en cada corrida (01/06→ayer va creciendo);
        #   · al cubrir los 30 días → se SELLA como `..._b<inicio>_FULL.csv` y se
        #     BORRA el parcial; un bloque sellado ya no se vuelve a descargar;
        #   · un rango sin datos no deja archivo (se borra el incompleto).
        # El maestro lee todos los informe_completo_recetas*.csv y deduplica por
        # ID Receta Detalle, así que los bloques (rangos sin solape) no doble-cuentan.
        if solo_stock:
            print("\n[2/9] Recetas — omitido (--solo-stock).")
        else:
            print("\n[2/9] Módulo RECETA — informe completo por bloques de 30 días...")
            await entrar_receta(page)

            await page.goto(RECETA_INFORME)
            await _networkidle(page)
            await page.wait_for_timeout(1_500)

            if not await page.evaluate("() => !!document.getElementById('fechaInicio')"):
                print("  [AVISO] No cargó el formulario sabana — omito recetas.")
                await page.screenshot(path=str(MAESTRO_DIR / "debug_receta.png"))
            else:
                bstart = ANCLA_HISTORICO
                while bstart <= ayer:
                    bend  = bstart + timedelta(days=BLOQUE_DIAS - 1)   # bloque de 30 días (inclusive)
                    q_end = min(bend, ayer)
                    base    = f"informe_completo_recetas_b{bstart:%Y%m%d}"
                    parcial = MAESTRO_DIR / f"{base}.csv"        # rango aún sin completar 30 días
                    full    = MAESTRO_DIR / f"{base}_FULL.csv"   # bloque sellado (30 días completos)
                    # Bloque ya sellado → saltar (no se vuelve a descargar nunca).
                    if full.exists() and full.stat().st_size > 200:
                        bstart = bend + timedelta(days=1)
                        continue
                    # ¿Ya tenemos los 30 días del bloque disponibles (bend ≤ ayer)?
                    es_completo = (q_end == bend)
                    fi, ff = fmt(bstart), fmt(q_end)
                    print(f"  Bloque {fi} → {ff}  [{'completo → sella' if es_completo else 'parcial'}]")
                    try:
                        csv = await descargar_sabana(page, fi, ff)
                        cabecera = csv.lstrip()[:200].lower()
                        if "<html" in cabecera or "<!doctype" in cabecera or "iniciar sesión" in cabecera:
                            print(f"    [ERROR] Bloque {fi}→{ff}: la respuesta parece ser la página de "
                                  f"login (sesión expirada) — NO se guarda ni se sella. Reintenta la corrida.")
                            bstart = bend + timedelta(days=1)
                            continue
                        n_filas = max(sum(1 for l in csv.splitlines() if l.strip()) - 1, 0)
                        if n_filas > 0 or es_completo:
                            # Si el bloque está completo (30 días ya pasaron) se sella aunque
                            # tenga 0 filas, o se re-descargaría indefinidamente en cada corrida.
                            dest = full if es_completo else parcial
                            with open(dest, "w", encoding="latin-1", newline="") as fcsv:
                                fcsv.write(csv)
                            if es_completo:
                                parcial.unlink(missing_ok=True)   # borrar el incompleto al sellar
                            print(f"    ✓ {dest.name} · {dest.stat().st_size // 1024:,} KB · {n_filas:,} filas")
                        else:
                            parcial.unlink(missing_ok=True)       # rango sin datos: sin incompletos
                            print("    (0 filas en el rango — nada que guardar)")
                    except Exception as e:
                        print(f"    [AVISO] Falló el bloque {fi}→{ff}: {e}")
                    bstart = bend + timedelta(days=1)

        if solo_recetas:
            print("\n[solo-recetas] Modo prueba: omito stock, maestro y publicación.")
            await browser.close()
            return

        # ════════════════════════════════════════════════════════════════════
        #  PASO 3 — GT  (gestión territorial — mismo módulo RECETA, sesión 629)
        # ════════════════════════════════════════════════════════════════════
        gt_dest = None   # path del xlsx descargado; None si se omite o falla
        if no_gt:
            print("\n[3/9] Modalidad de Despacho (GT) — omitido (--no-gt).")
        else:
            print(f"\n[3/9] Modalidad de Despacho — Informe GT ({desde_gt} → {hasta_gt})...")
            try:
                gt_dest, n_gt = await paso_gt(page, desde_gt, hasta_gt, debug_gt)
            except Exception as e:
                # A diferencia de los demás pasos, esta llamada no estaba protegida:
                # un TargetClosedError (browser muerto a mitad de corrida, ver
                # auto_ssasur_error.log) tumbaba TODO el proceso acá y nunca se
                # llegaba a PASO 4/maestro_aa.py/publicación — se perdía el reporte
                # diario completo por un fallo solo del módulo GT (detectado
                # 31-08-2026: 3 días sin ningún reporte por esta causa).
                gt_dest, n_gt = None, None
                print(f"  [ERROR] Informe GT falló: {e} — continúo con el resto.")
            if gt_dest:
                cnt = f"{n_gt} recetas" if isinstance(n_gt, int) and n_gt >= 0 else "recetas ?"
                print(f"  ✓ {gt_dest.name}  ({cnt})")
                print(f"  → {gt_dest}")
            elif n_gt == 0:
                print(f"  · Sin recetas de despacho en {desde_gt} → {hasta_gt}.")
                print(f"    (Usa --desde dd/mm/yyyy --hasta dd/mm/yyyy para otro rango)")
            else:
                print("  [AVISO] No se generó el Excel GT — continúo con el resto.")
                print(f"    Revisa debug_gt.png en {MAESTRO_DIR}")

        # ── Chequeo de backlog GT (fin de semana / feriado) ────────────────────
        # El reporte de arriba solo ve PENDIENTES — una receta despachada un
        # día sin corrida (sábado, domingo, feriado) desaparece de ahí para
        # siempre sin dejar nómina. Se cruza el histórico real
        # (informe_completo_recetas) contra el maestro para no perderlas y,
        # si encuentra alguna, GENERA y REGISTRA la Nómina de Envío que
        # falta (ver rango_backlog_gt/verificar_backlog_gt/_generar_backlog_gt).
        # Corre siempre — no hace nada si no hay hueco/backlog que revisar,
        # pero cuando SÍ encuentra algo escribe en gt_maestro.xlsx y en
        # 04_Farmacia_Gestion_Territorial (ya no es de solo lectura).
        if not no_gt:
            try:
                verificar_backlog_gt(today)
            except Exception as e:
                print(f"  [AVISO] Chequeo de backlog GT falló: {e} — continúo con el resto.")

        # ════════════════════════════════════════════════════════════════════
        #  PASO 3b — INFORME MEDICAMENTOS CONTROLADOS (Farmacia AT Abierta)
        # ════════════════════════════════════════════════════════════════════
        if no_controlados:
            print("\n[3b/9] Informe Medicamentos Controlados — omitido (--no-controlados).")
        else:
            print(f"\n[3b/9] Informe Medicamentos Controlados (AT Abierta, {fecha_controlados})...")
            try:
                ctrl_dest, n_ctrl = await paso_controlados(page, fecha_controlados, debug_gt)
                if ctrl_dest:
                    cnt = f"{n_ctrl} filas" if isinstance(n_ctrl, int) and n_ctrl >= 0 else "filas ?"
                    print(f"  ✓ {ctrl_dest.name}  ({cnt})")
                    print(f"  → {ctrl_dest}")
                elif n_ctrl == 0:
                    print(f"  · Sin despachos de controlados en {fecha_controlados}.")
                else:
                    print("  [AVISO] No se generó el Excel de controlados — continúo con el resto.")
                    print(f"    Revisa debug_controlados.png en {MAESTRO_DIR}")
            except Exception as e:
                print(f"  [AVISO] Informe de controlados falló: {e} — continúo con el resto.")

        # ════════════════════════════════════════════════════════════════════
        #  PASO 4 — ABASTECIMIENTO  (volver al dashboard → entrar → stock)
        # ════════════════════════════════════════════════════════════════════
        print("\n[4/9] Módulo ABASTECIMIENTO...")
        bodega_ok = True
        try:
            await page.goto(DASHBOARD_URL)
            await _networkidle(page)
            await page.wait_for_timeout(1_500)
            await entrar_modulo(page, "ABASTECIMIENTO")

            await page.goto(STOCK_REPORTE)
            await _networkidle(page)
            await page.wait_for_timeout(2_500)

            if "login.ssasur.cl" in page.url:
                # Por si la sesión del módulo no quedó lista: reintentar una vez
                print("  [AVISO] Rebote al login — reintentando entrar a ABASTECIMIENTO...")
                await page.goto(DASHBOARD_URL)
                await _networkidle(page)
                await page.wait_for_timeout(1_500)
                await entrar_modulo(page, "ABASTECIMIENTO")
                await page.goto(STOCK_REPORTE)
                await _networkidle(page)
                await page.wait_for_timeout(2_500)
        except Exception as e:
            # Igual que paso_gt (ver fix 31-08-2026): la navegación de entrada no
            # estaba protegida — un fallo transitorio del paso anterior (ej. la
            # página quedó en chrome-error:// tras un ERR_CONNECTION_TIMED_OUT de
            # [3b/9]) hacía que este goto() chocara con una navegación aún en
            # curso y tumbara TODO el resto del pipeline (4b/4c/5-9, sin llegar
            # nunca a maestro_aa.py). Detectado en vivo el mismo día.
            print(f"  [ERROR] No se pudo entrar a ABASTECIMIENTO — omito stock: {e}")
            bodega_ok = False

        # Configurar el reporte: bodega = TODAS.
        # IMPORTANTE: NO seleccionar el establecimiento — hacerlo dispara un modal
        # ("Selección Proyecto") cuyo backdrop bloquea el botón Generar XLS.
        # El establecimiento ya viene en "PITRUFQUEN HOSP." (única opción).
        if bodega_ok:
            try:
                await page.select_option("#bodega", BODEGA_TODAS)
                await page.wait_for_timeout(1_000)
                valor_actual = await page.eval_on_selector("#bodega", "el => el.value")
                if valor_actual != BODEGA_TODAS:
                    raise RuntimeError(f"selector #bodega quedó en '{valor_actual}', no en TODAS ('{BODEGA_TODAS}')")
                print("  Bodega: TODAS")
            except Exception as e:
                print(f"  [ERROR] No se pudo seleccionar bodega TODAS — omito stock para no generar reporte parcial: {e}")
                await page.screenshot(path=str(MAESTRO_DIR / "debug_stock.png"))
                bodega_ok = False

        if bodega_ok:
            print("  Generando XLS de stock... (puede tardar varios minutos)")
            try:
                # force=True: el overlay "Generando reporte. Espere un momento."
                # puede interponerse al clic; force lo atraviesa.
                _dest_stock = await descargar(
                    page, MAESTRO_DIR,
                    lambda: page.click("#generarXLS_stock", force=True),
                    "ABASTECIMIENTO",
                )
                _limpiar_versiones_previas(MAESTRO_DIR, "reporte_de_stock_*.xlsx", _dest_stock)
            except Exception as e:
                print(f"  [ERROR] Descarga ABASTECIMIENTO falló: {e}")
                await page.screenshot(path=str(MAESTRO_DIR / "debug_stock.png"))
                print("  Screenshot guardado: debug_stock.png")

        # ════════════════════════════════════════════════════════════════════
        #  PASO 4b — PROGRAMACIÓN AA  (mismo módulo ABASTECIMIENTO, mes en curso)
        # ════════════════════════════════════════════════════════════════════
        # Consumido por programacion_aa.py (Cantidad Programada / Cantidad
        # Solicitada del ciclo Bodega AA). A diferencia del stock, este reporte
        # SÍ tiene <select> nativos (confirmado por inspección del DOM
        # 15-07-2026) — no hace falta clic+búsqueda por texto:
        #   #ano, #mes (1-12), #distribucion (5=FARMACOS),
        #   #select_proyecto (3=LOCAL), #cc (75=FARMACIA), #generarXLS_salida
        # "Mes en curso" = mes de hoy: la programación de meses futuros puede
        # existir con Cantidad Solicitada en 0 (nadie ha solicitado todavía),
        # así que no tiene sentido adelantarse — se pide siempre el mes actual.
        async def _seleccionar_diag(page, selector, valor, campo):
            """select_option con diagnostico: si falla, imprime las opciones reales
            disponibles antes de relanzar la excepcion (para saber, sin re-loguearse,
            si cambio el ID o el formato del value — ej. mes con cero a la izquierda).
            force=True: SSASUR empezo a envolver estos <select> con un widget JS que
            los oculta visualmente (reemplazados por una caja con flecha) — el
            elemento real sigue en el DOM con sus <option> intactas, pero Playwright
            rechaza select_option() por su chequeo de visibilidad. force=True se
            salta ese chequeo y opera igual sobre el <select> real (confirmado en
            vivo 17-07-2026: sin force, timeout 'element is not visible' pese a que
            la opcion buscada SI figura en la lista de opciones del diagnostico)."""
            try:
                await page.select_option(selector, valor, force=True)
            except Exception:
                try:
                    opciones = await page.eval_on_selector(
                        selector,
                        "el => Array.from(el.options).map(o => o.value + ':' + o.textContent.trim())",
                    )
                    print(f"  [diag] {campo} ({selector}): no se pudo fijar '{valor}'. "
                          f"Opciones reales: {opciones}")
                except Exception as e2:
                    print(f"  [diag] {campo} ({selector}): no se encontro el elemento ({e2})")
                raise

        if not no_programacion:
            print("\n[4b/9] Reporte de Programación AA (mes en curso)...")
            try:
                await page.goto(PROGRAMACION_REPORTE)
                await _networkidle(page)
                await page.wait_for_timeout(1_500)
                await _seleccionar_diag(page, "#ano", str(today.year), "Año")
                await _seleccionar_diag(page, "#mes", str(today.month), "Mes")
                await _seleccionar_diag(page, "#distribucion", DISTRIB_FARMACOS, "Distribucion")
                await _seleccionar_diag(page, "#select_proyecto", PROYECTO_LOCAL, "Proyecto")
                await _seleccionar_diag(page, "#cc", CC_FARMACIA, "Centro de Costo")
                await page.wait_for_timeout(500)
                _dest_prog = await descargar(
                    page, MAESTRO_DIR,
                    lambda: page.click("#generarXLS_salida", force=True),
                    "PROGRAMACION AA",
                )
                _limpiar_versiones_previas(
                    MAESTRO_DIR,
                    "cantidad_de_productos_consumidos_en_centro_de_costo_farmacia*.xlsx",
                    _dest_prog,
                )
            except Exception as e:
                print(f"  [ERROR] Descarga Programación AA falló: {e}")
                await page.screenshot(path=str(MAESTRO_DIR / "debug_programacion.png"))
                print("  Screenshot guardado: debug_programacion.png")

        # ── PASO 4c — CONSOLIDADO HEMOGRAMAS CLOZAPINA (OPT-IN, --clozapina) ─────
        # Busca en HCE el hemograma más reciente de cada paciente con Clozapina
        # despachada (Atención Abierta) y arma el Excel de ingreso a MINSAL.
        # Reutiliza esta MISMA sesión/pestaña de Playwright — por eso corre ACÁ,
        # ANTES de browser.close(), y no como subprocess.run() más abajo (bug
        # real detectado en vivo 23-07-2026: colocado después de browser.close()
        # producía "Page.goto: Target page, context or browser has been closed").
        if con_clozapina:
            cloz_py = MAESTRO_DIR / "clozapina_hce_hemogramas.py"
            if cloz_py.exists():
                print(f"\n[4c/9] Consolidado de hemogramas Clozapina...")
                try:
                    from clozapina_hce_hemogramas import (
                        entrar_hce, esperar_firma_electronica, buscar_paciente,
                        descargar_hemograma_mas_reciente, limpiar_busqueda, _run_a_partes,
                        HEMOGRAMAS_DIR, _construir_filas_desde_cache, _cache_key,
                    )
                    from clozapina_consolidar import (
                        cargar_despachos_clozapina, extraer_hemograma_pdf,
                        construir_fila, guardar_excel, FilaConsolidada, ruta_archivo_mes,
                    )
                    def _run_sin_guion_cloz(run):
                        return run.replace("-", "").replace(".", "")

                    async def _procesar_dia_clozapina(fecha_objetivo):
                        despachos_cloz = cargar_despachos_clozapina(str(MAESTRO_DIR), desde=fecha_objetivo, hasta=fecha_objetivo)
                        if despachos_cloz.empty:
                            print(f"  Sin despachos de Clozapina el {fecha_objetivo.strftime('%d-%m-%Y')}.")
                            return
                        # .last() por seguridad si el mismo paciente aparece dos veces ese
                        # día (ej. receta fraccionada en cuotas despachadas juntas).
                        por_paciente_cloz = despachos_cloz.sort_values("fecha_despacho").groupby("run_normalizado").last()
                        por_paciente_cloz = por_paciente_cloz.reset_index().sort_values("fecha_despacho")
                        print(f"  {len(por_paciente_cloz)} paciente(s) con Clozapina despachada el {fecha_objetivo.strftime('%d-%m-%Y')}.")

                        _hay_pendientes = any(
                            not (HEMOGRAMAS_DIR / f"{_cache_key(_run_sin_guion_cloz(r), f)}.pdf").exists()
                            for r, f in zip(por_paciente_cloz["run_normalizado"], por_paciente_cloz["fecha_despacho"])
                        )
                        if not _hay_pendientes:
                            # Mismo criterio que clozapina_hce_hemogramas.py standalone:
                            # si TODOS los pacientes ya tienen hemograma cacheado, no hace
                            # falta entrar a HCE ni esperar Firma Electrónica (que puede
                            # quedar colgada o fallar si no hay nadie frente al equipo).
                            print("  Sin despachos nuevos — todos cacheados. No se entra a HCE.")
                            filas_minsal_cloz, filas_obs_cloz, _fallidos_cloz = _construir_filas_desde_cache(por_paciente_cloz)
                            _ruta_cloz = ruta_archivo_mes(fecha_objetivo)
                            guardar_excel(filas_minsal_cloz, filas_obs_cloz, str(_ruta_cloz))
                            print(f"  ✓ {_ruta_cloz.name} ({len(por_paciente_cloz)} pacientes"
                                  f" — 0 nuevos buscados en HCE, {len(por_paciente_cloz)} ya cacheados)")
                        else:
                            await entrar_hce(page)
                            await esperar_firma_electronica(page)
                            filas_minsal_cloz, filas_obs_cloz = [], []
                            nuevos_cloz = 0
                            for _, prow in por_paciente_cloz.iterrows():
                                run = prow["run_normalizado"]
                                run_sin_guion = run.replace("-", "").replace(".", "")
                                run_num, dv = _run_a_partes(run)
                                nombre = f"{prow['Nombre']} {prow['Apellido Paterno']} {prow['Apellido Materno']}"
                                hemograma = None
                                cache_key_cloz = _cache_key(run_sin_guion, prow["fecha_despacho"])
                                pdf_cacheado = HEMOGRAMAS_DIR / f"{cache_key_cloz}.pdf"
                                if pdf_cacheado.exists():
                                    # Ya se descargó el hemograma de ESTE despacho puntual
                                    # (mismo paciente + misma fecha) en una corrida anterior.
                                    # Un despacho nuevo del mismo paciente en otra fecha
                                    # siempre vuelve a buscar en HCE (bug real 11-08-2026:
                                    # cache solo por RUN reutilizaba para siempre el PDF del
                                    # primer despacho, ignorando exámenes más nuevos).
                                    try:
                                        hemograma = extraer_hemograma_pdf(pdf_cacheado)
                                    except Exception:
                                        hemograma = None
                                if hemograma is None:
                                    nuevos_cloz += 1
                                    if await buscar_paciente(page, run_num, dv):
                                        pdf_path = await descargar_hemograma_mas_reciente(context, page, run_sin_guion, cache_key_cloz)
                                        if pdf_path:
                                            try:
                                                hemograma = extraer_hemograma_pdf(pdf_path)
                                            except Exception as e:
                                                print(f"    [aviso] no se pudo leer PDF de {run_sin_guion}: {e}")
                                    await limpiar_busqueda(page)
                                fc = FilaConsolidada(
                                    run=run, nombre=nombre,
                                    fecha_despacho=prow["fecha_despacho"].date() if prow["fecha_despacho"] == prow["fecha_despacho"] else None,
                                    dosis_mg_dia=prow["dosis_mg_dia"], cuota_receta=prow["cuota_receta"] or "",
                                    posologia_texto=prow["posologia_texto"] or "", hemograma=hemograma,
                                )
                                fm, fo = construir_fila(fc, today)
                                filas_minsal_cloz.append(fm)
                                filas_obs_cloz.append(fo)
                            _ruta_cloz = ruta_archivo_mes(fecha_objetivo)
                            guardar_excel(filas_minsal_cloz, filas_obs_cloz, str(_ruta_cloz))
                            print(f"  ✓ {_ruta_cloz.name} ({len(por_paciente_cloz)} pacientes"
                                  f" — {nuevos_cloz} nuevos buscados en HCE, {len(por_paciente_cloz) - nuevos_cloz} ya cacheados)")

                    # A pedido del usuario (11-08-2026): cada corrida solo busca en HCE
                    # los despachos del día hábil ANTERIOR al de hoy — lunes mira el
                    # viernes (fin de semana no hay despacho), el resto de los días
                    # mira el día calendario anterior. Reemplaza el barrido de "todo
                    # el mes" que se usaba antes (más lento, revisaba pacientes que no
                    # tenían nada nuevo). OJO: si un día se salta la corrida de
                    # CLOZAPINA.bat, los despachos de ESE día quedan sin buscar en HCE
                    # (no hay reintento automático más adelante, a diferencia del
                    # barrido mensual anterior) — correr manualmente con
                    # --desde/--hasta si se detecta un día saltado.
                    fecha_objetivo_cloz = today - timedelta(days=3 if today.weekday() == 0 else 1)
                    await _procesar_dia_clozapina(fecha_objetivo_cloz)

                    # Sube el Consolidado a su carpeta Drive privada en ESTA misma
                    # corrida — antes quedaba pendiente de --solo-clozapina manual,
                    # que nadie disparaba, y el reporte nunca llegaba a Drive
                    # (detectado 30-07-2026). Mismo patrón que usa Centinela arriba.
                    pub_py_cloz = MAESTRO_DIR / "publicar_drive.py"
                    if pub_py_cloz.exists() and (MAESTRO_DIR / "token_drive.json").exists():
                        print("  [Clozapina] Sincronizando reporte a Drive...")
                        pret_cloz = subprocess.run(
                            [sys.executable, str(pub_py_cloz), "--solo-clozapina"],
                            cwd=str(MAESTRO_DIR),
                            env={**os.environ, "PYTHONUTF8": "1", "PYTHONIOENCODING": "utf-8"},
                        )
                        if pret_cloz.returncode != 0:
                            print(f"  [aviso] publicar_drive.py --solo-clozapina terminó con código {pret_cloz.returncode}")
                except Exception as e:
                    import traceback
                    print(f"  [aviso] Consolidado Clozapina falló: {e}")
                    traceback.print_exc()

        # ── PASO 4e — SERVICIOS FARMACÉUTICOS (mensual, Agenda Médica) ──────────
        # Recuento de actividades QF del MES ANTERIOR. Igual que Centinela
        # (semanal): dispara en los primeros días hábiles del mes con un
        # marcador para no duplicar, y hace catch-up si el 1er día hábil no
        # corrió (PC apagado, sesión no lista, etc.). --servicios-farmaceuticos
        # fuerza la corrida ignorando el gatillo (para la 1ª prueba supervisada
        # o para reprocesar un mes puntual con --mes-servicios MM-AAAA).
        if no_servicios:
            print("\n[4e/9] Servicios Farmacéuticos (Agenda Médica) — omitido (--no-servicios-farmaceuticos).")
        else:
            import json as _json
            if _mes_servicios_arg:
                _mm, _aaaa = _mes_servicios_arg.split("-")
                anio_srv, mes_srv = int(_aaaa), int(_mm)
            else:
                anio_srv, mes_srv = _mes_anterior(today)

            marcador = SERVICIOS_MARCADOR
            ya_generado = False
            if marcador.exists():
                try:
                    prev = _json.loads(marcador.read_text(encoding="utf-8"))
                    ya_generado = (prev.get("anio") == anio_srv and prev.get("mes") == mes_srv)
                except Exception:
                    ya_generado = False

            primer_habil = primer_dia_habil(today.year, today.month)
            dentro_ventana = today <= primer_habil + timedelta(days=4)   # catch-up: primeros ~5 días hábiles/corridos del mes

            if forzar_servicios:
                disparar = True
                motivo = "forzado (--servicios-farmaceuticos)"
            elif ya_generado:
                disparar = False
                motivo = f"ya generado para {mes_srv:02d}-{anio_srv} — omitido"
            elif not dentro_ventana:
                disparar = False
                motivo = f"fuera de ventana mensual (1er día hábil fue {fmt(primer_habil)}) — se generará el próximo mes"
            else:
                disparar = True
                motivo = "catch-up" if today != primer_habil else "1er día hábil del mes (ciclo normal)"

            if not disparar:
                print(f"\n[4e/9] Servicios Farmacéuticos — {motivo}.")
            else:
                print(f"\n[4e/9] Servicios Farmacéuticos — {motivo}...")
                desde_srv = fmt(date(anio_srv, mes_srv, 1))
                hasta_srv = fmt(_ultimo_dia_mes(anio_srv, mes_srv))
                try:
                    dest_srv, n_srv = await paso_agenda_actividades(page, desde_srv, hasta_srv, debug_agenda)
                except Exception as e:
                    print(f"  [ERROR] Agenda Médica falló: {e}")
                    dest_srv, n_srv = None, -1

                if dest_srv:
                    print(f"  ✓ {dest_srv.name} ({n_srv} filas)")
                    env_utf8_srv = {**os.environ, "PYTHONUTF8": "1", "PYTHONIOENCODING": "utf-8"}
                    resumen_py = MAESTRO_DIR / "servicios_farmaceuticos.py"
                    sret = subprocess.run(
                        [sys.executable, str(resumen_py), str(dest_srv), "--mes", f"{mes_srv:02d}-{anio_srv}"],
                        cwd=str(MAESTRO_DIR), env=env_utf8_srv,
                    )
                    if sret.returncode == 0:
                        try:
                            marcador.parent.mkdir(parents=True, exist_ok=True)
                            marcador.write_text(_json.dumps(
                                {"anio": anio_srv, "mes": mes_srv, "fecha": today.isoformat()},
                                ensure_ascii=False, indent=2), encoding="utf-8")
                        except Exception as e:
                            print(f"  [aviso] no se pudo escribir el marcador mensual: {e}")
                        pub_py_srv = MAESTRO_DIR / "publicar_drive.py"
                        if pub_py_srv.exists() and (MAESTRO_DIR / "token_drive.json").exists():
                            print("  [Servicios Farmacéuticos] Sincronizando reporte a Drive...")
                            pret_srv = subprocess.run(
                                [sys.executable, str(pub_py_srv), "--solo-servicios"],
                                cwd=str(MAESTRO_DIR), env=env_utf8_srv,
                            )
                            if pret_srv.returncode != 0:
                                print(f"  [aviso] publicar_drive.py --solo-servicios terminó con código {pret_srv.returncode}")
                    else:
                        print(f"  [aviso] servicios_farmaceuticos.py terminó con código {sret.returncode} — no se marca el mes como hecho")
                elif n_srv == 0:
                    print(f"  · Sin actividades en Agenda Médica para {desde_srv} → {hasta_srv}.")
                else:
                    print("  [AVISO] No se generó el Excel de Agenda Médica — revisa debug_agenda.png.")
                    print("  Si es la 1ª corrida: vuelve a intentar con --servicios-farmaceuticos --debug-agenda")
                    print("  y revisa los volcados [DESCUBRIR] para ajustar los selectores en AUTO_SSASUR.py.")

        try:
            await browser.close()
        except Exception as e:
            # Si el browser ya murió a mitad de corrida (TargetClosedError, ver
            # PASO 3), volver a cerrarlo puede fallar también — no debe impedir
            # llegar a PASO 5 (maestro_aa.py), que es lo que realmente genera
            # el reporte diario.
            print(f"  [aviso] browser.close() falló (probablemente ya estaba cerrado): {e}")

    # ── PASO 5 — MAESTRO AA ────────────────────────────────────────────────────
    print("\n[5/9] Actualizando Maestro AA...")
    env_utf8 = {**os.environ, "PYTHONUTF8": "1", "PYTHONIOENCODING": "utf-8"}
    result = subprocess.run(
        [sys.executable, str(MAESTRO_DIR / "maestro_aa.py")],
        cwd=str(MAESTRO_DIR),
        env=env_utf8,
    )

    print()
    if result.returncode == 0:
        print("═" * 62)
        print("  ✓ COMPLETADO")
        print("  → Consolidado_AA_MAESTRO.xlsx  actualizado")
        print("  → Resumen_Pedidos_AA.xlsx       actualizado")
        print("═" * 62)

        # Marca HOY como la última corrida exitosa — la próxima corrida usa
        # esto para detectar huecos reales (no solo fin de semana/feriado) en
        # el rango GT a revisar. Ver ultima_corrida_ok()/rango_backlog_gt().
        marcar_corrida_ok(today)

        # ── PASO 5b — AUDITORÍA DE PRESCRIPCIÓN ──────────────────────────────
        audit_py = MAESTRO_DIR / "auditoria_prescripcion.py"
        if audit_py.exists():
            print("\n[5b/9] Generando auditoría de prescripción...")
            aret = subprocess.run([sys.executable, str(audit_py)], cwd=str(MAESTRO_DIR), env=env_utf8)
            if aret.returncode != 0:
                print(f"  [aviso] auditoria_prescripcion.py terminó con código {aret.returncode}")

        # ── PASO 5c — CRUCE GT + PLANILLAS ───────────────────────────────────
        if gt_dest:
            out_gt = gt_salida(gt_dest)   # carpeta propia por rango de fechas
            print(f"\n[5c/9] Cruce GT + generando planillas → out_gt/{out_gt.name}/ ...")
            cruce = MAESTRO_DIR / "cruce_gt.py"
            if cruce.exists():
                cret = subprocess.run(
                    [sys.executable, str(cruce), str(gt_dest),
                     "--salida", str(out_gt), "--generar"],
                    cwd=str(MAESTRO_DIR), env=env_utf8,
                )
                if cret.returncode != 0:
                    # Antes esto no se chequeaba (a diferencia de 5d/5e/5f, que sí
                    # avisan) — si cruce_gt.py/generar.py fallaba a medio camino,
                    # el día se quedaba sin nóminas y no quedaba ningún rastro en
                    # el log para notarlo (mismo patrón silencioso que el bug de
                    # backlog GT arreglado arriba).
                    print(f"  [aviso] cruce_gt.py terminó con código {cret.returncode} — revisa si "
                          f"out_gt/{out_gt.name}/ quedó con todas las planillas esperadas.")
            else:
                print(f"  [aviso] cruce_gt.py no encontrado — omitiendo cruce.")

        # ── PASO 5c2 — FUSIONAR NÓMINAS GT (auto + manual/backlog) ───────────
        # verificar_backlog_gt() (PASO 3) puede haber generado hoy una o más
        # "Nomina_Manual_*.xlsx" para el mismo destino que cruce_gt.py acaba
        # de procesar arriba. Sin este paso, ambos archivos quedaban sueltos
        # y se publicaban tal cual — la QF seguía recibiendo nóminas
        # rotuladas "manual" (esa distinción ya no aplica en Gestión
        # Territorial, ver fusionar_nominas_gt.py) porque nadie corría
        # fusionar_nominas_gt.py a mano. Corre SIEMPRE (idempotente si no
        # hay nada que fusionar) para que lo que se publique en el PASO 7-9
        # sea siempre el nombre oficial "<destino>_Planilla.xlsx".
        fusion_py = MAESTRO_DIR / "fusionar_nominas_gt.py"
        if fusion_py.exists():
            print(f"\n[5c2/9] Fusionando nóminas GT (automática + manual/backlog)...")
            fret = subprocess.run(
                [sys.executable, str(fusion_py), "--todos"],
                cwd=str(MAESTRO_DIR), env=env_utf8,
            )
            if fret.returncode != 0:
                print(f"  [aviso] fusionar_nominas_gt.py terminó con código {fret.returncode}")

        # ── PASO 5d — REGISTRO ISP RECETAS CHEQUE ────────────────────────────
        # Consume la MISMA sábana ya descargada: filtra recetas cheque AT Abierta
        # y agrega los folios nuevos al formulario ISP del mes vigente. El
        # formulario vive fuera del repo (carpeta de la QF) → no se publica.
        rch_py = MAESTRO_DIR / "recetas_cheque.py"
        if not no_rch and rch_py.exists():
            print(f"\n[5d/9] Registro ISP de recetas cheque (Farmacia AT Abierta)...")
            dret = subprocess.run(
                [sys.executable, str(rch_py), "--no-pause"],
                cwd=str(MAESTRO_DIR), env=env_utf8,
            )
            if dret.returncode != 0:
                print(f"  [aviso] recetas_cheque.py terminó con código {dret.returncode}")

        # ── PASO 5e — PEDIDO FUSIONADO ───────────────────────────────────────
        pedido_py = MAESTRO_DIR / "pedido_fusion.py"
        if pedido_py.exists():
            print(f"\n[5e/9] Generando Pedido Fusionado (Farm_Bod + Bod_Farmacos + Dialisis)...")
            pret = subprocess.run(
                [sys.executable, str(pedido_py)],
                cwd=str(MAESTRO_DIR), env=env_utf8,
            )
            if pret.returncode != 0:
                print(f"  [aviso] pedido_fusion.py terminó con código {pret.returncode}")

        # ── PASO 5f — PEDIDO FUSIONADO SIMPLE (Farm_Bodega + Faltantes_AA) ────
        pedido_simple_py = MAESTRO_DIR / "pedido_fusion_simple.py"
        if pedido_simple_py.exists():
            print(f"\n[5f/9] Generando Pedido Fusionado Simple (Farm_Bodega + Faltantes_AA)...")
            psret = subprocess.run(
                [sys.executable, str(pedido_simple_py)],
                cwd=str(MAESTRO_DIR), env=env_utf8,
            )
            if psret.returncode != 0:
                print(f"  [aviso] pedido_fusion_simple.py terminó con código {psret.returncode}")

        # NOTA: sgli_historico.py NO se corre aquí — maestro_aa.py YA lo llama
        # internamente al final de su propio main() (import sgli_historico;
        # sgli_historico.main()). Correrlo de nuevo en este paso lo duplicaba
        # (2 Excel SGLI_Historico_*.xlsx por corrida, ~15s desperdiciados) —
        # visto en vivo el 2026-07-13: SGLI_Historico_..._1801 y _1802.

        # ── PASO 6 — DEDUPLICAR RECETAS GT ────────────────────────────────────
        # Detecta y limpia recetas duplicadas entre archivos GT descargados con
        # rangos solapados (sobre-extracción). Actúa en modo --limpiar: crea .bak
        # antes de modificar cualquier archivo. Corre ANTES de publicar para que
        # GitHub/Escritorio/Drive reciban la versión ya limpia (antes el dedup
        # corría a mitad del publicado y GitHub/Escritorio se quedaban con la
        # versión sin deduplicar hasta la corrida siguiente).
        dedup_py = MAESTRO_DIR / "dedup_recetas.py"
        if dedup_py.exists():
            print("\n[6/9] Deduplicando recetas GT por sobre-extracción...")
            ddup = subprocess.run(
                [sys.executable, str(dedup_py), "--limpiar"],
                cwd=str(MAESTRO_DIR), env=env_utf8,
            )
            if ddup.returncode != 0:
                print(f"  [aviso] dedup_recetas.py terminó con código {ddup.returncode}")

        # ── PASO 7-9 — SINCRONIZAR TODO (Escritorio + GitHub + Drive + RCh) ────
        # SINCRONIZAR_TODO.bat es el mismo script que corre el acceso directo
        # "Sincronizar Todo" del Escritorio — una sola fuente de verdad para
        # "cómo se publica todo", en vez de duplicar las 4 llamadas aquí y allá.
        # Incluye Recetas Cheque ISP → Drive (excepción autorizada por el usuario
        # 2026-06-30, confirmada AUTOMÁTICA en cada corrida 2026-07-15 — sube RUT
        # de pacientes sin confirmación puntual; --no-rch la desactiva).
        # OJO: --no-publicar solo agrega "--no-git" acá abajo — Escritorio y Drive
        # se sincronizan SIEMPRE que exista SINCRONIZAR_TODO.bat, con datos
        # reales. No hay flag para saltarse este paso completo.
        sync_bat = MAESTRO_DIR / "SINCRONIZAR_TODO.bat"
        if sync_bat.exists():
            print("\n[7-9/9] Sincronizando todo (Escritorio + GitHub + Drive + Recetas Cheque)...")
            args = ["cmd", "/c", str(sync_bat)]
            if no_publicar:
                args.append("--no-git")
            if no_rch:
                args.append("--no-rch")
            args.append("--no-pause")
            sret = subprocess.run(args, cwd=str(MAESTRO_DIR), env=env_utf8)
            if sret.returncode != 0:
                print(f"  [aviso] SINCRONIZAR_TODO.bat terminó con código {sret.returncode}")
        else:
            print("\n[7-9/9] SINCRONIZAR_TODO.bat no encontrado — omitiendo publicación.")
    else:
        print("  [ERROR] maestro_aa.py falló — revisa los mensajes arriba")

    # ── CENTINELA — REPORTE SEMANAL con catch-up (una vez por semana) ─────────
    # Genera el PDF semanal para el MINSAL con los CSV+XLSX recién descargados
    # y lo sincroniza a Drive en la MISMA corrida.
    #
    # Antes: sólo corría si hoy==lunes, acoplado a que AUTO_SSASUR llegara
    # completo hasta aquí ESE lunes. Si el lunes la tarea no disparaba (PC
    # apagado), el login SSASUR no se completaba, o un paso previo caía
    # (TargetClosedError / token Drive invalid_grant), el reporte se perdía en
    # silencio toda la semana: martes-viernes lo omitían por "no ser lunes".
    # Peor aún, el PDF se generaba DESPUÉS del SINCRONIZAR_TODO, así que ni
    # siquiera se subía a Drive ese mismo día. (Incidente lunes 20-07-2026.)
    #
    # Ahora: se ejecuta en la PRIMERA corrida hábil (lun-vie) de cada semana
    # ISO en que aún no se haya generado — normalmente el lunes, y si el lunes
    # se perdió, el primer día siguiente que AUTO_SSASUR corra hace el
    # catch-up. Un marcador por semana ISO evita duplicados, y tras generarlo
    # se publica a Drive con --solo-centinela sin esperar a la corrida del día
    # siguiente. Sigue siendo UN reporte por semana epidemiológica.
    centinela_py = MAESTRO_DIR / "centinela_reporte.py"
    if centinela_py.exists():
        import json
        anio_iso, sem_iso, _ = today.isocalendar()
        marcador = MAESTRO_DIR / "Centinela_Reportes" / "_ultima_semana.json"
        ya_generado = False
        if marcador.exists():
            try:
                prev = json.loads(marcador.read_text(encoding="utf-8"))
                ya_generado = (prev.get("anio_iso") == anio_iso
                               and prev.get("sem_iso") == sem_iso)
            except Exception:
                ya_generado = False

        if today.weekday() >= 5:
            print(f"\n[CENTINELA] Fin de semana ({today.strftime('%A')}) — se generará el lunes.")
        elif ya_generado:
            print(f"\n[CENTINELA] Ya generado esta semana (ISO {anio_iso}-S{sem_iso}) — omitido.")
        else:
            es_catchup = today.weekday() != 0
            motivo = "catch-up: el lunes no se generó" if es_catchup else "lunes (ciclo normal)"
            print(f"\n[CENTINELA] Reporte Centinela — {motivo}...")
            cret = subprocess.run(
                [sys.executable, str(centinela_py), "--no-pause"],
                cwd=str(MAESTRO_DIR), env=env_utf8,
            )
            if cret.returncode == 0:
                # Marca la semana ISO como hecha para no duplicar en corridas siguientes.
                try:
                    marcador.parent.mkdir(parents=True, exist_ok=True)
                    marcador.write_text(json.dumps(
                        {"anio_iso": anio_iso, "sem_iso": sem_iso, "fecha": today.isoformat()},
                        ensure_ascii=False, indent=2), encoding="utf-8")
                except Exception as e:
                    print(f"  [aviso] no se pudo escribir el marcador semanal: {e}")
                # Sube el PDF a Drive en ESTA corrida (antes nacía tras el sync
                # y quedaba para el día siguiente).
                pub_py = MAESTRO_DIR / "publicar_drive.py"
                if pub_py.exists() and (MAESTRO_DIR / "token_drive.json").exists():
                    print("  [CENTINELA] Sincronizando reporte a Drive...")
                    pret = subprocess.run(
                        [sys.executable, str(pub_py), "--solo-centinela"],
                        cwd=str(MAESTRO_DIR), env=env_utf8,
                    )
                    if pret.returncode != 0:
                        print(f"  [aviso] publicar_drive.py --solo-centinela terminó con código {pret.returncode}")
            else:
                print(f"  [aviso] centinela_reporte.py terminó con código {cret.returncode} — revisa el reporte centinela")

    if not no_pause:
        try:
            input("\nPresiona Enter para cerrar...")
        except EOFError:
            pass


if __name__ == "__main__":
    # Impide que el equipo (notebook) suspenda por inactividad durante la
    # corrida desatendida: la tarea programada no cuenta como "actividad" para
    # el detector de idle de Windows, y el timeout en batería es de solo 25 min
    # — bastante menos que lo que tarda la corrida completa. Si el equipo se
    # suspende a mitad de camino, Chromium queda con la conexión cortada y
    # cualquier operación posterior de Playwright falla con TargetClosedError
    # (causa raíz identificada 31-08-2026: 3 días sin reporte diario porque la
    # corrida de las 9:00 moría así). ES_CONTINUOUS|ES_SYSTEM_REQUIRED|
    # ES_DISPLAY_REQUIRED se libera siempre al terminar, sea éxito o error.
    ES_CONTINUOUS       = 0x80000000
    ES_SYSTEM_REQUIRED  = 0x00000001
    ES_DISPLAY_REQUIRED = 0x00000002
    try:
        ctypes.windll.kernel32.SetThreadExecutionState(
            ES_CONTINUOUS | ES_SYSTEM_REQUIRED | ES_DISPLAY_REQUIRED
        )
    except Exception:
        pass   # no-Windows / entorno sin kernel32 (no debería pasar aquí, pero no es motivo de aborto)
    try:
        asyncio.run(main())
    finally:
        try:
            ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS)
        except Exception:
            pass
