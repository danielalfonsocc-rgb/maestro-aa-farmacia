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
  · --no-publicar      no publica en GitHub (debug)

El registro ISP de recetas cheque (recetas_cheque.py) corre como PASO 5d: usa
la MISMA sábana ya descargada para agregar los folios cheque nuevos de Farmacia
AT Abierta al formulario ISP del mes vigente (carpeta de la QF, fuera del repo).

Tras el Consolidado corre también (sin argumentos, sin llamadas a IA):
  5f) reposicion_dias_habiles.py → Reposicion_DiasHabiles_AA_<fecha>.xlsx
Ese archivo ya era esperado por publicar_escritorio.py / publicar_drive.py; sin
este paso quedaba publicando una copia vieja del plan.
(sgli_historico.py → SGLI_Historico_<fecha>.xlsx NO se agrega aquí: maestro_aa.py
YA lo corre solo, al final de su propio main() — agregarlo lo duplicaba.)

NOTA: agente_duplicados.py y auditoria_duplicados_profunda.py (llaman a la API
de Claude) NO corren aquí — se mantienen en sus .bat propios para ejecutarse
a demanda y no gastar tokens en cada corrida de AUTO_SSASUR.
═══════════════════════════════════════════════════════════════
"""
import asyncio
import os
import subprocess
import sys
from pathlib import Path
from datetime import date, timedelta

# Forzar UTF-8 en la salida — si no, al redirigir a un archivo/pipe Windows usa
# cp1252 y los caracteres ═ ✓ → ✗ revientan el script antes de empezar.
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass


# ── Configuración ─────────────────────────────────────────────────────────────
MAESTRO_DIR      = Path(__file__).parent
SESSION_FILE     = MAESTRO_DIR / ".ssasur_session.json"
DASHBOARD_URL    = "https://login.ssasur.cl/dashboard"
RECETA_INFORME   = "https://www.ssasur.cl/receta/informes/sabana"
STOCK_REPORTE    = "https://www.ssasur.cl/abastecimiento/reportes/stock_en_momento_bodega"
BODEGA_TODAS     = "0"    # bodega → "TODAS" (todas las bodegas en un archivo)
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


def fmt(d: date) -> str:
    return d.strftime("%d/%m/%Y")


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
    await page.wait_for_load_state("networkidle")
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


async def entrar_receta(page):
    """Dashboard → tarjeta RECETA → (modal proyecto) → www.ssasur.cl/receta.
    Si ya estamos dentro del módulo receta, vuelve al inicio del módulo sin
    pasar por el dashboard (evita el timeout del botón RECETA)."""
    if "www.ssasur.cl/receta" in page.url:
        await page.goto("https://www.ssasur.cl/receta", wait_until="networkidle")
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
    await page.wait_for_load_state("networkidle")
    await page.wait_for_timeout(1_500)
    print(f"  ✓ En módulo receta: {page.url}")


async def descargar_sabana(page, fi: str, ft: str) -> str:
    """
    Descarga el informe completo de recetas por POST directo (con la sesión ya
    acuñada al proyecto). Replica los dos POST exactos del flujo manual:
      1) /receta/informes/jsonValidaSabana   (valida; fija contexto)
      2) .../informeCompletoReceta.php        (devuelve el CSV ';' latin-1)
    fi/ft en formato dd/mm/yyyy. Devuelve el CSV como texto latin-1.
    """
    return await page.evaluate(
        r"""async ({fi, ft}) => {
          const enc = encodeURIComponent;
          const csrf = (document.querySelector('input[name=csrf_token]')||{}).value || '';
          try {
            await fetch('https://www.ssasur.cl/receta/informes/jsonValidaSabana', {
              method:'POST',
              headers:{'content-type':'application/x-www-form-urlencoded; charset=UTF-8','x-requested-with':'XMLHttpRequest'},
              body:`csrf_token=${csrf}&fechaInicio=${enc(fi)}&fechaTermino=${enc(ft)}&estadoSolicitadoHidden=F`,
              credentials:'include',
            });
          } catch(e){}
          const r = await fetch('https://www.ssasur.cl/application/sistemas/receta/reportes/informeCompletoReceta.php', {
            method:'POST', headers:{'content-type':'application/x-www-form-urlencoded'},
            body:`fechaInicio=${enc(fi)}&fechaTermino=${enc(ft)}&estadoSolicitadoHidden=F&entregaDomicilio=F&estadoAnulada=N`,
            credentials:'include',
          });
          const buf = await r.arrayBuffer();
          const b = new Uint8Array(buf); let s='';
          for (let i=0;i<b.length;i++) s += String.fromCharCode(b[i]);
          return s;
        }""",
        {"fi": fi, "ft": ft},
    )


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
        await page.wait_for_load_state("networkidle")
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
    await page.wait_for_load_state("networkidle")
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


async def _set_fechas(page, desde, hasta):
    """Rellena fechaInicio/fechaTermino simulando tipeo real (click + seleccionar
    todo + escribir carácter por carácter + Tab). Necesario porque asignar
    .value vía JS y disparar eventos sintéticos (el método anterior) dejaba
    fechaInicio VACÍO de forma intermitente — visto en vivo 14-07-2026: el
    campo mostraba value='' pese al assignment, mientras fechaTermino sí
    tomaba el valor. El campo tiene máscara/validación propia que solo
    reacciona a eventos de teclado reales, no a Event() sintéticos en bloque.
    Cae a la heurística JS anterior si el campo por id no existe. Fechas en
    dd/mm/yyyy."""
    async def _escribir(field_id, valor):
        try:
            loc = page.locator(f"#{field_id}")
            if await loc.count() == 0:
                return False
            await loc.click()
            await loc.press("Control+A")
            await loc.press("Delete")
            await loc.type(valor, delay=30)
            await page.keyboard.press("Tab")
            await page.wait_for_timeout(250)
            real = await loc.input_value()
            return real == valor
        except Exception:
            return False

    ok_ini = await _escribir(SEL_FECHA_INI, desde) if desde else True
    ok_fin = await _escribir(SEL_FECHA_FIN, hasta) if hasta else True
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
    await page.wait_for_load_state("networkidle")
    await page.wait_for_timeout(2_500)
    if debug:
        await _dump_formulario(page, "post-buscar")
        await page.screenshot(path=str(MAESTRO_DIR / "debug_gt.png"))

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
    # No publicar en GitHub al final (útil para corridas de prueba/debug).
    no_publicar  = "--no-publicar" in sys.argv
    # Modo GT exclusivo: solo gestión territorial y termina (sin sábana/stock/maestro).
    gt_mode  = ("--gt" in sys.argv) or ("--solo-gt" in sys.argv)
    # Solo stock: salta recetas y GT, baja solo stock y corre maestro.
    solo_stock = "--solo-stock" in sys.argv
    # En corrida completa, saltar GT con --no-gt (si ya se corrió hoy por separado).
    no_gt    = "--no-gt" in sys.argv or solo_stock
    # Saltar el registro ISP de recetas cheque con --no-rch.
    no_rch   = "--no-rch" in sys.argv
    debug_gt = "--debug-gt" in sys.argv        # volcados [DESCUBRIR …] + screenshots
    _fecha   = _arg_val("--fecha")             # atajo: mismo día en desde/hasta
    today = date.today()
    ayer  = today - timedelta(days=1)
    # Default GT: ayer → +13 días (2 semanas hacia adelante). Captura despachos
    # próximos programados (el reporte muestra pendientes, no histórico entregado).
    # Los despachos pasados desaparecen del listado al ser procesados, por eso se
    # mira al FUTURO en lugar de hacia atrás.
    # Se puede acotar con --desde dd/mm/yyyy o --fecha dd/mm/yyyy.
    _gt_inicio = _fecha or fmt(today - timedelta(days=1))
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
            try:
                await pg.wait_for_selector(
                    'button:has-text("ABASTECIMIENTO"), div:has-text("ABASTECIMIENTO")',
                    timeout=TIMEOUT_LOGIN,
                )
            except Exception:
                await pg.wait_for_function(
                    "document.body.innerText.includes('ABASTECIMIENTO')",
                    timeout=TIMEOUT_LOGIN,
                )

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
            await page.wait_for_load_state("networkidle")
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
            gt_dest, n_gt = await paso_gt(page, desde_gt, hasta_gt, debug_gt)
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

        # ════════════════════════════════════════════════════════════════════
        #  PASO 4 — ABASTECIMIENTO  (volver al dashboard → entrar → stock)
        # ════════════════════════════════════════════════════════════════════
        print("\n[4/9] Módulo ABASTECIMIENTO...")
        await page.goto(DASHBOARD_URL)
        await page.wait_for_load_state("networkidle")
        await page.wait_for_timeout(1_500)
        await entrar_modulo(page, "ABASTECIMIENTO")

        await page.goto(STOCK_REPORTE)
        await page.wait_for_load_state("networkidle")
        await page.wait_for_timeout(2_500)

        if "login.ssasur.cl" in page.url:
            # Por si la sesión del módulo no quedó lista: reintentar una vez
            print("  [AVISO] Rebote al login — reintentando entrar a ABASTECIMIENTO...")
            await page.goto(DASHBOARD_URL)
            await page.wait_for_load_state("networkidle")
            await page.wait_for_timeout(1_500)
            await entrar_modulo(page, "ABASTECIMIENTO")
            await page.goto(STOCK_REPORTE)
            await page.wait_for_load_state("networkidle")
            await page.wait_for_timeout(2_500)

        # Configurar el reporte: bodega = TODAS.
        # IMPORTANTE: NO seleccionar el establecimiento — hacerlo dispara un modal
        # ("Selección Proyecto") cuyo backdrop bloquea el botón Generar XLS.
        # El establecimiento ya viene en "PITRUFQUEN HOSP." (única opción).
        try:
            await page.select_option("#bodega", BODEGA_TODAS)
            await page.wait_for_timeout(1_000)
            valor_actual = await page.eval_on_selector("#bodega", "el => el.value")
            if valor_actual != BODEGA_TODAS:
                raise RuntimeError(f"selector #bodega quedó en '{valor_actual}', no en TODAS ('{BODEGA_TODAS}')")
            print("  Bodega: TODAS")
        except Exception as e:
            print(f"  [ERROR] No se pudo seleccionar bodega TODAS — abortando para no generar stock parcial: {e}")
            await page.screenshot(path=str(MAESTRO_DIR / "debug_stock.png"))
            raise

        print("  Generando XLS de stock... (puede tardar varios minutos)")
        try:
            # force=True: el overlay "Generando reporte. Espere un momento."
            # puede interponerse al clic; force lo atraviesa.
            await descargar(
                page, MAESTRO_DIR,
                lambda: page.click("#generarXLS_stock", force=True),
                "ABASTECIMIENTO",
            )
        except Exception as e:
            print(f"  [ERROR] Descarga ABASTECIMIENTO falló: {e}")
            await page.screenshot(path=str(MAESTRO_DIR / "debug_stock.png"))
            print("  Screenshot guardado: debug_stock.png")

        await browser.close()

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
                subprocess.run(
                    [sys.executable, str(cruce), str(gt_dest),
                     "--salida", str(out_gt), "--generar"],
                    cwd=str(MAESTRO_DIR), env=env_utf8,
                )
            else:
                print(f"  [aviso] cruce_gt.py no encontrado — omitiendo cruce.")

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

        # ── PASO 5f — REPOSICIÓN DÍAS HÁBILES ────────────────────────────────
        # Plan Bodega→Farmacia ajustado por feriados (reposicion_dias_habiles.py).
        # publicar_escritorio.py y publicar_drive.py ya esperan este archivo —
        # sin este paso quedaban copiando una versión vieja del plan.
        repo_py = MAESTRO_DIR / "reposicion_dias_habiles.py"
        if repo_py.exists():
            print(f"\n[5f/9] Generando Reposición Días Hábiles (Bodega→Farmacia)...")
            rret = subprocess.run(
                [sys.executable, str(repo_py)],
                cwd=str(MAESTRO_DIR), env=env_utf8,
            )
            if rret.returncode != 0:
                print(f"  [aviso] reposicion_dias_habiles.py terminó con código {rret.returncode}")

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

    # ── CENTINELA — REPORTE SEMANAL (solo lunes, independiente del maestro) ────
    # Auto-detecta los CSV + XLSX ya descargados en este mismo flujo y
    # genera el PDF semanal para el MINSAL. Se ejecuta UNA VEZ a la semana
    # (los lunes) para coincidir con el ciclo epidemiológico MINSAL.
    centinela_py = MAESTRO_DIR / "centinela_reporte.py"
    if centinela_py.exists():
        if today.weekday() == 0:   # 0 = lunes
            print(f"\n[CENTINELA] Reporte Centinela — Campaña Invierno 2026...")
            cret = subprocess.run(
                [sys.executable, str(centinela_py), "--no-pause"],
                cwd=str(MAESTRO_DIR), env=env_utf8,
            )
            if cret.returncode != 0:
                print(f"  [aviso] centinela_reporte.py terminó con código {cret.returncode} — revisa el reporte centinela")
        else:
            print(f"\n[CENTINELA] Solo se ejecuta los lunes — omitido hoy ({today.strftime('%A')}).")

    if not no_pause:
        try:
            input("\nPresiona Enter para cerrar...")
        except EOFError:
            pass


if __name__ == "__main__":
    asyncio.run(main())
