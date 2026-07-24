#!/usr/bin/env python3
"""stock_controlados.py — Stock del DÍA ANTERIOR para el conteo de controlados.

Baja el "Reporte Stock en Fecha" de SSASUR
(https://www.ssasur.cl/abastecimiento/reportes/stockFecha) del día anterior para
las bodegas de controlados (FARMACIA AT ABIERTA y FARMACIA AT CERRADA), lo cruza
con el mapeo de controlados_config.json y escribe:

    Conteo_Controlados/stock_sistema.json

que la app conteo_controlados_app.py inyecta al formulario para PRE-LLENAR la
columna "Sistema" con el stock del día anterior. La columna Acta/Vencimiento NO
se toca nunca (la maneja la QF a mano).

Estructura del JSON:
    {
      "fecha": "2026-07-23",                 # día consultado (día anterior)
      "generadoEn": "2026-07-24T08:15:00",
      "porFarmacia": {
        "abierta": { "Alprazolam|0,5 mg cm": 120, ... },
        "cerrada": { "Morfina|10 mg/ml am": 30, ... }
      }
    }

Uso:
  · Normalmente NO se corre suelto: AUTO_SSASUR.py lo llama dentro del módulo
    ABASTECIMIENTO (reutiliza la sesión ya logueada). Ver PASO 4d en AUTO_SSASUR.
  · Suelto para probar/backfill:  py stock_controlados.py            (día anterior)
                                  py stock_controlados.py 23/07/2026 (fecha fija)

NOTA — primera corrida en vivo: los IDs exactos del formulario stockFecha
(select de bodega, campo de fecha, botón Generar XLS) NO están confirmados por
inspección todavía. Este script los busca de forma DEFENSIVA (bodega por texto
de la opción, fecha por el input con formato dd/mm/aaaa, botón por su texto) y,
si algo falla, vuelca [DESCUBRIR …] y guarda debug_stock_controlados.png para
ajustar los selectores — mismo patrón que paso_gt / paso_controlados en
AUTO_SSASUR.py.
"""
import asyncio
import json
import sys
from datetime import date, datetime, timedelta
from pathlib import Path

for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

MAESTRO_DIR = Path(__file__).parent
CONFIG_FILE = MAESTRO_DIR / "controlados_config.json"
SALIDA_DIR  = MAESTRO_DIR / "Conteo_Controlados"
SALIDA_JSON = SALIDA_DIR / "stock_sistema.json"
STOCKFECHA_URL = "https://www.ssasur.cl/abastecimiento/reportes/stockFecha"
TIMEOUT_DESCARGA = 600_000


def _fmt(d: date) -> str:
    return d.strftime("%d/%m/%Y")


def cargar_config() -> dict:
    try:
        return json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
    except Exception as e:
        print(f"  [ERROR] No se pudo leer {CONFIG_FILE.name}: {e}")
        return {}


# ── Descubrimiento defensivo del formulario ────────────────────────────────────
async def _dump_formulario(page, etiqueta="stockFecha"):
    try:
        info = await page.evaluate(r"""() => {
          const sels = [...document.querySelectorAll('select')].map(s => ({
            id: s.id, name: s.name,
            opts: [...s.options].slice(0, 20).map(o => `${o.value}=${(o.textContent||'').trim()}`)
          }));
          const inps = [...document.querySelectorAll('input')]
            .filter(i => !['button','submit','image','hidden'].includes(i.type))
            .map(i => ({id: i.id, name: i.name, type: i.type, value: i.value, ph: i.placeholder || ''}));
          const btns = [...document.querySelectorAll('button,a,input[type=button],input[type=submit]')]
            .map(b => (b.textContent || b.value || '').trim()).filter(Boolean).slice(0, 40);
          return {sels, inps, btns, url: location.href};
        }""")
    except Exception as e:
        print(f"  [dump] no se pudo inspeccionar: {e}")
        return
    print(f"  [DESCUBRIR · {etiqueta}] {info['url']}")
    for s in info["sels"]:
        print(f"    <select id='{s['id']}' name='{s['name']}'>  {s['opts']}")
    for i in info["inps"]:
        print(f"    <input id='{i['id']}' name='{i['name']}' type='{i['type']}'> "
              f"value='{i['value']}' ph='{i['ph']}'")
    print(f"    botones/enlaces: {info['btns']}")


async def _seleccionar_bodega(page, texto_bodega: str):
    """Selecciona en el <select> de bodega la opción cuyo texto contiene
    `texto_bodega` (ej. 'FARMACIA AT ABIERTA'). Busca por texto — no por id/value
    fijo — igual que _seleccionar_bodega_at_abierta() de AUTO_SSASUR."""
    return await page.evaluate(
        r"""(texto) => {
          const norm = s => (s || '').toUpperCase().replace(/\s+/g, ' ').trim();
          const objetivo = norm(texto);
          for (const s of document.querySelectorAll('select')) {
            const o = [...s.options].find(o => norm(o.textContent).includes(objetivo));
            if (o) {
              if (s.value !== o.value) {
                s.value = o.value;
                s.dispatchEvent(new Event('change', {bubbles: true}));
              }
              return {sel: s.id || s.name || 'bodega', val: o.value, label: (o.textContent || '').trim()};
            }
          }
          return null;
        }""",
        texto_bodega,
    )


async def _set_fecha(page, fecha_str: str):
    """Rellena el campo de Fecha (dd/mm/aaaa) tipeando de verdad (click + limpiar
    + escribir + Tab). Busca el input por id 'fecha*', o el primer input de texto
    cuyo valor tenga forma de fecha dd/mm/aaaa. Mismo enfoque de tipeo real que
    _set_fechas() de AUTO_SSASUR (los campos con máscara ignoran .value sintético)."""
    handle = await page.evaluate_handle(r"""() => {
      const cand = [...document.querySelectorAll('input')].filter(i =>
        i.type !== 'hidden' && i.type !== 'button' && i.type !== 'submit');
      // 1) por id/name que mencione 'fecha'
      let el = cand.find(i => /fecha/i.test((i.id || '') + ' ' + (i.name || '')));
      // 2) por valor con forma dd/mm/aaaa
      if (!el) el = cand.find(i => /^\d{2}\/\d{2}\/\d{4}$/.test((i.value || '').trim()));
      // 3) input type=date
      if (!el) el = cand.find(i => i.type === 'date');
      return el || null;
    }""")
    el = handle.as_element()
    if not el:
        return False
    try:
        await el.click()
        await el.press("Control+A")
        await el.press("Delete")
        await el.type(fecha_str, delay=30)
        await page.keyboard.press("Tab")
        await page.wait_for_timeout(300)
        real = await el.input_value()
        return real == fecha_str
    except Exception:
        return False


async def _click_generar_xls(page):
    """Clic en 'Generar XLS' (o equivalente). Prueba id conocido del otro reporte
    de stock y, si no, por texto."""
    selectores = (
        "#generarXLS_stock", "#generarXLS", "#generarXLS_salida",
        'button:has-text("Generar XLS")', 'a:has-text("Generar XLS")',
        'button:has-text("Generar XLS")', 'input[value*="XLS" i]',
        'button:has-text("XLS")', 'a:has-text("XLS")',
    )
    for sel in selectores:
        try:
            await page.click(sel, timeout=4_000, force=True)
            return sel
        except Exception:
            continue
    raise RuntimeError(f"No encontré el botón 'Generar XLS' (probé: {selectores})")


def _parse_stock_xlsx(path: Path) -> dict:
    """Lee el xlsx del Reporte Stock en Fecha → {DESCRIPCIÓN_SSASUR: cantidad}.
    Detecta la fila de encabezado buscando una columna de nombre (descripción/
    producto/glosa/artículo) y una de cantidad (cantidad/stock/saldo/existencia).
    Best-effort; {} si no se reconoce el formato."""
    try:
        import openpyxl
    except Exception as e:
        print(f"  [ERROR] falta openpyxl: {e}")
        return {}
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    except Exception as e:
        print(f"  [ERROR] no se pudo abrir {path.name}: {e}")
        return {}
    ws = wb.worksheets[0]
    filas = [list(r) for r in ws.iter_rows(values_only=True)]
    wb.close()

    def _norm(v):
        return str(v).strip().lower() if v is not None else ""

    col_desc = col_cant = header_i = -1
    for i, fila in enumerate(filas[:25]):
        celdas = [_norm(c) for c in fila]
        cd = next((j for j, c in enumerate(celdas)
                   if any(k in c for k in ("descrip", "producto", "glosa", "artículo", "articulo"))), -1)
        cc = next((j for j, c in enumerate(celdas)
                   if any(k in c for k in ("cantidad", "stock", "saldo", "existencia"))), -1)
        if cd != -1 and cc != -1:
            col_desc, col_cant, header_i = cd, cc, i
            break
    if header_i == -1:
        print(f"  [AVISO] {path.name}: no reconocí columnas Descripción/Cantidad.")
        return {}

    stock = {}
    for fila in filas[header_i + 1:]:
        if col_desc >= len(fila) or col_cant >= len(fila):
            continue
        desc = fila[col_desc]
        if desc is None or not str(desc).strip():
            continue
        try:
            cant = float(fila[col_cant]) if fila[col_cant] is not None else 0.0
        except (ValueError, TypeError):
            cant = 0.0
        stock[str(desc).strip()] = int(cant) if cant == int(cant) else cant
    return stock


def _mapear(stock_ssasur: dict, mapeo: dict) -> dict:
    """Traduce {DESCRIPCIÓN_SSASUR: cant} → {'nombre|presentacion': cant} usando
    el mapeo de controlados_config.json. Compara normalizado (mayúsculas, espacios
    colapsados) para tolerar dobles espacios del reporte."""
    def _n(s):
        return " ".join(str(s).upper().split())
    stock_norm = {_n(k): v for k, v in stock_ssasur.items()}
    out = {}
    for key, nombre_ssasur in mapeo.items():
        v = stock_norm.get(_n(nombre_ssasur))
        if v is not None:
            out[key] = v
    return out


async def descargar_stock_controlados(page, fecha: date, config: dict | None = None):
    """Descarga el stock en `fecha` para cada bodega de controlados y escribe
    Conteo_Controlados/stock_sistema.json. Reutiliza `page` YA dentro del módulo
    ABASTECIMIENTO (el llamador debe haber hecho entrar_modulo(page,'ABASTECIMIENTO')).
    Devuelve el Path del JSON, o None si no se pudo generar."""
    config = config or cargar_config()
    farmacias = config.get("farmacias", [])
    bodega_ssasur = config.get("bodega_ssasur", {})
    mapeo = config.get("mapeo_ssasur", {})
    if not mapeo:
        print("  [ERROR] controlados_config.json sin mapeo_ssasur — abortando.")
        return None

    SALIDA_DIR.mkdir(parents=True, exist_ok=True)
    fecha_str = _fmt(fecha)
    por_farmacia = {}

    for f in farmacias:
        fid = f.get("id")
        bodega = bodega_ssasur.get(fid)
        if not bodega:
            continue  # Urgencia u otra sin bodega → se ingresa a mano en la app
        print(f"  [{bodega}] stock al {fecha_str} ...")
        try:
            await page.goto(STOCKFECHA_URL)
            await page.wait_for_load_state("networkidle")
            await page.wait_for_timeout(1_500)

            sel = await _seleccionar_bodega(page, bodega)
            if not sel:
                print(f"    [AVISO] no encontré la bodega '{bodega}' en el select.")
                await _dump_formulario(page, f"stockFecha-{fid}")
                await page.screenshot(path=str(MAESTRO_DIR / "debug_stock_controlados.png"))
                continue
            print(f"    Bodega → {sel['label']}")

            ok_fecha = await _set_fecha(page, fecha_str)
            if not ok_fecha:
                print(f"    [AVISO] no pude fijar la fecha {fecha_str} — sigo con la del formulario.")
            await page.wait_for_timeout(400)

            tmp = SALIDA_DIR / f"_stock_{fid}_{fecha.isoformat()}.xlsx"
            async with page.expect_download(timeout=TIMEOUT_DESCARGA) as dl_info:
                await _click_generar_xls(page)
            dl = await dl_info.value
            await dl.save_as(tmp)
            print(f"    ✓ {tmp.name} ({tmp.stat().st_size // 1024:,} KB)")

            crudo = _parse_stock_xlsx(tmp)
            mapeado = _mapear(crudo, mapeo)
            por_farmacia[fid] = mapeado
            print(f"    → {len(mapeado)}/{len(mapeo)} controlados mapeados "
                  f"(de {len(crudo)} filas del reporte)")
            try:
                tmp.unlink()  # el xlsx crudo trae RUTs de otras bodegas? no; igual no se versiona
            except Exception:
                pass
        except Exception as e:
            print(f"    [ERROR] {bodega}: {e}")
            await page.screenshot(path=str(MAESTRO_DIR / "debug_stock_controlados.png"))

    if not por_farmacia:
        print("  [AVISO] No se obtuvo stock de ninguna bodega — no se escribe el JSON.")
        return None

    payload = {
        "fecha": fecha.isoformat(),
        "generadoEn": datetime.now().isoformat(timespec="seconds"),
        "porFarmacia": por_farmacia,
    }
    SALIDA_JSON.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    total = sum(len(v) for v in por_farmacia.values())
    print(f"  ✓ {SALIDA_JSON.name} — {total} valores, bodegas: {', '.join(por_farmacia)}")
    return SALIDA_JSON


# ── Runner suelto (para probar/backfill sin AUTO_SSASUR) ────────────────────────
async def _standalone(fecha: date):
    try:
        from playwright.async_api import async_playwright
    except ImportError:
        print("[ERROR] Playwright no está instalado. Ejecuta AUTO_SSASUR.bat una vez.")
        return
    # Reutiliza el ingreso al dashboard/módulo de AUTO_SSASUR (importar NO corre su main()).
    from AUTO_SSASUR import entrar_modulo, DASHBOARD_URL, SESSION_FILE

    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=False,
            args=["--start-maximized", "--disable-blink-features=AutomationControlled"],
        )
        if SESSION_FILE.exists():
            context = await browser.new_context(accept_downloads=True, storage_state=str(SESSION_FILE))
        else:
            context = await browser.new_context(accept_downloads=True)
        page = await context.new_page()
        print("Logéate en SSASUR si hace falta (tienes 5 min)...")
        await page.goto(DASHBOARD_URL)
        try:
            await page.wait_for_selector(
                'button:has-text("ABASTECIMIENTO"), div:has-text("ABASTECIMIENTO")',
                timeout=300_000)
        except Exception:
            print("[ERROR] No se detectó el dashboard — ¿login incompleto?")
            await browser.close()
            return
        await context.storage_state(path=str(SESSION_FILE))
        await entrar_modulo(page, "ABASTECIMIENTO")
        await descargar_stock_controlados(page, fecha)
        await browser.close()


if __name__ == "__main__":
    arg = next((a for a in sys.argv[1:] if not a.startswith("--")), None)
    if arg:
        d, m, y = arg.split("/")
        fecha = date(int(y), int(m), int(d))
    else:
        fecha = date.today() - timedelta(days=1)   # día anterior
    print(f"Stock de controlados — fecha consultada: {_fmt(fecha)}")
    asyncio.run(_standalone(fecha))
