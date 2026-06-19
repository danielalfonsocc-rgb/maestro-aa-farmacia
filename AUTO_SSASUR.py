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


def fmt(d: date) -> str:
    return d.strftime("%d/%m/%Y")


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
    """Dashboard → tarjeta RECETA → (modal proyecto) → www.ssasur.cl/receta."""
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


async def main():
    # ── Verificar Playwright ───────────────────────────────────────────────────
    try:
        from playwright.async_api import async_playwright
    except ImportError:
        print("\n[ERROR] Playwright no está instalado.")
        print("  Ejecuta AUTO_SSASUR.bat para instalarlo automáticamente.\n")
        try:
            input("Presiona Enter para cerrar...")
        except EOFError:
            pass
        sys.exit(1)

    # Modo prueba: descarga SOLO recetas y termina (sin stock, maestro ni publicar).
    solo_recetas = "--solo-recetas" in sys.argv
    # No publicar en GitHub al final (útil para corridas de prueba/debug).
    no_publicar  = "--no-publicar" in sys.argv

    today = date.today()
    ayer  = today - timedelta(days=1)

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
        print("\n[1/5] Logéate en SSASUR (tienes 5 minutos)...")
        await page.goto(DASHBOARD_URL)
        try:
            await page.wait_for_selector(
                'button:has-text("ABASTECIMIENTO"), div:has-text("ABASTECIMIENTO")',
                timeout=TIMEOUT_LOGIN,
            )
        except Exception:
            await page.wait_for_function(
                "document.body.innerText.includes('ABASTECIMIENTO')",
                timeout=TIMEOUT_LOGIN,
            )
        await context.storage_state(path=str(SESSION_FILE))
        print("  ✓ Sesión detectada")

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
        print("\n[2/5] Módulo RECETA — informe completo por bloques de 30 días...")
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
                    if n_filas > 0:
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
        #  PASO 3 — ABASTECIMIENTO  (volver al dashboard → entrar → stock)
        # ════════════════════════════════════════════════════════════════════
        print("\n[3/5] Módulo ABASTECIMIENTO...")
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
            print("  Bodega: TODAS")
        except Exception as e:
            print(f"  [AVISO] No se pudo seleccionar bodega TODAS: {e}")

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

    # ── PASO 4 — MAESTRO AA ────────────────────────────────────────────────────
    print("\n[4/5] Actualizando Maestro AA...")
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

        # ── PASO 4b — AUDITORÍA DE PRESCRIPCIÓN (para la pestaña en la nube) ──
        audit_py = MAESTRO_DIR / "auditoria_prescripcion.py"
        if audit_py.exists():
            print("\n  Generando auditoría de prescripción...")
            subprocess.run([sys.executable, str(audit_py)], cwd=str(MAESTRO_DIR), env=env_utf8)

        # ── PASO 5 — PUBLICAR EN GITHUB ───────────────────────────────────────
        git_dir  = MAESTRO_DIR / ".git"
        publicar = MAESTRO_DIR / "PUBLICAR_DATOS.bat"
        if no_publicar:
            print("\n[5/5] --no-publicar: omito la publicación en GitHub (modo prueba).")
        elif git_dir.exists() and publicar.exists():
            print("\n[5/5] Publicando datos en GitHub...")
            pub = subprocess.run(
                ["cmd", "/c", str(publicar)],
                cwd=str(MAESTRO_DIR),
            )
            if pub.returncode != 0:
                print("  [AVISO] Publicación falló — ejecuta PUBLICAR_DATOS.bat manualmente.")
        else:
            print("\n[5/5] GitHub no configurado — omitiendo publicación.")
            print("      Ejecuta CONFIGURAR_GITHUB.bat para activarlo.")
    else:
        print("  [ERROR] maestro_aa.py falló — revisa los mensajes arriba")

    try:
        input("\nPresiona Enter para cerrar...")
    except EOFError:
        pass


if __name__ == "__main__":
    asyncio.run(main())
