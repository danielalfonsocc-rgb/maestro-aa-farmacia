---
name: qa-app-pedidos-pitrufquen
description: >
  Verificación por regresión de app_pedidos.py / app_maestro.py contra el
  Consolidado_AA_MAESTRO.xlsx y los scripts fuente (pedido_fusion.py, maestro_aa.py)
  tras un cambio de código: levanta el servidor Streamlit, recorre cada pestaña,
  cruza las cifras contra el output del script fuente y revisa la consola del
  navegador en busca de errores. Reemplaza el chequeo manual pestaña-por-pestaña
  que se repite después de casi cualquier cambio a la app.

  USAR CUANDO el usuario diga o implique:
  "revisa que quede todo bien en la app", "verifica la app después del cambio",
  "compara app vs Excel", "revisa la app_pedidos", "chequea que la app cargue bien",
  "¿la app sigue coincidiendo con el Excel?", "prueba la app".

  USAR PROACTIVAMENTE (sin que lo pidan) después de modificar app_pedidos.py,
  app_maestro.py, pedido_fusion.py, pedido_fusion_simple.py o maestro_aa.py en la
  misma sesión — antes de dar el cambio por terminado.
---

# Skill: qa-app-pedidos-pitrufquen
# QA de regresión — app Streamlit vs. Excel/scripts fuente

Después de tocar `app_pedidos.py`, `app_maestro.py`, `pedido_fusion.py`,
`pedido_fusion_simple.py` o `maestro_aa.py`, la app puede seguir "cargando bien"
visualmente y aun así mostrar cifras desincronizadas del Excel o de los scripts
fuente que la alimentan — porque el bug está en un cálculo, no en la sintaxis.
Este skill automatiza el chequeo cruzado que si no, hay que hacer a mano
pestaña por pestaña cada vez.

No genera ningún archivo nuevo: es control de calidad, no una tarea de reporte.
El resultado es un veredicto pestaña-por-pestaña.

## Cuándo NO usarlo

- Si el cambio fue puramente cosmético (CSS, texto, orden de columnas visual)
  y no toca ningún cálculo, un `py_compile` + una mirada rápida basta —no hace
  falta el cruce completo de cifras.
- Si no hay datos reales disponibles (`Consolidado_AA_MAESTRO.xlsx` no existe
  o está vacío), avisa al usuario en vez de forzar el chequeo.

## Pasos

### 1 — Sanity check de sintaxis

`py_compile` (o simplemente `py -c "import ast; ast.parse(open('archivo.py').read())"`)
de cada módulo tocado. Si falla acá, ni vale la pena levantar el servidor.

### 2 — Levantar el servidor Streamlit

Usar `preview_start` (nunca `Bash`/`PowerShell` para el server — ver reglas
generales de la herramienta).

- Si se tocó solo `app_pedidos.py` / `pedido_fusion*.py` / `maestro_aa.py`:
  levantar `app_pedidos.py` en el puerto **8501**
  (`py -m streamlit run app_pedidos.py --server.headless true --browser.gatherUsageStats false --server.port 8501`).
- Si se tocó `app_maestro.py` o algo bajo `paginas/`: levantar `app_maestro.py`
  en el puerto **8502**. `app_maestro.py` es un hub (`st.navigation`) que
  embebe `app_pedidos.py` como página "Pedidos AA" — si el cambio afecta
  ambos, chequear ambos puertos.

Revisar `preview_logs` inmediatamente: un traceback de arranque (import roto,
columna faltante en el Excel) corta el chequeo ahí — repórtalo y no sigas
navegando pestañas fantasma.

### 3 — Recorrer las pestañas

`app_pedidos.py` tiene estas pestañas (confirmar en el código si cambiaron —
no asumir esta lista ciega si el diff tocó la definición de `st.tabs`):

| Pestaña | Qué valida |
|---|---|
| 📝 Pedido a Bodega AA | Farm→Bod: nº de meds y unidades a pedir |
| 🏭 Pedido a Bodega Fármacos | Bod→Fármacos: nº de meds con reposición |
| 💉 Diálisis | meds con faltante de diálisis (días **corridos**, no hábiles) |
| 🚨 Faltantes | meds sin poder despachar en Atención Abierta |
| ⛔ Faltantes Absolutos (30d) | meds con Bodega AA en 0 |
| 🚦 SGLI · Capacidad | alertas de estrés / nivel objetivo |
| 🔬 Auditoría de prescripción | hallazgos de `auditoria_prescripcion.json` |
| 💬 Diagnóstico y Sugerencias | no lleva cifras — solo confirmar que carga sin error |

Para cada una: `navigate`/click en la pestaña → `get_page_text` → anotar los
totales (nº de medicamentos, unidades) que muestre el encabezado o resumen de
la tabla.

### 4 — Cruzar contra el script fuente

Correr en paralelo el script fuente equivalente y comparar sus totales impresos
contra lo que muestra la pestaña:

```bash
py pedido_fusion.py
```

Esto imprime, entre otras líneas:
```
Farm->Bod       : N meds (n con pedido)
Bod->Farmacos   : N meds (n con reposicion)
Dialisis        : N meds (n con faltante)
Faltantes AA 30d: N meds sin poder despachar en Atencion Abierta
Por agotarse    : N meds con Bodega AA en 0 y cobertura ...
```

Estos números deben calzar exactamente con lo que la pestaña correspondiente
muestra (📝 ↔ Farm->Bod, 🏭 ↔ Bod->Farmacos, 💉 ↔ Dialisis, 🚨 ↔ Faltantes AA 30d,
⛔ ↔ Por agotarse). Un desfase acá es la señal real de un bug — no un logging
extra, es la razón de ser de este skill.

### 5 — Consola del navegador

`read_console_messages` con `onlyErrors: true` después de recorrer todas las
pestañas. Cualquier error JS/red se reporta aunque las cifras hayan calzado.

### 6 — Cerrar el servidor

`preview_stop` al terminar, para no dejar el puerto ocupado en la siguiente
sesión.

## Formato de reporte

No repetir el diff ni re-explicar qué se cambió — solo el veredicto:

```
## QA app vs. Excel/script — <archivo(s) tocado(s)>

| Pestaña | Carga | Cifras | Consola |
|---|---|---|---|
| 📝 Pedido a Bodega AA | ✅ | ✅ 91 meds / 44.683 ud | ✅ |
| 🏭 Pedido a Bodega Fármacos | ✅ | ❌ app: 38 meds / script: 41 meds | ✅ |
| ... | | | |

Desfases encontrados: <detalle solo de lo que no calzó, con la línea de código sospechosa si se identifica>
```

Si todo calza, un solo bloque corto basta ("Las 8 pestañas cargan bien y las
cifras coinciden con `pedido_fusion.py`. Sin errores de consola.") — no hace
falta la tabla completa cuando no hay nada que señalar.

## Datos sensibles

Los CSV de recetas y el reporte de stock pueden traer RUT de pacientes.
`get_page_text` puede exponer filas con RUT si la pestaña las muestra
(p.ej. Auditoría de prescripción) — no reproducir esas filas en el reporte,
solo los totales agregados.
