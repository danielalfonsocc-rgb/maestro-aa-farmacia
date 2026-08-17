---
name: inventario-programacion
description: >
  Genera el instrumento de conteo físico de Bodega AA y calcula la diferencia
  contra Stock Sistema y Cantidad Programada/Solicitada, entregando al final
  la lista de qué pedir — SIN necesitar el Consolidado_AA_MAESTRO.xlsx ni el
  historial completo de recetas. Pensado para cuando el usuario NO tiene el
  repo Maestro AA clonado ni Python corriendo localmente (otro computador,
  celular, claude.ai) y solo cuenta con los dos reportes crudos de SSASUR:
  el reporte de Programación (Consumo por centro de costo → FARMACIA) y el
  reporte de stock de bodega.

  USAR CUANDO el usuario diga o implique cualquiera de estas cosas:
  "no puedo usar Maestro AA en este computador", "necesito hacer el pedido
  hoy pero no tengo el repo acá", "sácame el inventario de la programación",
  "instrumento para contar la bodega", "planilla para contar físicamente",
  "compara el stock del sistema con lo que cuento", "diferencia entre stock
  y programación", "qué pedir sin el consolidado", "tengo el reporte de
  programación y el de stock, ayúdame a armar el pedido", "hoy toca hacer
  pedido y estoy sin mi compu de siempre". También úsalo proactivamente
  cuando el usuario adjunte o pegue datos de un reporte de "cantidad de
  productos consumidos en centro de costo farmacia" y/o un "reporte de
  stock" y hable de contar, inventariar o pedir.
---

# Skill: inventario-programacion
# Instrumento de conteo + diferencia + pedido — Bodega AA (sin Consolidado)

Versión liviana de `programacion_aa.py`, para el día a día en que el usuario
(QF de Farmacia AT Abierta, Hospital de Pitrufquén) necesita hacer el pedido
del ciclo Bodega AA pero está en una máquina sin el repo Maestro AA clonado,
sin Python, y sin el historial de 9 meses de recetas que alimenta el
Consolidado. Solo hacen falta dos reportes que se descargan directo de
SSASUR y que cualquiera puede tener a mano ese día:

1. **Reporte de Programación**: SSASUR → Reportes → Consumo por centro de
   costo → Centro de Costo = FARMACIA → Generar XLS.
   Da `Cantidad Programada` / `Cantidad Solicitada` del ciclo.
2. **Reporte de stock**: `reporte_de_stock_*.xlsx` (volcado crudo con
   columna Bodega) o el reporte **"Existencias"** (ya acotado a una sola
   bodega/farmacia, columnas Producto/Stock Disponible, sin columna Bodega
   — el script detecta cuál es cuál). Da `Stock Sistema` de la bodega (por
   defecto BODEGA AT ABIERTA).

El script `inventario_rapido.py` (en la raíz del proyecto) hace el trabajo
en dos pasos. **No calcula Consumo Promedio Mensual ni Sugerencia de
programación** — eso requiere el historial completo vía `maestro_aa.py` /
`sgli.py`, que aquí no está disponible. Lo que sí entrega:

**Importante — universo AA**: los dos reportes crudos de SSASUR traen de
todo (inyectables de pabellón, hospitalización cerrada, etc.), no solo lo
que Farmacia AA dispensa. Si en la carpeta del proyecto hay un
`Consolidado_AA_MAESTRO*.xlsx` (se autodetecta el más reciente, aunque sea
de días atrás — no hace falta que esté fresco del día), el script lo usa
para acotar el instrumento a la lista real de medicamentos AA (hoja
`Pedido_Repos_Bodega`); lo que quede fuera de ese universo se guarda en una
hoja aparte `Fuera_Universo_AA` del mismo Excel, no se pierde ni se mezcla
con el conteo. **Si no hay Consolidado disponible, avísale al usuario
explícitamente** que el instrumento es la unión cruda de ambos reportes y
puede traer medicamentos que Farmacia AA no dispensa — no lo des por
"limpio" sin decirlo.

- Un **instrumento de conteo** imprimible/editable, medicamento por
  medicamento, con una columna en blanco para el conteo físico.
- La **Diferencia** entre lo que dice el sistema y lo que se contó
  físicamente (para pescar errores de inventario, mermas, etc.).
- La **Cantidad a Pedir**: `max(0, objetivo − Stock Real)`, donde objetivo
  es Cantidad Programada (o Solicitada si no hay Programada). Es una
  fórmula simple de "reponer hasta lo programado del ciclo" — no reemplaza
  el modelo SGLI completo, pero alcanza para armar el pedido de hoy.

## Paso 1 — conseguir los dos reportes

Pídeselos al usuario si no los ha adjuntado todavía. El orden de preferencia:

1. **Archivos adjuntos** (.xlsx tal como los bajó de SSASUR) — lo ideal,
   porque el script espera el formato exacto de SSASUR (fila 1 = título,
   fila 2 = metadata, fila 3 = encabezados). Guárdalos en el filesystem
   (p.ej. con el tool de escritura de archivos que tengas disponible según
   el entorno) y pásale la ruta al script.
2. Si el usuario solo puede **pegar el contenido como texto/tabla** en el
   chat (porque está en un celular o no puede adjuntar), reconstruye tú
   mismo un .xlsx que respete el formato esperado (fila 1 título cualquiera,
   fila 2 con "...mes de <MES> de <AÑO>..." si lo tiene, fila 3 con los
   encabezados `Centro Costo | Producto | Total de Productos Programados |
   Productos Solicitado` para el reporte de Programación, o `Descripcion |
   Bodega | Cantidad` — el orden exacto de columnas no importa, los nombres
   sí — para el de stock) antes de correr el script. Explícale brevemente
   al usuario que hiciste esto para que no le extrañe si algo no calza.
3. Si falta alguno de los dos reportes, dile exactamente qué descargar y
   desde dónde (ver arriba) — no inventes números.

## Paso 2 — generar el instrumento de conteo

```bash
cd /home/user/maestro-aa-farmacia   # o la raíz del repo en este entorno
python3 inventario_rapido.py --programacion ruta_reporte_prog.xlsx --stock ruta_reporte_stock.xlsx
```

Si la bodega del reporte de stock no es "BODEGA AT ABIERTA" (por ejemplo el
usuario trabaja en otra farmacia/bodega de la red), el script se detiene y
lista las bodegas que sí encontró en el archivo — vuelve a correrlo con
`--bodega "NOMBRE EXACTO"`.

El script deja `Instrumento_Conteo_AA_<fecha>.xlsx` en `Programacion_AA/`.
**Entrégaselo al usuario** (con la herramienta de envío de archivos que
tengas disponible) y explícale en 2-3 líneas:
- Cuántos medicamentos quedaron en la planilla, y cuántos aparecen "solo en
  Stock" (tienen inventario pero no están en el reporte de Programación —
  vale la pena que los revise, puede que falte agregarlos a la programación).
- Que debe contar físicamente Bodega AA e ir llenando la columna
  "Stock Real" — directo en el Excel (si tiene Excel/Sheets a mano) o
  avisándote los números por chat, lo que le sea más cómodo.

## Paso 3 — aplicar el conteo y sacar la lista de pedido

Cuando el usuario tenga los números contados, hay dos caminos — usa el que
corresponda según cómo te los pase:

**A) Te reenvía el mismo Excel con "Stock Real" ya lleno:**
```bash
python3 inventario_rapido.py --aplicar-conteo Instrumento_Conteo_AA_<fecha>.xlsx
```

**B) Te dicta o pega los conteos por chat** (p.ej. "Paracetamol 500: 90,
Ibuprofeno 400: 310..."): arma un JSON `{"MEDICAMENTO TAL COMO SALE EN LA
PLANILLA": cantidad, ...}` con exactamente los nombres que aparecen en la
columna Medicamento del instrumento (cópialos de ahí, no los reescribas) y
corre:
```bash
python3 inventario_rapido.py --aplicar-conteo conteo.json
```
No hace falta que el usuario cuente todo de una vez — puedes correr este
paso varias veces a medida que te va pasando más medicamentos; los que ya
tenían Stock Real en el Excel base se mantienen y solo se sobreescriben los
que vengan en el JSON nuevo.

Esto genera `Resumen_Conteo_AA_<fecha_hora>.xlsx` en `Programacion_AA/`,
con Diferencia y Cantidad a Pedir resaltadas. **Entrégaselo al usuario** y
además de mandarle el archivo, **responde en el chat con la lista de qué
pedir** (el script ya la imprime ordenada de mayor a menor cantidad) — el
usuario no debería tener que abrir el Excel solo para saber qué pedir hoy.
Menciona también:
- Cuántos medicamentos quedaron "sin contar todavía" (por si se le olvidó
  alguno).
- Los medicamentos con Diferencia grande entre Stock Sistema y Stock Real
  (posible error de inventario/merma), aunque no requieran pedido.

## Ejemplo de resumen para el chat

```
Instrumento generado: 342 medicamentos (338 en Programación, 4 solo en Stock
— revisa si conviene agregarlos a la programación).

[usuario cuenta y te pasa los números]

Resumen aplicado — 12 sin contar todavía.
Diferencias stock vs sistema: 27 medicamentos.

Qué pedir hoy (17 medicamentos, 2.340 unidades):
  - PARACETAMOL 500 MG COMPRIMIDO: 410 ud
  - OMEPRAZOL 20 MG CAPSULA: 125 ud
  - ...
```

## Errores comunes

| Error | Causa | Solución |
|---|---|---|
| `El reporte de Programación no tiene las columnas esperadas` | El archivo no es el reporte correcto, o se pegó/reconstruyó mal el encabezado | Verificar que sea "Consumo por centro de costo → Generar XLS"; revisar fila 3 = encabezados |
| `No se encontró la bodega "..." en el reporte de stock` | El reporte de stock no tiene "BODEGA AT ABIERTA" (otra farmacia/bodega) | El script lista las bodegas disponibles — reintentar con `--bodega "NOMBRE EXACTO"` |
| Un medicamento sale "solo en Stock" (sin Cantidad Programada) | No está en el reporte de Programación de este ciclo | Normal — revisar si conviene incorporarlo a la programación (sin el historial CMP no se puede confirmar automáticamente, a diferencia de `programacion_aa.py`) |
| Los números no coinciden con lo esperado | Nombres de medicamento con grafías distintas entre los dos reportes | Revisar si hace falta agregar una homologación nueva en `utils_aa.py → HOMOLOGACION_RAW` (nunca duplicarla en otro script) |

## Cuándo usar el flujo completo en vez de este

Si el usuario SÍ tiene el repo Maestro AA con el Consolidado_AA_MAESTRO.xlsx
generado (p.ej. está en su computador habitual), es mejor usar
`programacion_aa.py` directamente — da además Consumo Promedio Mensual y
Sugerencia de programación basada en el historial real de recetas. Este
skill es específicamente para cuando esa opción no está disponible.
