# Proyecto: Maestro AA — Consolidado Operacional Farmacia AT Abierta

Herramienta de gestión farmacéutica para Hospital de Pitrufquén (SSASur).
Universo: **378 medicamentos AA**. Fuente de datos: SSASur (stock + recetas).

## Archivos principales

| Archivo | Rol |
|---|---|
| `maestro_aa.py` | Consolidación principal → `Consolidado_AA_MAESTRO.xlsx` (14 hojas). Infraestructura: alimenta Pedido Fusión y el resto de las 6 categorías, aunque ya no tiene dashboard propio |
| `sgli.py` / `sgli_historico.py` | Motor SGLI (reposición basada en demanda) y su planilla histórica ABC-XYZ. Dependencia interna de `maestro_aa.py` |
| `utils_aa.py` | **Módulo compartido**: norm_erp, HOMOLOGACION (20 entradas), cargar_recetas_csv |
| `cruce_gt.py`, `gt_maestro.py`, `agente_gt_pendientes.py`, `dedup_recetas.py`, `procesar_establecimiento_maestro.py`, `descargar_recetas_pdf.py`, `publicar_gt_sheets.py` y demás scripts `gt_*`/`*_gt_*` | Bloque de **Gestión Territorial** (una de las 6 categorías) |
| `recetas_cheque.py`, `subir_recetas_cheque_drive.py` | **Controlados**: formulario ISP recetas cheque (estupefacientes/psicotrópicos) — obligación legal |
| `pedido_fusion.py`, `pedido_fusion_simple.py` | **Fusión AA**: Pedido_Fusion_AA_<fecha>.xlsx (Farm_Bod + Bod_Farmacos + Dialisis + Faltantes_AA) |
| `centinela_reporte.py` | **Centinela Invierno**: reporte semanal campaña invierno (PDF MINSAL) |
| `centinela_inyectables_sm.py` | **Centinela SM**: stock de antipsicóticos de depósito (salud mental ambulatoria) |
| `clozapina_consolidar.py`, `clozapina_hce_hemogramas.py`, `motor_reglas_clozapina_v3.py` | **Clozapinas**: consolidado de hemogramas MINSAL |
| `servicios_farmaceuticos.py` | **Servicios Farmacéuticos**: recuento mensual QF desde "Hoja Diaria de Profesional" (Agenda Médica SSASUR) → `Servicios_Farmaceuticos/<MES AÑO>/Resumen_Servicios_Farmaceuticos_*.xlsx`, SIN RUT. Lo llama AUTO_SSASUR |
| `AUTO_SSASUR.py` | Descarga automatizada SSASur (recetas + stock + GT) → dedup → Drive |
| `publicar_drive.py` | Sube salidas a Google Drive (requiere `credentials.json` + `SETUP_DRIVE.bat`) — recortado 04-09-2026 a solo las 6 categorías |
| `publicar_escritorio.py` | Copia salidas al Escritorio\Farmacia AA\ (acceso rápido local) — recortado 04-09-2026 a solo las 6 categorías |
| `aa_colors.py` | Paleta de colores compartida (impresión económica) |
| `skill_gt/scripts/generar.py` | Generador de planillas + letreros GT por establecimiento destino |

**Eliminados 04-09-2026** (dashboard Streamlit + auditorías/programación sueltas, fuera de las 6 categorías que el usuario pidió conservar): `app_pedidos.py`, `app_maestro.py` (+ `paginas/*.py` y `estilo_maestro.py`, exclusivos del hub), `agente_duplicados.py`, `auditoria_medicamento.py`, `auditoria_duplicados_profunda.py`, `auditoria_prescripcion.py`, `auditoria_cantidad_posologia.py`, `auditoria_insulinas.py`, `app_auditoria_retiros.py`, `programacion_aa.py`, `actualizar_programacion_sept.py`, `_generar_glosario.py`, `crear_acta_vencimiento.py`, `subir_prueba_sheets.py`. No reintroducir sin confirmar con el usuario (mismo criterio que otras eliminaciones — ver memoria del proyecto).

**Eliminado 07-09-2026:** `fusionar_nominas_gt.py` (y el PASO 5c2 que lo llamaba en `AUTO_SSASUR.py`). Fusionaba en un paso posterior, dependiente del orden, la Planilla "automática" (`cruce_gt.py`) con la "manual" (`agregar_gt_manual.py`) — causaba el bug real de 2 nóminas separadas del 04-09-2026 (ver memoria `gt-manual-vs-pipeline-auto`). Reemplazado por fusión al momento de escribir en ambos lados: `agregar_gt_manual.py` ahora fusiona directo sobre `<destino>_Planilla.xlsx` (nunca crea `Nomina_Manual_*`), y `publicar_drive._depositar_arbol_local()` fusiona sola si la Planilla automática llega después el mismo día. No reintroducir sin confirmar.

## Reglas de arquitectura

- **Nuevas homologaciones de nombres**: SOLO en `utils_aa.py → HOMOLOGACION_RAW`. Nunca duplicar en scripts individuales.
- **El modelo SGLI no tiene techo de capacidad**: el Nivel Objetivo T se calcula desde la demanda; Cap_Max es informativo y solo activa [ALERTA_ESTRES].
- **RUTs**: nunca a la API. `agente_gt_pendientes.py` no incluye RUT ni nombre en el prompt (solo medicamento/fecha/estado).
- **GT raw downloads**: van a `../04_Farmacia_Gestion_Territorial/` (carpeta hermana del repo). Nombrado: `reporteGestionTerritorial_<desde>_<hasta>.xlsx`. `dedup_recetas.py` busca ahí.
- **Reporte de Programación AA (PASO 4b de `AUTO_SSASUR.py`)**: sigue descargándose a diario (es parte del mismo scrape del módulo ABASTECIMIENTO, sin costo aparte) aunque `programacion_aa.py` ya no existe — nadie lo procesa hoy. Se dejó así porque tocar el scraper es riesgoso y de bajo beneficio; si se quiere ahorrar el tiempo de descarga, usar `--no-programacion` explícitamente.
- **Drive/Escritorio recortados 04-09-2026**: `publicar_drive.py` y `publicar_escritorio.py` solo sincronizan las 6 categorías vigentes (Fusión AA, Gestión Territorial, Controlados, Clozapinas, Centinela Invierno/SM, Servicios Farmacéuticos) + infraestructura. Las carpetas "App Pedidos", "Auditoria Prescripcion" y "Programacion AA" (en Drive y en Escritorio\Farmacia AA\) se eliminaron — no reintroducir sin confirmar.
- **Ventana del informe GT**: la calcula `AUTO_SSASUR.gt_desde_a_consultar()` a
  partir de `_gt_cobertura.json` (qué fecha de despacho se consultó y en qué día),
  NO de `_ultima_corrida_ok.json`. Retrocede hasta el primer día hábil que nadie
  consultó *después* de que terminara, y siempre re-consulta los últimos
  `GT_RECHEQUEO_DIAS` (7), con tope `GT_MAX_RETRO_DIAS` (30). Motivo: la ventana
  de un día se consulta una sola vez a media mañana, y lo que SSASur registre
  después con esa misma fecha de entrega no lo veía nadie nunca más; además, si
  el programa no corría un día, ese día se perdía. Un rango ancho es inofensivo
  porque el dedup por nómina real impide reimprimir. La cobertura se anota solo
  con `n >= 0` (incluido "0 filas"): si la descarga falla, el día queda ABIERTO
  a propósito para que la próxima corrida lo recupere.
- **Dedup GT (qué receta ya tiene nómina)**: la fuente es la PLANILLA en disco
  (`cruce_gt._recetas_con_nomina`: `out_gt/**/*Planilla*.xlsx`, el árbol
  `<ESTAB>/Nóminas de Envío/**` y los `Nomina_Manual_*` legados), reforzada con
  las recetas que `gt_maestro.xlsx` marca en un estado avanzado (lista para
  retiro / enviada / entregada). NUNCA usar la mera presencia de la fila en
  `gt_maestro.xlsx`: ese libro es el registro anual de GT y lo llena
  `_sincronizar_maestro()` con el reporte crudo COMPLETO en estado "EN
  PREPARACIÓN" — usarlo como dedup dejó recetas de Quepe/Teodoro/Toltén sin
  nómina para siempre (corregido 11-09-2026).
- **Orden cronológico de los reportes GT**: se deriva de la fecha `<desde>` del
  nombre (ancla del día de descarga), NUNCA de `<hasta>`, que es el borde de la
  ventana consultada y puede apuntar al futuro. `dedup_recetas._clave_orden` y
  `publicar_drive.sync_gt` ya lo hacen así; usar `<hasta>` hacía que el dedup
  vaciara el reporte recién descargado.
- **Drive**: NO subir CSV sábanas ni stock xlsx (RUTs / Ley 19.628). Carpeta raíz `Farmacia AA` en Drive — IDs fijos en `_drive_folders.json`. Para activar: `SETUP_DRIVE.bat`.
- **Rutas fuera del repo (otra máquina)**: `RCH_DIR` y `PLANTILLA_BLANCO_RCH` (carpeta/plantilla del formulario ISP de Recetas Cheque) viven en `utils_aa.py`, configurables por variable de entorno — `MAESTRO_RCH_DIR` y `MAESTRO_PLANTILLA_RCH` — para no hardcodear la ruta de la QF al correr esto en otro equipo. Default = la ruta actual de esta máquina.

## Economía de modelos (OBLIGATORIO respetar)

### Haiku — tareas mecánicas (<1 min de razonamiento)
- Preguntas "¿qué hace esta función?"
- Verificar un import o nombre de variable
- Formatear o convertir una lista pequeña
- "¿Cuál es el dtype de esta columna pandas?"
- Dudas de sintaxis Python/Streamlit puntuales

### Sonnet — trabajo estándar (mayoría de las tareas)
- Agregar/modificar columnas en `maestro_aa.py`
- Nuevos filtros o vistas en `app_pedidos.py`
- Ajustar CSS/layout Streamlit
- Bugs predecibles (KeyError, dtype mismatch, merge incorrecto)
- Scripts nuevos similares a los existentes (nueva auditoría, nuevo cruce)
- Ajustar umbrales o fórmulas en `sgli.py`
- Generación de nuevas hojas Excel con ReportLab/openpyxl

### Opus — solo razonamiento complejo
- Rediseñar la lógica SGLI desde cero
- Diseñar una nueva fórmula de reposición con múltiples variables
- Automatización SSASur cuando el sitio cambia (web scraping complejo)
- Refactorizar entre varios archivos simultáneamente
- Decisiones de arquitectura nuevas

## Reglas de prompt para reducir coste

### 1. Declarar el archivo al inicio
```
Archivo: app_pedidos.py, función cargar_datos(), línea ~180.
Problema: [descripción concreta].
Solución esperada: [qué cambio quiero].
```

### 2. No pegar DataFrames enteros
```python
# Bien: df.dtypes, df.shape, df.head(3)
```

### 3. Reutilizar contexto de sesión
Dentro de una sesión larga, no re-explicar el proyecto. El contexto está cacheado.

### 4. Limitar verbosidad en sesiones de diagnóstico
Para tareas del tipo "detecte problemas" o "corríjelo":
```
Diagnóstico en máx. 4 bullet points.
Corrección: solo el bloque de código cambiado, sin re-imprimir el archivo completo.
Sin explicaciones de lo que no cambió.
```
El ratio output/input fue 65× en la última jornada — tokens de salida cuestan
3–5× más que los de entrada. Un diagnóstico verboso de 10 K tokens sale igual
que 10 preguntas de Haiku.

## Privacidad y datos sensibles

Los archivos `informe_completo_recetas*.csv` y `reporte_de_stock_*.xlsx`
pueden contener **RUT de pacientes** sujetos a la **Ley 19.628**.

- **NUNCA** pegar filas con RUT en el chat.
- Para debugging con datos reales: anonimizar primero (reemplazar RUT por `XXXXX`).
- El directorio `maestro/` no debe versionarse en GitHub público.

## Stack técnico

- Python 3.10 vía lanzador `py`
- Streamlit ≥ 1.35 en puerto **8501**
- pandas, numpy, openpyxl, reportlab, rapidfuzz, anthropic

## Iniciar el servidor

```bat
py -m streamlit run app_pedidos.py --server.headless true --browser.gatherUsageStats false --server.port 8501
```

O simplemente: `ABRIR_APP.bat`
