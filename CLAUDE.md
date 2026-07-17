# Proyecto: Maestro AA — Consolidado Operacional Farmacia AT Abierta

Herramienta de gestión farmacéutica para Hospital de Pitrufquén (SSASur).
Universo: **378 medicamentos AA**. Fuente de datos: SSASur (stock + recetas).

## Archivos principales

| Archivo | Rol |
|---|---|
| `maestro_aa.py` | Consolidación principal → `Consolidado_AA_MAESTRO.xlsx` (14 hojas) |
| `app_pedidos.py` | Dashboard Streamlit — pedidos, faltantes, alertas |
| `sgli.py` | Motor SGLI: reposición basada en demanda (sin techo de capacidad) |
| `utils_aa.py` | **Módulo compartido**: norm_erp, HOMOLOGACION (20 entradas), cargar_recetas_csv |
| `cruce_gt.py` | Cruce con Guías de Tratamiento |
| `recetas_cheque.py` | Formulario ISP recetas cheque (estupefacientes/psicotrópicos) — obligación legal |
| `agente_duplicados.py` | Agente IA (Haiku) — duplicados operacionales del día |
| `agente_gt_pendientes.py` | Agente IA (Haiku) — clasifica PENDIENTES de GT (URGENTE/RUTINARIO/DIFERIBLE) |
| `auditoria_medicamento.py` | Auditoría clínica genérica: `--contiene NOMBRE --dosis X` |
| `auditoria_duplicados_profunda.py` | Auditoría duplicados histórica con vigencia actual y propuestas IA |
| `auditoria_prescripcion.py` | Pre-calcula `auditoria_prescripcion.json` (consumido por app_pedidos) |
| `pedido_fusion.py` | Genera Pedido_Fusion_AA_<fecha>.xlsx (Farm_Bod + Bod_Farmacos + Dialisis) |
| `programacion_aa.py` | Planilla ciclo Bodega AA: Cantidad Programada/Solicitada (reporte SSASUR) vs Stock Bodega AA vs Stock Real (conteo). `--aplicar-conteo` genera el Resumen final en `Programacion_AA\`. Sin IA |
| `sgli_historico.py` | Planilla SGLI histórica — clasificación ABC-XYZ |
| `centinela_reporte.py` | Reporte semanal centinela campaña invierno (PDF MINSAL) |
| `AUTO_SSASUR.py` | Descarga automatizada SSASur (recetas + stock + GT) → dedup → Drive |
| `dedup_recetas.py` | Detecta/limpia recetas duplicadas por sobre-extracción de GT (crea .bak) |
| `publicar_drive.py` | Sube salidas a Google Drive (requiere `credentials.json` + `SETUP_DRIVE.bat`) |
| `publicar_escritorio.py` | Copia salidas al Escritorio\Farmacia AA\ (acceso rápido local) |
| `aa_colors.py` | Paleta de colores compartida (impresión económica) |
| `_generar_glosario.py` | Genera Glosario_Maestro_AA.pdf |
| `skill_gt/scripts/generar.py` | Generador de planillas + letreros GT por establecimiento destino |

## Reglas de arquitectura

- **Nuevas homologaciones de nombres**: SOLO en `utils_aa.py → HOMOLOGACION_RAW`. Nunca duplicar en scripts individuales.
- **El modelo SGLI no tiene techo de capacidad**: el Nivel Objetivo T se calcula desde la demanda; Cap_Max es informativo y solo activa [ALERTA_ESTRES].
- **RUTs**: nunca a la API. `agente_duplicados.py` anonimiza con SHA-256 antes de llamar a Claude.
- **GT raw downloads**: van a `../04_Farmacia_Gestion_Territorial/` (carpeta hermana del repo). Nombrado: `reporteGestionTerritorial_<desde>_<hasta>.xlsx`. `dedup_recetas.py` busca ahí.
- **Drive**: NO subir CSV sábanas ni stock xlsx (RUTs / Ley 19.628). Carpeta raíz `Farmacia AA` en Drive — IDs fijos en `_drive_folders.json`. Para activar: `SETUP_DRIVE.bat`.
- **Rutas fuera del repo (otra máquina)**: `RCH_DIR` y `PLANTILLA_BLANCO_RCH` (carpeta/plantilla del formulario ISP de Recetas Cheque) viven en `utils_aa.py`, configurables por variable de entorno — `MAESTRO_RCH_DIR` y `MAESTRO_PLANTILLA_RCH` — para no hardcodear la ruta de la QF al correr esto en otro equipo. Default = la ruta actual de esta máquina.

## Economía de modelos (OBLIGATORIO respetar)

### Haiku — tareas mecánicas (<1 min de razonamiento)
- Preguntas "¿qué hace esta función?"
- Verificar un import o nombre de variable
- Formatear o convertir una lista pequeña
- "¿Cuál es el dtype de esta columna pandas?"
- Dudas de sintaxis Python/Streamlit puntuales
- `agente_duplicados.py` usa Haiku por defecto (análisis mecánico-analítico)

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
