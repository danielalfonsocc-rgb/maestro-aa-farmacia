# Proyecto: Maestro AA — Consolidado Operacional Farmacia AT Abierta

Herramienta de gestión farmacéutica para Hospital de Pitrufquén (SSASur).
Universo: **378 medicamentos AA**. Fuente de datos: SSASur (stock + recetas).

## Archivos principales

| Archivo | Rol |
|---|---|
| `maestro_aa.py` | Consolidación principal → `Consolidado_AA_MAESTRO.xlsx` (14 hojas) |
| `app_pedidos.py` | Dashboard Streamlit — pedidos, faltantes, alertas |
| `sgli.py` | Motor SGLI: reposición basada en demanda (sin techo de capacidad) |
| `reposicion_dias_habiles.py` | Reposición ajustada por feriados y Freq_Revision |
| `utils_aa.py` | **Módulo compartido**: norm_erp, HOMOLOGACION (20 entradas), cargar_recetas_csv |
| `cruce_gt.py` | Cruce con Guías de Tratamiento |
| `recetas_cheque.py` / `recetas_duplicadas.py` | Auditorías de recetas (ISP / histórico) |
| `agente_duplicados.py` | Agente IA (Haiku) — duplicados operacionales del día |
| `auditoria_medicamento.py` | Auditoría clínica genérica para cualquier medicamento |
| `auditoria_prescripcion.py` / `auditoria_empagliflozina.py` | Auditorías clínicas especializadas |
| `centinela_reporte.py` | Reporte semanal centinela campaña invierno (PDF MINSAL) |
| `AUTO_SSASUR.py` | Descarga automatizada SSASur (recetas + stock) |
| `aa_colors.py` | Paleta de colores compartida (impresión económica) |
| `_generar_glosario.py` | Genera Glosario_Maestro_AA.pdf |

## Reglas de arquitectura

- **Nuevas homologaciones de nombres**: SOLO en `utils_aa.py → HOMOLOGACION_RAW`. Nunca duplicar en scripts individuales.
- **El modelo SGLI no tiene techo de capacidad**: el Nivel Objetivo T se calcula desde la demanda; Cap_Max es informativo y solo activa [ALERTA_ESTRES].
- **RUTs**: nunca a la API. `agente_duplicados.py` anonimiza con SHA-256 antes de llamar a Claude.

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
- Ajustar umbrales o fórmulas en `sgli.py` o `reposicion_dias_habiles.py`
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
