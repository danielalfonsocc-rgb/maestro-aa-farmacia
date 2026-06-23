# Proyecto: Maestro AA — Consolidado Operacional Farmacia AT Abierta

Herramienta de gestión farmacéutica para Hospital de Pitrufquén (SSASur).
Universo: **378 medicamentos AA**. Fuente de datos: SSASur (stock + recetas).

## Archivos principales

| Archivo | Rol |
|---|---|
| `maestro_aa.py` | Consolidación principal → `Consolidado_AA_MAESTRO.xlsx` (14 hojas) |
| `app_pedidos.py` | Dashboard Streamlit — pedidos, faltantes, alertas |
| `sgli.py` | Cálculo SGLI (stock/licitación/días hábiles) |
| `reposicion_dias_habiles.py` | Reposición ajustada por feriados |
| `cruce_gt.py` | Cruce con Guías de Tratamiento |
| `recetas_cheque.py` / `recetas_duplicadas.py` | Auditorías de recetas |
| `auditoria_prescripcion.py` / `auditoria_empagliflozina.py` | Auditorías clínicas |
| `AUTO_SSASUR.py` / `descargar_ssasur.py` | Descarga automatizada SSASur |
| `evaluar_compra_agil.py` | Evaluación de compras ágiles |
| `aa_colors.py` | Paleta de colores compartida |

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
Esto evita que Claude explore el proyecto entero antes de responder.

### 2. Limitar el output explícitamente
Para Opus, agregar siempre al final:
```
Respuesta concisa. Solo el código modificado + 1 línea de explicación.
```

### 3. Preguntas de datos: usar `--json` mode del auditor
Para análisis de CSV/Excel puntuales, usar el script directo en lugar
de pegar datos en el chat.

### 4. No pegar DataFrames enteros
En lugar de pegar 50 filas de un DataFrame:
```
# Bien
df.dtypes, df.shape, df.head(3) — pegar solo esto.
```

### 5. Reutilizar contexto de sesión
Dentro de una sesión larga, no re-explicar el proyecto cada mensaje.
El contexto ya está cacheado (95% cache hit en este proyecto).

## Privacidad y datos sensibles

Los archivos `informe_completo_recetas*.csv` y `reporte_de_stock_*.xlsx`
pueden contener **RUT de pacientes** sujetos a la **Ley 19.628**.

- **NUNCA** pegar filas con RUT en el chat.
- Para debugging con datos reales: anonimizar primero (reemplazar RUT por `XXXXX`).
- El directorio `maestro/` no debe versionarse en GitHub público.

## Stack técnico

- Python 3.10 vía lanzador `py`
- Streamlit ≥ 1.35 en puerto **8501** (`py -m streamlit run app_pedidos.py`)
- pandas, numpy, openpyxl, reportlab, rapidfuzz
- anthropic (Claude API directa dentro de la app)
- Playwright / Selenium para descarga SSASur

## Iniciar el servidor

```bat
py -m streamlit run app_pedidos.py --server.headless true --browser.gatherUsageStats false --server.port 8501
```

O simplemente: `ABRIR_APP.bat`
