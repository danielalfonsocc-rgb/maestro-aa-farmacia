---
name: revisar-bugs
description: >
  Auditoría sistemática de bugs en el proyecto Maestro AA. Usa subagentes paralelos
  para analizar todos los scripts en simultáneo contra las reglas de arquitectura de
  CLAUDE.md, y aplica los fixes confirmados. Tiempo estimado: 3-5 min en lugar de
  un contexto completo.

  USAR CUANDO el usuario diga o implique:
  "busca bugs", "revisa el código", "reanaliza el programa", "hay errores?",
  "chequea el proyecto", "qué está mal", "audit", "revisar bugs",
  "encontrar problemas", "corrije los bugs", "qué bugs hay".

  También usar proactivamente después de añadir un script nuevo o refactorizar
  más de 2 archivos en la misma sesión.
---

# Skill: revisar-bugs
# Auditoría paralela de bugs — Proyecto Maestro AA

Ejecuta una auditoría completa del proyecto usando subagentes paralelos (un agente
por archivo o grupo de archivos), reporta los bugs por severidad y aplica los fixes
confirmados. No re-explica el proyecto al usuario — trabaja directamente.

## Checklist de bugs a buscar (aplicar en CADA archivo)

### 🔴 CRÍTICO — causaría crash o datos incorrectos silenciosos

1. **Empty-before-concat**: `pd.concat([])` sin verificar que la lista no esté vacía
   → `ValueError: No objects to concatenate`. Buscar patrón: `chunks = []` + loop + `pd.concat(chunks)` sin `if not chunks:` previo.

2. **Empaque inconsistente**: `Necesidad_Farm` o `Necesidad_Bod` calculados con `.round(0)` en lugar de `redondear_empaque()`. En la sección principal se usa `redondear_empaque`; en pipelines secundarios (_pipeline_ped_dial, hojas de diálisis en pedido_fusion.py) puede faltar.

3. **KeyError silenciado**: uso de `df.get('col', default)` donde `default` es escalar (ej. `0`) y el resultado se usa como máscara booleana. Si la columna no existe, `df[0 > 0]` tira `TypeError`.

4. **Día-de-semana no chequeado**: reportes/procesos semanales (centinela_reporte.py) invocados sin verificar `date.today().weekday() == 0` (lunes).

5. **Self-mapping en HOMOLOGACION_RAW**: entradas donde `norm_erp(key) == norm_erp(value)` (ej. doble espacio en ambos lados → colapsan igual). Dead code que enmascara la ausencia real de homologación.

### 🟡 MODERADO — output incorrecto pero no crash

6. **Computación duplicada**: misma columna asignada dos veces con lógica idéntica (ej. `Prof_Norm` en secciones 7 y 15 de maestro_aa.py). La segunda sobreescribe sin añadir valor.

7. **Etiquetas de paso inconsistentes**: en AUTO_SSASUR.py los sub-pasos `[Nb/TOTAL]` deben coincidir con el total real de pasos. Si se añadieron pasos nuevos sin actualizar los sub-pasos, el conteo queda mal.

8. **Cobertura de días incorrecta**: en `_dias_ef()` y `_dias_ciclo()` de pedido_fusion.py verificar que el conteo incluya hoy y excluya fines de semana/feriados correctamente.

9. **`fillna(0)` en columna de fecha**: `pd.to_datetime(...).fillna(0)` en lugar de `pd.NaT`. El 0 se convierte en epoch (1970-01-01) y distorsiona cálculos de semana/periodo.

### 🔵 MENOR — código muerto o confuso

10. **Entradas redundantes en HOMOLOGACION_RAW**: después de la limpieza de self-maps, verificar que no queden entradas apuntando al mismo destino canónico por rutas distintas.

11. **Comentarios con número de paso obsoleto**: bloques `# ── PASO X` en AUTO_SSASUR.py fuera de sincronía con la numeración real.

12. **`import` no utilizado**: en scripts nuevos o refactorizados.

---

## Cómo ejecutar este skill

### Paso 1 — Inventariar los archivos a auditar

```python
# Archivos Python del proyecto (excluir __pycache__, .venv, tests)
ARCHIVOS_PRINCIPALES = [
    "maestro_aa.py",           # 2000+ líneas — dividir en grupos de secciones
    "utils_aa.py",
    "AUTO_SSASUR.py",
    "app_pedidos.py",
    "sgli.py",
    "reposicion_dias_habiles.py",
    "pedido_fusion.py",
    "agente_duplicados.py",
    "auditoria_duplicados_profunda.py",
    "auditoria_prescripcion.py",
    "centinela_reporte.py",
    "cruce_gt.py",
    "publicar_escritorio.py",
    "publicar_drive.py",
    "dedup_recetas.py",
    "sgli_historico.py",
    "skill_gt/scripts/generar.py",
]
```

### Paso 2 — Lanzar Workflow con agentes paralelos

Usar la herramienta `Workflow` con un script que distribuya los archivos entre
agentes paralelos. Cada agente recibe:
- La ruta del archivo
- El checklist completo (items 1-12 del apartado anterior)
- El contenido de CLAUDE.md (reglas de arquitectura)
- Instrucción: devolver JSON estructurado `{bugs: [{linea, severidad, descripcion, fix_sugerido}]}`

Esquema del script de Workflow:
```javascript
export const meta = {
  name: 'audit-maestro-aa',
  description: 'Auditoría paralela de bugs Maestro AA',
  phases: [
    { title: 'Análisis', detail: 'Leer y auditar cada archivo' },
    { title: 'Consolidar', detail: 'Agrupar por severidad y archivo' },
    { title: 'Fixes', detail: 'Aplicar correcciones confirmadas' },
  ],
}
// 1) Leer CLAUDE.md para pasar el contexto de arquitectura a cada agente
// 2) pipeline(ARCHIVOS_PRINCIPALES, archivo => agent(prompt_auditoria(archivo), {schema: BUGS_SCHEMA}))
// 3) Aplanar resultados, ordenar por severidad, devolver lista consolidada
```

### Paso 3 — Reportar y aplicar fixes

Formato de reporte obligatorio:

```
## Bugs encontrados — Maestro AA

### 🔴 Críticos (N)
| Archivo | Línea | Bug | Fix |
|---|---|---|---|
| maestro_aa.py | 393 | pd.concat sin check vacío | Añadir `if not chunks: raise FileNotFoundError(...)` |

### 🟡 Moderados (N)
...

### 🔵 Menores (N)
...

Total: N bugs | Aplicados: N | Pendientes revisión: N
```

Aplicar automáticamente los bugs **críticos y moderados** con `Edit`.
Presentar los **menores** al usuario para que decida.

---

## Reglas de arquitectura que NUNCA se pueden violar (de CLAUDE.md)

- Nuevas homologaciones: **SOLO** en `utils_aa.py → HOMOLOGACION_RAW`
- SGLI no tiene techo de capacidad (Cap_Max es informativo)
- RUTs: **nunca** a la API ni al chat
- `agente_duplicados.py` usa SHA-256 antes de llamar a Claude
- GT raw downloads van a `../04_Farmacia_Gestion_Territorial/` (carpeta hermana)
- Drive: **no subir** CSV sábanas ni stock xlsx

## Historial de bugs conocidos (ya corregidos — no re-reportar)

| Fecha | Archivo | Bug | Estado |
|---|---|---|---|
| 2026-06-29 | maestro_aa.py | pd.concat sin check empty csv_files | ✅ Corregido |
| 2026-06-29 | utils_aa.py | TRAZODONA self-map en HOMOLOGACION_RAW | ✅ Corregido |
| 2026-06-29 | maestro_aa.py | Prof_Norm computado dos veces (secciones 7 y 15) | ✅ Corregido |
| 2026-06-29 | maestro_aa.py | _pipeline_ped_dial sin redondear_empaque en Necesidad_Farm/Bod | ✅ Corregido |
| 2026-06-29 | AUTO_SSASUR.py | [5b/7][5c/7][5d/7] debían ser /9 | ✅ Corregido |
| 2026-06-29 | AUTO_SSASUR.py | Centinela corría todos los días (debe ser solo lunes) | ✅ Corregido |
| 2026-06-29 | AUTO_SSASUR.py | Comentario "PASO 5e" duplicado con Pedido Fusionado | ✅ Corregido |
| 2026-06-29 | pedido_fusion.py | apfarm/apbod diálisis sin redondeo de empaque | ✅ Corregido |
| 2026-06-29 | maestro_aa.py | df_ped_dial.get('Necesidad_Farm/Bod',0)>0 → TypeError (KeyError silenciado) | ✅ Corregido |
| 2026-06-29 | maestro_aa.py | LEYENDA_CRIT incompleta: faltaban niveles Bodega (4-MODERADO, 5-BAJO, 3-ALTO, 3-MODERADO) | ✅ Corregido |
| 2026-06-29 | cruce_gt.py | c_rec fallback a 0 leía columna índice 0 si no había columna de receta | ✅ Corregido |
| 2026-06-29 | cruce_gt.py | RUT de paciente serializado en gt_enriquecido.json (violación Ley 19.628) | ✅ Corregido |
| 2026-06-29 | sgli_historico.py | merged.get('Stock_Farm', 0) en DataFrame → AttributeError al llamar .fillna() | ✅ Corregido |
| 2026-06-30 | agente_duplicados.py | import VERDE no usado | ✅ Corregido |
| 2026-06-30 | auditoria_duplicados_profunda.py | imports NARANJA/AMBAR/VERDE no usados + anon_map muerto en _anon_run | ✅ Corregido |
| 2026-06-30 | auditoria_prescripcion.py | import sys no usado | ✅ Corregido |
| 2026-06-30 | publicar_drive.py | imports base64/io no usados | ✅ Corregido |
| 2026-06-30 | sgli.py | import re no usado | ✅ Corregido |
| 2026-06-30 | cruce_gt.py | (falso positivo descartado) csv SÍ se usa en csv.reader línea 155 | ⛔ No es bug |
| 2026-06-30 | cruce_gt.py | (falso positivo descartado) variable 'g' en sorted() de escribir_excel() no genera shadowing real | ⛔ No es bug |
| 2026-06-30 | utils_aa.py | (falso positivo descartado) lambda c,_c=cols es patrón intencional anti late-binding | ⛔ No es bug |

## Datos sensibles

Los CSV de recetas y el reporte de stock pueden contener RUTs (Ley 19.628).
**Nunca mostrar filas con RUT en la respuesta.** Si hay un bug relacionado con
un CSV, trabajar solo con dtypes, shape y nombres de columnas.
