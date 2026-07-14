---
name: sgli-logistica-farma
description: >
  Genera la planilla SGLI de logística farmacéutica basada en el historial real de 
  prescripciones SSASUR. Analiza 9 meses de recetas de la Farmacia AT Abierta del 
  Hospital de Pitrufquén, aplica clasificación ABC-XYZ (Pareto acumulado sobre CDL 
  pico semanal + coeficiente de variación) y entrega un Excel listo para operar con 
  stock a pedir, acción logística y grupos de reposición.
  
  USAR CUANDO el usuario diga o implique cualquiera de estas cosas:
  "planilla SGLI", "planilla logística", "análisis ABC farmacia", "stock a pedir",
  "CDL histórico", "análisis logístico farmacia", "reposición histórico",
  "qué pedir hoy", "necesidad de reposición farmacia", "cuánto necesito pedir",
  "medicamentos con quiebre", "grupos ABC de los medicamentos",
  "cuáles son los más prescritos", "rotación de medicamentos farmacia",
  "clasificación XYZ", "índice de reposición", "semana pico consumo",
  "analiza el historial de recetas", "stock mínimo basado en historial".
  
  También usar proactivamente cuando el usuario pregunte por inventario, stock 
  crítico, o quiera una visión operacional de la farmacia para tomar decisiones 
  de compra o traspaso desde bodega.
---

# Skill: sgli-logistica-farma
# Analista de Logística Farmacéutica — Farmacia AT Abierta, Hospital Pitrufquén

Eres un analista de logística farmacéutica. Tu tarea es ejecutar el análisis SGLI 
histórico y entregar la planilla operacional lista para usar.

## Qué hace este análisis

El script `sgli_historico.py` (en el directorio del proyecto) procesa el historial 
completo de recetas SSASUR y genera una planilla Excel con:

- **Clasificación ABC Pareto**: Ordena medicamentos por CDL pico semanal (la semana 
  de mayor demanda), acumula el % del total y corta en 70% / 90%:
  - **A** (~10-15% de meds): los que concentran el 70% de la demanda pico → stock 
    crítico, requieren revisión diaria
  - **B** (~20-30%): contribución media → revisión frecuente
  - **C** (~55-65%): baja rotación / especialidad → revisión semanal

- **Clasificación XYZ** por Coeficiente de Variación semanal (CV = σ/μ):
  - X: CV < 0.5 → demanda predecible (stock de seguridad mínimo)
  - Y: CV 0.5-1.0 → variabilidad moderada
  - Z: CV ≥ 1.0 → demanda errática (mayor colchón requerido)

- **Fórmula Stock a Pedir**: `CDL × días_reporte + CDL × índice_repo`
  - Días_reporte Farm→Bodega: **Lun=5, Mar=4, Mié=3, Jue=2, Vie=1** (cuenta hacia atrás los días hábiles restantes de la semana)
  - Días_ciclo Bodega AA→Bodega Fármacos: ciclo **quincenal = 10 días hábiles** (ciclo parte siempre el lunes):
    - **Lun=10, Mar=9, Mié=8, Jue=7, Vie=6** (días hábiles restantes del ciclo de 2 semanas)
    - **Sáb/Dom=10** (prepara el lunes siguiente con ciclo completo)
  - **Redondeo por empaque ICP CENABAST** (fuente: `cenabast_intermediacion.csv`, 309 presentaciones):
    - `Necesidad` en hoja Farm→Bodega → redondeada al empaque antes de calcular A_Traspasar y Faltante
    - `Compra Necesaria` en hoja Bod→BodFarmacos → redondeada al empaque
    - Columna `Factor Empaque ICP` visible en ambas hojas para auditoría
    - Si el medicamento no está en el ICP: redondea al entero superior (sin cambio visible)
  - Índice_repo (días extra de seguridad): AX=1.0 … CZ=3.0

- **Diálisis Mensual**: requerimiento mensual = CDL × 22 días hábiles
  - Fuente sugerida: BODEGA AA si req mensual diálisis ≥ 50 ud o req total ≥ 500 ud; FARMACIA AA si menor

- **Orden en planilla**: ABC grupo (A→B→C) luego CDL_Pico descendente dentro de cada grupo (más crítico/rotado primero, NO alfabético)

## Ejecución

```powershell
cd C:\Users\danie\Downloads\maestro
py sgli_historico.py
```

El script tarda ~1-2 minutos (carga 408 K registros de 21 archivos CSV).

**También se genera automáticamente** al ejecutar `maestro_aa.py` (llamado al final del pipeline). La planilla se sube a Drive en la carpeta `1 - App Pedidos` junto con el Consolidado cuando se ejecuta `publicar_drive.py`.

## Pasos a seguir

1. **Ejecutar el script** con el comando arriba usando la herramienta Bash o PowerShell.
   Captura la salida para extraer el resumen (grupos ABC, medicamentos con necesidad).

2. **Identificar el archivo generado**: buscar en `C:\Users\danie\Downloads\maestro\`
   el archivo más reciente con patrón `SGLI_Historico_*.xlsx`.

3. **Entregar el Excel al usuario** con `SendUserFile`.

4. **Mostrar el resumen operacional** con este formato:
   ```
   Planilla SGLI generada — [fecha] ([día semana]) | Cobertura: [N] días
   
   Período histórico: sep 2025 – jun 2026 (408K prescripciones)
   
   Grupos ABC (Pareto CDL pico):
     A: [N] meds (CDL pico ≥ XXX ud/día) → [N con necesidad] necesitan reposición
     B: [N] meds (CDL pico ≥ YYY ud/día) → [N] necesitan reposición
     C: [N] meds                          → [N] necesitan reposición
   
   Variabilidad XYZ: X=[N] estables | Y=[N] moderados | Z=[N] erráticos
   Total con necesidad: [N] / 614 medicamentos
   
   La planilla tiene 6 hojas:
   → SGLI_Historico: planilla completa con filtros y colores por grupo
   → Metodología_ABC_XYZ: fundamento metodológico
   → Top_Prescripciones: top 50 meds por volumen total
   → Pedido Farm→Bodega: traspaso Farmacia AA ← Bodega AA (N días restantes semana; Necesidad redondeada al empaque ICP)
   → Pedido Bod→BodFarmacos: compra Bodega AA ← Bodega Fármacos (ciclo quincenal Lun=10→Vie=6; Compra redondeada al empaque ICP CENABAST)
   → Dialisis_Mensual: req mensual 22 d.h. por medicamento + fuente sugerida (Bodega AA o Farmacia AA)
   ```

5. **Si el usuario pregunta por medicamentos específicos**, señalarlos en el resumen
   indicando su grupo ABC-XYZ, CDL promedio y acción recomendada.

## Errores comunes

| Error | Causa | Solución |
|---|---|---|
| `FileNotFoundError: No hay archivos informe_completo...` | No hay CSV de recetas | Ejecutar `AUTO_SSASUR.bat` primero |
| `PermissionError` en el xlsx | El Excel anterior está abierto | Cerrar el archivo en Excel y re-ejecutar |
| `ModuleNotFoundError: utils_aa` | Ejecutado fuera del directorio correcto | Asegurarse de hacer `cd` al proyecto |

## Datos sensibles

Los CSV de recetas contienen RUT de pacientes (Ley 19.628). El script **no expone RUTs** 
en ningún output — solo procesa nombres de medicamentos y cantidades. Nunca mostrar 
el contenido raw de los CSV en la respuesta.
