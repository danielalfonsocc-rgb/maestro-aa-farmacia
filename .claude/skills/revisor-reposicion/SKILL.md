---
name: revisor-reposicion
description: Analiza el estado de reposición del Almacén de Abastecimiento (AA) del Hospital de Pitrufquén. Lee Consolidado_AA_MAESTRO.xlsx, detecta medicamentos bajo stock mínimo, calcula días hábiles de cobertura y genera el resumen de pedido. Usar ante "qué reponer hoy", "revisar stock", "generar pedido AA", "cuántos días de stock".
---

# Skill: revisor-reposicion

Analiza el stock actual del AA y produce la lista de reposición del día. Trabaja
siempre sobre el archivo más reciente de `Consolidado_AA_MAESTRO*.xlsx` o el
indicado por el usuario.

## Contexto del proyecto

- Hospital de Pitrufquén, SSASUR.
- 378 medicamentos activos en el AA.
- Archivos de salida diarios: `Reposicion_DiasHabiles_AA_YYYYMMDD.xlsx`,
  `Resumen_Pedidos_AA_YYYYMMDD_HHMM.xlsx`.
- Script de pedidos: `app_pedidos.py` (Streamlit).
- Automatización SSASur: `AUTO_SSASUR.py`.

## Pasos del agente

1. **Identificar archivo fuente**: buscar `Consolidado_AA_MAESTRO.xlsx` en el
   directorio del proyecto. Si hay versiones con fecha (`_YYYYMMDD_HHMM.xlsx`),
   usar la más reciente. Confirmar al usuario cuál se usará.

2. **Leer stock actual**: columnas clave a buscar (nombres aproximados, confirmar
   con el usuario si difieren):
   - Nombre / Descripción del medicamento
   - Stock actual (unidades)
   - Stock mínimo / Punto de pedido
   - Consumo promedio diario o mensual
   - Unidad de medida

3. **Calcular días de cobertura**:
   - `dias_cobertura = stock_actual / consumo_diario_promedio`
   - Si no hay consumo_diario disponible: `consumo_diario = consumo_mensual / 22`
     (días hábiles promedio por mes).

4. **Detectar alertas**:
   - 🔴 Crítico: `stock_actual <= stock_minimo` → pedir inmediato.
   - 🟡 Alerta: `dias_cobertura < 7` → pedir en los próximos 2 días hábiles.
   - 🟢 Normal: `dias_cobertura >= 7` → sin acción.

5. **Generar resumen**: tabla ordenada por criticidad con columnas:
   - Medicamento | Stock actual | Mínimo | Días cobertura | Estado | Cantidad a pedir

6. **Proponer nombre de archivo de salida**:
   `Reposicion_DiasHabiles_AA_YYYYMMDD.xlsx` con la fecha del día.

## Restricciones

- No modificar `Consolidado_AA_MAESTRO.xlsx` (solo lectura).
- No incluir datos de pacientes en la salida (Ley 19.628).
- Si un medicamento no tiene consumo registrado, marcarlo como "sin datos" y
  no calcular días de cobertura (no inventar valores).
- Las cantidades a pedir son estimaciones: el QF valida antes de enviar el pedido.

## Archivos relacionados

- `app_pedidos.py` — interfaz Streamlit para gestión de pedidos.
- `AUTO_SSASUR.py` — automatización de reportes para SSASur.
- `auditoria_prescripcion.py` — auditoría de prescripciones (flujo separado).
- `cenabast_intermediacion.csv` — datos de intermediación CENABAST.
