---
name: auto-ssasur
description: Diagnóstico y corrección rápida de AUTO_SSASUR.py / AUTO_SSASUR.bat — descarga automatizada de recetas, stock y GT desde SSASur. Usar ante: "AUTO_SSASUR no funciona", "no me abre", "no puedo ejecutarlo", "el browser no aparece", "se cierra solo", "error en descarga SSASur", "no reconoce el comando", "bat no encontrado", "falla la automatización", "TargetClosedError", "sesión expirada".
---

# Skill: auto-ssasur

Diagnostica y corrige el pipeline de descarga automatizada SSASur del proyecto
Maestro AA (Hospital de Pitrufquén). Trabaja sobre `AUTO_SSASUR.py` y
`AUTO_SSASUR.bat`.

## Errores frecuentes y solución rápida

### 1. El browser no abre / se cierra inmediatamente (`TargetClosedError`)
**Causa**: la sesión guardada (`.ssasur_session.json`) está expirada. Al
navegar al dashboard con cookies vencidas, SSASur redirige y Playwright cierra
la página antes de que el usuario pueda logarse.

**Diagnóstico**: buscar en `auto_ssasur_error.log`:
```
playwright._impl._errors.TargetClosedError: Page.wait_for_function: Target page, context or browser has been closed
```

**Solución inmediata**:
```powershell
Remove-Item "C:\Users\danie\Downloads\maestro\.ssasur_session.json" -ErrorAction SilentlyContinue
```
Luego volver a ejecutar `AUTO_SSASUR.bat`. El browser abrirá SSASur limpio
para que el usuario ingrese sus credenciales.

**Solución permanente** (ya aplicada en `AUTO_SSASUR.py`): el script detecta
`TargetClosedError` automáticamente, borra la sesión expirada y reabre el
browser sin sesión. Si el error vuelve a ocurrir verificar que el bloque
`try/except _TCE` esté en el paso 1 de `main()`.

---

### 2. "AUTO_SSASUR.bat no se reconoce como comando"
**Causa**: `py` (lanzador Windows) no está en PATH, o el `.bat` tiene ruta
relativa incorrecta.
**Solución**:
1. Verificar `py --version` en PowerShell desde la raíz del proyecto.
2. Revisar que la línea de ejecución del `.bat` sea `py AUTO_SSASUR.py %*`.
3. Si el error persiste, ejecutar directamente: `py AUTO_SSASUR.py`.

---

### 3. Error de credenciales / Google Drive (`credentials.json`)
**Causa**: `credentials.json` ausente o `token_drive.json` expirado.
**Solución**: ejecutar `SETUP_DRIVE.bat` para regenerar el token OAuth.
`credentials.json` debe estar en la raíz del proyecto (no en subcarpeta).

---

### 4. Error de deduplicación de recetas
**Causa**: `dedup_recetas.py` detecta recetas duplicadas entre archivos GT y
crea `.bak` del archivo original.
**Solución**: revisar los `.bak` generados, confirmar duplicados reales
(sobre-extracción GT) y volver a ejecutar.

---

### 5. SSASur no responde / timeout
**Causa**: mantenimiento del portal SSASur o VPN hospitalaria inactiva.
**Solución**: verificar conectividad con VPN, esperar 5 min y reintentar.
No modificar el script — el problema es de red, no de código.

---

## Pasos del diagnóstico

1. **Leer `auto_ssasur_error.log`** completo (no solo la primera línea).
2. **Identificar el tipo de error**:
   - `TargetClosedError` → Caso 1 (sesión expirada)
   - `NameError` / `ImportError` → bug en el script, localizar por número de línea
   - `TimeoutError` → Caso 5 (red/VPN)
   - `FileNotFoundError` credentials → Caso 3
3. **No tocar** `dedup_recetas.py` ni `utils_aa.py` sin confirmar con el usuario.
4. Confirmar la corrección revisando la última línea del log de salida.

## Contexto del pipeline

```
AUTO_SSASUR.bat
  └─ py AUTO_SSASUR.py
       ├─ [1] Login SSASur (Playwright/Chromium, headless=False)
       │        └─ .ssasur_session.json  ← si expiró: TargetClosedError
       ├─ [2] Descarga recetas → CSV (informe_completo_recetas*.csv)
       ├─ [3] Descarga GT → XLSX (04_Farmacia_Gestion_Territorial/)
       ├─ [4] Descarga stock → XLSX (reporte_de_stock_*.xlsx)
       ├─ [5] cruce_gt.py --generar  (planillas + letreros)
       ├─ [6] maestro_aa.py (Consolidado)
       ├─ [7] dedup_recetas.py --limpiar
       └─ [8] publicar_drive.py  ← puede fallar sin token_drive.json
```

## Archivos clave de diagnóstico

| Archivo | Qué muestra |
|---|---|
| `auto_ssasur_error.log` | Stderr completo del último run |
| `auto_ssasur_stdout.log` | Salida normal del último run |
| `.ssasur_session.json` | Sesión Playwright (borrar si expirada) |
| `token_drive.json` | Token OAuth Drive (borrar + re-auth si expirado) |

## Datos sensibles

Los CSV de recetas contienen RUT de pacientes (Ley 19.628). Nunca mostrar
su contenido en la respuesta, solo confirmar si existen o no.
