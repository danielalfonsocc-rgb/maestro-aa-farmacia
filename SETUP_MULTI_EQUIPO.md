# Usar Maestro AA en otros computadores de la farmacia

## 0. Si solo necesitas VER el estado actual — no instales nada

La planilla de Gestión Territorial ya quedó publicada como versión
**definitiva** en Google Sheets (19-07-2026). Cualquiera con el link la ve
en tiempo real, actualizada, con colores y todo el formato — **sin instalar
Python, Claude Code, ni clonar nada**:

**https://docs.google.com/spreadsheets/d/1ydzzscntphNzjjQq9pnt6ZXffneMHjwjjuHzaHGkihQ**

⚠️ Ojo: hoy es un mirror de una sola vía. `gt_maestro.py` sigue escribiendo
al `.xlsx` local, y `py publicar_gt_sheets.py` empuja ese estado al link de
arriba — pero nadie puede *editar* la planilla directamente desde el link y
que eso vuelva al Excel. Si alguien edita el Sheet a mano, la próxima
republicación se lo pisa. Para editar de verdad (registrar recetas, cambiar
Estado, etc.) todavía hace falta el setup completo de abajo, desde el equipo
que tiene el `.xlsx`.

Lo mismo aplica a la carpeta de Drive `Farmacia AA/2 - Gestión Territorial/`
(Revisión de Solicitudes / Nóminas de Envío por establecimiento) — cualquiera
con acceso a esa carpeta compartida la ve, sin instalar nada.

## 1. Para EDITAR (correr gt_maestro.py, procesar solicitudes, etc.)

Esta sección cubre qué se necesita instalar/configurar en **cada computador
nuevo** para poder usar Claude Code + este repo tal como se usa hoy en el
equipo principal.

### 1.1 Lo que hay que instalar una vez por equipo

1. **Claude Code** (la app/CLI) — instalar y hacer login con la cuenta de
   Anthropic que corresponda usar en ese equipo.
2. **Python 3.10+** — vía el lanzador `py` (`python.org`, marcar "Add to PATH").
3. **Git** — para clonar/actualizar el repo.
4. **Google Chrome** — lo usa el navegador integrado de Claude Code para
   SSASUR y Google Sheets.

### 1.2 Clonar el repo y las dependencias

```bat
git clone https://github.com/danielalfonsocc-rgb/maestro-aa-farmacia.git maestro
cd maestro
py -m pip install -r requirements.txt
py -m playwright install chromium
```

El repo es público (decisión consciente, ver memoria del proyecto) — no hace
falta usuario/token de GitHub para clonarlo. `git pull` dentro de esa carpeta
trae las actualizaciones de código que se hagan desde cualquier otro equipo,
**incluido `_gt_sheets_id.json`** — así que un `git pull` en un equipo nuevo
ya deja apuntando al mismo Sheet definitivo sin configurar nada más para eso.

### 1.3 Rutas que son específicas de cada máquina

`utils_aa.py` lee estas rutas desde variables de entorno, con un valor por
defecto que es el de la máquina original. **En un equipo nuevo hay que
sobrescribirlas** (Panel de control → Sistema → Variables de entorno, o
`setx` en una consola):

| Variable de entorno | Para qué sirve |
|---|---|
| `MAESTRO_GT_MAESTRO_XLSX` | Ruta a `GT PITRUFQUEN 2026 (2).xlsx` (el maestro de Gestión Territorial) |
| `MAESTRO_RCH_DIR` | Carpeta del formulario ISP de Recetas Cheque |
| `MAESTRO_PLANTILLA_RCH` | Plantilla en blanco del formulario ISP |

```bat
setx MAESTRO_GT_MAESTRO_XLSX "D:\Farmacia\GT PITRUFQUEN 2026 (2).xlsx"
```

(cerrar y volver a abrir la consola/Claude Code para que tome el valor nuevo)

### 1.4 El archivo maestro sigue siendo un `.xlsx` local — ese es el límite real

`GT PITRUFQUEN 2026 (2).xlsx` vive en `Downloads` de un solo computador, y
`gt_maestro.py` **escribe ahí**, no en Sheets. El Sheet definitivo (sección 0)
es un espejo de solo lectura que se actualiza corriendo
`py publicar_gt_sheets.py` — no al revés.

Esto significa: si dos personas quieren *editar* (registrar recetas, correr
el workflow de una solicitud) desde dos equipos distintos el mismo día, van a
pisarse — solo un equipo a la vez debería tener la copia "buena" del `.xlsx`
abierta/editándose. En la práctica, hoy eso significa que el trabajo de
registro sigue centralizado en el equipo que tiene el archivo.

Dos caminos si esto se vuelve un problema real:

**Opción A — Carpeta compartida sincronizada (rápido, con riesgo).**
Poner el `.xlsx` en OneDrive/Google Drive de escritorio y apuntar
`MAESTRO_GT_MAESTRO_XLSX` a esa ruta sincronizada en cada equipo. Si dos
personas lo editan a la vez en dos equipos, el sync puede crear un archivo
duplicado ("conflicto de copia") que hay que fusionar a mano.

**Opción B — Reescribir `gt_maestro.py` para leer/escribir directo en Sheets
(pendiente, no iniciado).**
Es la solución de fondo: cualquier equipo editaría la misma hoja en la nube
sin depender de quién tiene el `.xlsx` "bueno". Es un trabajo de ingeniería
bastante más grande que publicar (hay que reescribir `upsert_receta`,
`buscar_receta_en_maestro`, etc. contra la API de Sheets en vez de openpyxl)
— avísame si quieres que lo aborde.

### 1.5 Credenciales — nunca van en el repo

- **`credentials.json`** y **`token_drive.json`** (acceso a Google
  Drive/Sheets): están en `.gitignore` a propósito — nunca se suben a GitHub.
  Para un equipo nuevo, cópialos tú mismo por un canal seguro (pendrive,
  Drive privado, etc.), o corre el flujo de autorización de nuevo en ese
  equipo (`publicar_drive.py` con `--setup` la primera vez) para generar un
  `token_drive.json` propio de esa máquina.
- **SSASUR**: el login es siempre manual, en el navegador, en cada sesión —
  esto es igual en todos los equipos y no requiere configuración.
- **Historial de recetas (`informe_completo_recetas*.csv`)**: no se
  sincronizan entre equipos (están en `.gitignore`, contienen RUT de
  pacientes). Cada equipo que necesite `gt_maestro.py` clasificando
  Refrigerado/Controlado debe correr `AUTO_SSASUR.bat` localmente para tener
  su propia copia actualizada.

### 1.6 Resumen del primer arranque en un equipo nuevo

```bat
git clone https://github.com/danielalfonsocc-rgb/maestro-aa-farmacia.git maestro
cd maestro
py -m pip install -r requirements.txt
py -m playwright install chromium
setx MAESTRO_GT_MAESTRO_XLSX "<ruta al xlsx en ese equipo, o a la carpeta sincronizada>"
:: copiar credentials.json / token_drive.json si vas a usar Drive/Sheets desde ahí
```

Después de eso, `ABRIR_APP.bat` y los scripts (`gt_maestro.py`,
`publicar_gt_sheets.py`, `publicar_drive.py`, etc.) funcionan igual que en el
equipo original.
