---
name: revision-solicitudes-gt-pitrufquen
description: Procesa una SOLICITUD DE ENVÍO de Gestión Territorial — cuando un establecimiento destino de la red (CESFAM Freire, CESFAM Hualpín, CESFAM Quepe, CESFAM Teodoro Schmidt, DSM/Hospital Gorbea, DSM/Hospital Loncoche, DSM/Hospital Toltén, PSR Comuy, PSR Queule, PSR Los Galpones) pide que se le envíen las recetas vigentes de sus pacientes crónicos, prescritas en Hospital Pitrufquén. Dispara ante: "llegó una solicitud de [establecimiento]", "solicitud de envío de Toltén/Loncoche/Freire/Teodoro Schmidt/etc", "nómina que me envió [CESFAM/DSM/Hospital]", "procesa esta lista de gestión territorial", "revisar solicitud GT", "busca estas recetas por SSASUR", o cuando el usuario adjunte un xlsx (columnas tipo RUT/NOMBRE, a veces con Nº de Receta sugerido) o PDFs sueltos (Receta_<n>_..., Historico_<n>_...) que sean una solicitud de un establecimiento de la red. Usar SIEMPRE que se mencione una solicitud de gestión territorial, aunque no se nombren los scripts por su nombre.
---

# Skill: revision-solicitudes-gt-pitrufquen

Orquesta el pipeline ya existente para procesar una solicitud de envío de
Gestión Territorial (GT): un establecimiento destino (CESFAM, DSM, hospital
o PSR de la red Araucanía Sur) pide las recetas **vigentes** de sus pacientes
crónicos que están **prescritas en Hospital Pitrufquén**, para retirarlas vía
GT en vez de que el paciente viaje hasta Pitrufquén.

Esta skill orquesta `revision_solicitudes.py` y `descargar_recetas_pdf.py`, ya
probados en producción (Freire 23-07-2026, Toltén 25-07-2026, Loncoche,
CESFAM Teodoro Schmidt 31-07-2026), más `scripts/parsear_solicitud.py`
(propio de esta skill) para el caso más común de entrada: texto suelto
pegado desde un correo.

## Regla de negocio central (no cambiar sin confirmar con el usuario)

Toda receta que se envíe **siempre** debe estar digitada en Hospital
Pitrufquén (`Establecimiento Digita` contiene "PITRUFQUEN"). Si la única
receta vigente de un paciente está digitada en otro establecimiento, **no se
usa** — se marca para revisión manual. El origen equivocado invalida incluso
la receta más reciente.

"Vigente" = Estado **no** está en `{ENTREGADA, CERRADA / INCOMPLETA, ANULADA}`
(incluye `SOLICITADA` y `PENDIENTE`). Si la receta que sugiere el
establecimiento ya salió `ENTREGADA`, o no existe en el sistema, hay que
reemplazarla por la vigente del mismo RUT — nunca despachar una receta ya
entregada.

## Paso 0 — identificar el establecimiento y guardar lo recibido

Mapea el nombre que use el usuario al establecimiento oficial (`_CARPETA_LOCAL`
en `agregar_gt_manual.py`):

| Dice el usuario | `--estab` | Carpeta local |
|---|---|---|
| Freire | `CESFAM FREIRE` | `CESFAM_FREIRE` |
| Hualpín | `CESFAM HUALPIN` | `CESFAM_HUALPIN` |
| Quepe | `CESFAM QUEPE` | `CESFAM_QUEPE` |
| Teodoro / Teodoro Schmidt | `CESFAM TEODORO SCHMIDT` | `CESFAM_TEODORO_SCHMIDT` |
| Gorbea (DSM) | `GORBEA DSM` | `DSM_GORBEA` |
| Loncoche (DSM) | `LONCOCHE DSM` | `DSM_LONCOCHE` |
| Toltén (DSM) | `TOLTEN DSM` | `DSM_TOLTEN` |
| Hospital Gorbea | `GORBEA HOSP` | `HOSPITAL_GORBEA` |
| Hospital Loncoche | `LONCOCHE HOSP` | `HOSPITAL_LONCOCHE` |
| Hospital Toltén | `TOLTEN HOSP` | `HOSPITAL_TOLTEN` |
| PSR Comuy | `PSR COMUY` | `PSR_COMUY` |
| PSR Queule | `PSR QUEULE` | `PSR_QUEULE` |
| PSR Los Galpones | `PSR LOS GALPONES` | `PSR_LOS_GALPONES` |

Si el usuario no da un nombre exacto, pregunta o infiere del nombre del
archivo — pero confirma antes de correr nada si hay ambigüedad real (ej.
"Loncoche" solo, sin decir DSM u Hospital).

Si la solicitud llegó como **PDFs sueltos**, deben quedar en
`04_Farmacia_Gestion_Territorial/<CARPETA_LOCAL>/Revisión de Solicitudes/`
antes del paso 1 (nombres `Receta_<n>_....pdf` / `Historico_<n>_....pdf`). Si
llegó como **planilla xlsx** (columnas tipo RUT/NOMBRE, a veces con Nº de
Receta ya sugerido por el establecimiento), no hace falta moverla — se pasa
la ruta directa con `--xlsx`.

### Caso más común: texto suelto pegado (correo, PDF de Gmail, foto transcrita)

La mayoría de las solicitudes NO llegan en xlsx — llegan como una lista
suelta en el cuerpo del correo, una línea por paciente, tipo
`nombre RUT establecimiento especialidad [nota opcional]`, muchas veces con
mayúsculas/minúsculas y formato inconsistente entre un establecimiento y
otro. **No transcribas esa lista a mano** escribiendo cada fila en un script
— es trabajo mecánico puro que un script determinístico hace mejor y sin
gastar tokens de salida en retipear ~30 filas de RUT/nombre:

1. Pega el bloque de texto tal cual (sin reformatear) en un `.txt` con Write.
2. Corre:
   ```bash
   py .claude/skills/revision-solicitudes-gt-pitrufquen/scripts/parsear_solicitud.py \
     --texto <archivo.txt> --estab "<alias del establecimiento, ej. tolten>" \
     --out Solicitud_<CARPETA_LOCAL>_<fecha>.xlsx
   ```
   El script ancla cada línea en el patrón de RUT (muy confiable), separa
   nombre/resto, saca el alias del establecimiento y un N° de receta
   sugerido si aparece (`N°123456`, `Receta completa N°123456`), y fusiona
   RUT repetidos (pasa seguido que el establecimiento manda la misma fila
   dos veces, o la repite agregando el N° de receta que le faltaba). No
   intenta separar perfectamente "especialidad" de "nota" — deja el resto
   de la línea junto en la columna ESPECIALIDAD, porque el cruce en
   `_elegir_receta()` compara por substring en ambos sentidos y matchea
   igual aunque venga texto de más pegado.
3. El script imprime cuántas líneas NO calzaron con el patrón de RUT (si
   las hay) — revisa esas puntuales a mano, no todo el listado. Si el motivo
   es que el establecimiento digitó el RUT sin el guion del dígito
   verificador (ej. "83444649"), el aviso ya trae calculado y verificado el
   RUT sugerido con guion (ej. `[posible RUT sin guion, DV verificado:
   8344464-9]`) — confírmalo contra el nombre del paciente y agrégalo a
   mano a la fila correspondiente del xlsx antes del Paso 1, no hace falta
   calcular el dígito verificador vos mismo.
4. Sigue con el Paso 1 normal usando el xlsx recién generado (`--xlsx`).

Si la solicitud SÍ llega ya en xlsx (algunos establecimientos ya mandan
planilla), sáltate este paso y usa `--xlsx` directo como dice arriba.

## Paso 1 — cruce contra el histórico local (rápido, sin tocar SSASUR)

```bash
py revision_solicitudes.py --estab "<ESTABLECIMIENTO>" --xlsx "<archivo.xlsx>" --dry-run
```
(sin `--xlsx` si son PDFs sueltos ya dejados en la carpeta).

Lee los `[AVISO]` en consola — indican pacientes cuya receta sugerida no era
válida y fue reemplazada, o casos sin ninguna receta vigente en Pitrufquén
(esos quedan pendientes de revisión en vivo, paso 2). Si el resultado se ve
razonable, repite **sin** `--dry-run`: genera
`Feedback_Solicitud_<ESTAB>_<fecha>.xlsx` y archiva el original recibido en
`Revisión de Solicitudes/<MES AÑO>/<fecha>/`.

Este paso también corre `gt_maestro.detectar_alertas_mismo_rut()` sobre todas
las recetas vigentes de cada RUT — si un paciente tiene 2+ recetas vigentes
que parecen duplicadas o ambiguas, queda anotado en la columna Observación
para que el QF lo revise antes de despachar.

## Paso 2 — verificación en vivo en SSASUR (para lo que el CSV local no resuelve)

El CSV local (`informe_completo_recetas*.csv`, actualizado por AUTO_SSASUR)
no siempre tiene la cuota más nueva de un paciente todavía — es el caso más
común de "[AVISO] ... se adjunta la más reciente disponible (ENTREGADA)".
Para esos casos, busca en vivo:

```bash
py descargar_recetas_pdf.py --estab "<ESTABLECIMIENTO>" --rut-live-desde-feedback "<ruta al Feedback_Solicitud...xlsx>"
```

Qué hace: abre un browser real (no headless) y pide login manual de SSASUR
(~5 min — el SSO no es automatizable, avísale al usuario que debe loguearse
cuando se abra la ventana). Reutiliza la sesión guardada si existe. Entra al
módulo Receta → Consultar Receta y, por cada RUT de la planilla, busca las
recetas vigentes de origen "PITRUFQUEN HOSP.", descarga el PDF oficial
directo de cada una (endpoint `/receta/impresion/pdf/<n>/undefined` — no usa
el botón "Imprimir" del sitio porque abre el diálogo nativo del SO), arma
`Recetas_Combinadas_<CARPETA_LOCAL>_<fecha>.pdf`, mueve los PDF individuales
a "PDFs individuales", y escribe
`Resumen_Busqueda_SSASUR_<CARPETA_LOCAL>_<fecha>.xlsx` con lo encontrado por
paciente (para el QF).

Correr esto en segundo plano (`run_in_background`) mientras el usuario se
loguea — no es instantáneo, tarda varios minutos según cuántos RUT haya.

**Variantes** del mismo script para casos más simples:
- `--feedback "<Feedback...xlsx>"` — si las recetas del Feedback ya son
  confiables (no hace falta buscar por RUT), descarga directo por Nº de
  receta.
- `--rut "<RUT>"` — un solo paciente puntual.
- `--recetas "n1,n2,..."` — lista de números de receta sueltos.

## Paso 3 — entregar

Todo queda en `04_Farmacia_Gestion_Territorial/<CARPETA_LOCAL>/Revisión de
Solicitudes/<MES AÑO>/<fecha>/`: el Feedback, el Resumen de búsqueda SSASUR,
el PDF combinado y los PDF individuales. Repórtale al usuario un resumen
agregado (cuántos pacientes, cuántos reemplazados, cuántos sin receta vigente
encontrada) — nunca listes RUT ni pacientes uno por uno en el chat.

## Datos sensibles (Ley 19.628)

Los RUT y nombres de pacientes viven solo en los xlsx/PDF locales generados.
**Nunca** pegar filas con RUT en el chat — solo confirmar cantidades y, si
hace falta señalar un caso puntual, usar el nombre del paciente sin el RUT.
