/**
 * SISTEMA DE GESTIÓN TERRITORIAL - HOSPITAL DE PITRUFQUÉN
 * Automatizaciones de Farmacia Central y Despacho Periférico.
 *
 * Basado en el borrador original del usuario, con 5 correcciones aplicadas
 * (ver plan de migración C:\Users\danie\.claude\plans\peaceful-forging-floyd.md):
 *   1. filtrarCadenaFrio() no existía — se implementó.
 *   2. ordenarPorReceta() ya no renumera correlativos ya asignados (rompía el
 *      control físico de valijas ya despachadas).
 *   3. El prefijo de mes ya no está hardcodeado a "Septiembre" — se deriva del
 *      nombre real de la hoja.
 *   4. onEdit() ya no borra la columna F (Establecimiento_Destino, dato manual)
 *      al vaciar el RUT — solo limpia D:E (las 2 fórmulas). Además soporta
 *      pegado de RUT en varias filas a la vez.
 *   5. generarRolloverMes() ya no promete en el comentario algo que no hacía
 *      (arrastre automático de pacientes) — Opción A del plan: crea la hoja
 *      vacía desde la plantilla y el QF completa a mano, apoyándose en la
 *      columna H (alerta) de la hoja anterior. También actualiza
 *      Parametros!G2 ("Mes Activo") para que el Dashboard con INDIRECT
 *      apunte solo ahí.
 *
 * Agregado 04-09-2026 (pedido directo del usuario):
 *   6. verPendientes() — filtra la hoja activa por Estado Despacho =
 *      "pendiente por falta stock" (Columna K), para ver de un vistazo las
 *      recetas que quedaron con algún fármaco sin entregar.
 *   7. generarRolloverMes() ahora reprotege Nombre/Teléfono (D:E) en la hoja
 *      nueva — copyTo() NO copia protecciones de rango, se confirmó con una
 *      prueba directa (crear+borrar una copia) que la copia queda con 0
 *      protecciones si no se rehacen a mano.
 *
 * Agregado 04-09-2026 (bug real detectado):
 *   8. generarRolloverMes() ya no deja que se tipee el nombre completo de la
 *      hoja a mano — bug real: se creó "despachos octubre" (minúscula, sin
 *      guion bajo) en vez de "Despachos_Octubre", y Parametros!G2 quedó
 *      apuntando a ese nombre, rompiendo la convención que usa
 *      _esHojaDespachos()/_numeroMesDesdeNombreHoja() y el INDIRECT del
 *      Dashboard. Ahora solo se pide el MES (ej. "Octubre") y la función arma
 *      "Despachos_<Mes>" sola, validando contra MESES_ES — no hay forma de
 *      que el nombre salga mal escrito.
 *   9. verPendientes() ahora filtra por Estado Despacho = "pendiente por
 *      falta stock" O Fármaco Pendiente/Stock no vacío (en vez de solo
 *      Estado). Ese segundo campo lo llena gt_pendientes_maestro_pacientes.py
 *      cruzando contra informe_completo_recetas*.csv — la Cantidad Pendiente
 *      real por prescripción, no un estado asignado a mano. Sin esto, una
 *      receta marcada "Enviada" con un despacho parcial (un fármaco puntual
 *      todavía pendiente) no aparecía en el filtro aunque el CSV sí mostrara
 *      la prescripción pendiente.
 *
 * Corregido 04-09-2026 (bug real detectado tras pegar el punto 9):
 *   10. La fórmula OR(...) de verPendientes() usaba "," como separador de
 *       argumentos — el Sheet tiene locale es_ES, que exige ";" (mismo bug
 *       ya resuelto antes en el colchón de Despachos_Septiembre, esta vez se
 *       me pasó porque escribí la fórmula directo sin probarla contra la
 *       API). Con "," la fórmula es inválida y el filtro ocultaba TODAS las
 *       filas en vez de solo las no-pendientes — confirmado probando ambas
 *       variantes directo contra el Sheet real vía la API antes de corregir.
 */

var MESES_ES = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO",
                 "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"];

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🏥 Gestión Territorial')
    .addItem('🔄 Ordenar por N° de Receta', 'ordenarPorReceta')
    .addItem('❄️ Filtrar Cadena de Frío (Mes)', 'filtrarCadenaFrio')
    .addItem('⏳ Ver Recetas Pendientes (Falta Stock)', 'verPendientes')
    .addSeparator()
    .addItem('🗓️ Generar Rollover (Pase al Mes Siguiente)', 'generarRolloverMes')
    .addToUi();
}

/** Deriva el número de mes (01-12) a partir del nombre de la hoja (ej.
 * "Despachos_Septiembre" -> "09"). Devuelve null si no reconoce ningún mes —
 * así el llamador puede decidir qué hacer en vez de asumir "XX" siempre. */
function _numeroMesDesdeNombreHoja(nombreHoja) {
  var upper = nombreHoja.toUpperCase();
  for (var i = 0; i < MESES_ES.length; i++) {
    if (upper.indexOf(MESES_ES[i]) !== -1) {
      var n = i + 1;
      return (n < 10 ? "0" + n : "" + n);
    }
  }
  return null;
}

function _esHojaDespachos(nombreHoja) {
  return nombreHoja.indexOf("Despachos") !== -1 || nombreHoja.indexOf("Plantilla_Mes_Nuevo") !== -1;
}

/**
 * Ordena la hoja activa correlativamente por N° de Receta (Col B) y asigna
 * N° Envío GT (Col A) SOLO a las filas que todavía no lo tengan — nunca
 * reasigna un correlativo ya impreso/usado en una valija física.
 */
function ordenarPorReceta() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var name = sheet.getName();

  if (!_esHojaDespachos(name)) {
    SpreadsheetApp.getUi().alert("Ubíquese en una hoja de despachos (Despachos_<Mes> o Plantilla_Mes_Nuevo).");
    return;
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 4) return;

  var numFilas = lastRow - 3;
  var range = sheet.getRange(4, 1, numFilas, 18);

  // Ordena las 18 columnas completas por Col B (N° de Receta) — el correlativo
  // de Col A viaja con su fila (comportamiento nativo de sort), así que un
  // correlativo ya asignado sigue ligado a la receta correcta tras ordenar.
  range.sort({column: 2, ascending: true});

  var colA = sheet.getRange(4, 1, numFilas, 1).getValues();
  var maxExistente = 0;
  for (var i = 0; i < colA.length; i++) {
    var v = String(colA[i][0] || "").trim();
    var m = v.match(/(\d+)\s*$/);
    if (m) {
      var n = parseInt(m[1], 10);
      if (n > maxExistente) maxExistente = n;
    }
  }

  var prefijoMes = _numeroMesDesdeNombreHoja(name) || "XX";
  var prefijo = "GT-" + prefijoMes + "-";
  var siguiente = maxExistente + 1;
  var actualizados = 0;

  for (var i = 0; i < colA.length; i++) {
    if (String(colA[i][0] || "").trim() === "") {
      var pad = siguiente < 10 ? "00" + siguiente : (siguiente < 100 ? "0" + siguiente : "" + siguiente);
      colA[i][0] = prefijo + pad;
      siguiente++;
      actualizados++;
    }
  }
  sheet.getRange(4, 1, numFilas, 1).setValues(colA);

  SpreadsheetApp.getActiveSpreadsheet().toast(
    "Ordenado por N° de Receta. " + actualizados + " fila(s) nueva(s) recibieron correlativo — " +
    "los ya asignados no se tocaron.", "Éxito", 4);
}

/** Filtra la hoja activa mostrando solo las filas con Refrigerado = "SÍ"
 * (Columna L). Usa un filtro nativo de Sheets para poder revertirlo desde
 * el menú Datos > Filtros sin perder los datos ocultos. */
function filtrarCadenaFrio() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var name = sheet.getName();
  if (!_esHojaDespachos(name)) {
    SpreadsheetApp.getUi().alert("Ubíquese en una hoja de despachos (Despachos_<Mes> o Plantilla_Mes_Nuevo).");
    return;
  }
  var lastRow = sheet.getLastRow();
  if (lastRow < 4) return;

  var range = sheet.getRange(3, 1, lastRow - 2, 18); // fila 3 = encabezado, filtro incluido
  var existente = sheet.getFilter();
  if (existente) existente.remove();
  var filter = range.createFilter();
  var criteria = SpreadsheetApp.newFilterCriteria().whenTextEqualTo("SÍ").build();
  filter.setColumnFilterCriteria(12, criteria); // Columna L = Refrigerado

  SpreadsheetApp.getActiveSpreadsheet().toast(
    "Filtro aplicado: solo Refrigerado = SÍ. Quítalo desde Datos > Quitar filtro.", "Cadena de Frío", 4);
}

/** Filtra la hoja activa mostrando las filas con Estado Despacho =
 * "pendiente por falta stock" (Columna K, criterio manual del QF) O con algo
 * escrito en Fármaco Pendiente/Stock (Columna N, criterio real: lo cruza
 * gt_pendientes_maestro_pacientes.py contra la Cantidad Pendiente por
 * prescripción de informe_completo_recetas*.csv, no un estado a mano).
 * Usa la unión de ambas para no perder un despacho PARCIAL — una receta
 * puede estar marcada "Enviada" en general y aun así tener un fármaco
 * puntual con Cantidad Pendiente > 0 en el informe completo. Mismo mecanismo
 * que filtrarCadenaFrio(): filtro nativo de Sheets, reversible desde
 * Datos > Quitar filtro sin perder datos. */
function verPendientes() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var name = sheet.getName();
  if (!_esHojaDespachos(name)) {
    SpreadsheetApp.getUi().alert("Ubíquese en una hoja de despachos (Despachos_<Mes> o Plantilla_Mes_Nuevo).");
    return;
  }
  var lastRow = sheet.getLastRow();
  if (lastRow < 4) return;

  var range = sheet.getRange(3, 1, lastRow - 2, 18); // fila 3 = encabezado, filtro incluido
  var existente = sheet.getFilter();
  if (existente) existente.remove();
  var filter = range.createFilter();
  // Formula relativa a la fila 3 (encabezado del rango) — Sheets la ajusta
  // fila por fila igual que un formato condicional. $K/$N fijan columna, la
  // fila queda relativa para que el desplazamiento sea automático.
  // OJO: separador de argumentos ";" (no ","), locale es_ES del Sheet —
  // con "," la fórmula es inválida y el filtro oculta TODO (bug real
  // 04-09-2026, confirmado probando ambas variantes contra la API).
  var criteria = SpreadsheetApp.newFilterCriteria()
    .whenFormulaSatisfied('=OR($K3="pendiente por falta stock";$N3<>"")')
    .build();
  filter.setColumnFilterCriteria(11, criteria); // se ancla en K, pero la formula mira K y N

  SpreadsheetApp.getActiveSpreadsheet().toast(
    "Filtro aplicado: Estado 'pendiente por falta stock' O Fármaco Pendiente/Stock con dato. " +
    "Quítalo desde Datos > Quitar filtro.", "Recetas Pendientes", 4);
}

/**
 * Sanitiza el RUT al digitarlo (Columna C). Soporta edición de una celda y
 * pegado de un rango de varias filas a la vez.
 *
 * OJO: NO toca las columnas D:E (Nombre/Teléfono) cuando el RUT queda vacío.
 * Versión anterior las limpiaba con clearContent() — eso borra la fórmula
 * de esa celda para siempre, no solo el valor mostrado. Es innecesario:
 * la fórmula de D/E ya trae IF(C="","",...), así que se autolimpia sola en
 * cuanto C queda vacío. Bug real detectado en pruebas 03-09-2026: al borrar
 * y volver a escribir un RUT, D:E quedaban en blanco permanente porque la
 * fórmula ya no existía para recalcular.
 */
function onEdit(e) {
  var sheet = e.source.getActiveSheet();
  var name = sheet.getName();
  if (name === "Maestro_Pacientes" || name === "Dashboard" || name === "Parametros") return;
  if (!_esHojaDespachos(name)) return;

  var range = e.range;
  if (range.getColumn() > 3 || range.getColumn() + range.getNumColumns() - 1 < 3) return; // no toca Col C

  var startRow = Math.max(range.getRow(), 4);
  var endRow = range.getRow() + range.getNumRows() - 1;
  if (endRow < 4) return;

  var colC = sheet.getRange(startRow, 3, endRow - startRow + 1, 1);
  var valores = colC.getValues();
  var cambiados = false;

  for (var i = 0; i < valores.length; i++) {
    var rawRut = String(valores[i][0] || "").trim();
    if (!rawRut || rawRut === "undefined") continue; // la fórmula de D:E ya muestra "" sola
    var cleanRut = rawRut.replace(/[.\s\u00A0]/g, "").toUpperCase();
    if (cleanRut.indexOf("-") === -1 && cleanRut.length >= 8 && cleanRut.length <= 9) {
      cleanRut = cleanRut.slice(0, -1) + "-" + cleanRut.slice(-1);
    }
    if (cleanRut !== rawRut) {
      valores[i][0] = cleanRut;
      cambiados = true;
    }
  }
  if (cambiados) colC.setValues(valores);
}

/**
 * Pase de Mes: crea la hoja del mes nuevo a partir de Plantilla_Mes_Nuevo y
 * actualiza Parametros!G2 (Mes Activo) para que el Dashboard apunte ahí sin
 * tocar ninguna fórmula a mano.
 *
 * OJO: esta función NO arrastra pacientes activos del mes anterior ni
 * incrementa el numerador de Periodo Receta automáticamente (ej. 4/12 -> 5/12)
 * — decidir quién sigue vigente es un criterio clínico, no mecánico. El QF
 * completa los despachos del nuevo mes a mano, apoyándose en la Columna H
 * (Alerta Tratamiento) de la hoja del mes anterior para saber quién debía
 * renovar. Si más adelante se quiere automatizar el arrastre, es un cambio
 * aparte que hay que revisar con el QF antes de activarlo.
 *
 * OJO 2: Plantilla_Mes_Nuevo!D:E está protegido (editar SOLO en
 * Maestro_Pacientes), pero copyTo() NO copia protecciones de rango a la hoja
 * nueva (confirmado por prueba directa 03-09-2026: se creó y borró una copia
 * de prueba, la copia tenía 0 protecciones) — por eso esta función vuelve a
 * protegerlas explícitamente sobre la hoja recién creada.
 */
function generarRolloverMes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var prompt = ui.prompt("Pase de Mes", "Ingrese SOLO el mes de la nueva hoja (ej: Octubre):", ui.ButtonSet.OK_CANCEL);
  if (prompt.getSelectedButton() !== ui.Button.OK) return;

  var mesIngresado = prompt.getResponseText().trim();
  // Ningún mes en español lleva tilde (SEPTIEMBRE, no SEPTIÉMBRE), así que
  // basta con mayúsculas — no hace falta normalizar acentos.
  var idxMes = MESES_ES.indexOf(mesIngresado.toUpperCase());
  if (idxMes === -1) {
    ui.alert("No reconozco '" + mesIngresado + "' como un mes válido. Escribe solo el nombre del mes " +
             "(ej: Octubre), sin 'Despachos_' ni otro texto — la función arma el nombre de la hoja sola.");
    return;
  }
  var mesCapitalizado = MESES_ES[idxMes].charAt(0) + MESES_ES[idxMes].slice(1).toLowerCase();
  var nuevoNombre = "Despachos_" + mesCapitalizado;

  if (ss.getSheetByName(nuevoNombre)) {
    ui.alert("La hoja '" + nuevoNombre + "' ya existe.");
    return;
  }

  var plantilla = ss.getSheetByName("Plantilla_Mes_Nuevo");
  if (!plantilla) {
    ui.alert("No se encontró la hoja 'Plantilla_Mes_Nuevo'.");
    return;
  }

  var nuevaHoja = plantilla.copyTo(ss).setName(nuevoNombre);
  nuevaHoja.showSheet();
  ss.setActiveSheet(nuevaHoja);

  var avisoProteccion = "";
  try {
    var rangoNombreTelefono = nuevaHoja.getRange(4, 4, Math.max(nuevaHoja.getMaxRows() - 3, 1), 2);
    var proteccion = rangoNombreTelefono.protect().setDescription(
      "Nombre/Telefono - editar SOLO en Maestro_Pacientes (evita que el numero diverja entre hojas)");
    proteccion.removeEditors(proteccion.getEditors());
  } catch (err) {
    avisoProteccion = "\n\nOJO: no se pudo proteger Nombre/Teléfono automáticamente (" + err.message + "). " +
                       "Protégelas a mano: seleccionar D4:E, Datos > Hojas y rangos protegidos.";
  }

  var wsParam = ss.getSheetByName("Parametros");
  if (wsParam) {
    wsParam.getRange("G2").setValue(nuevoNombre);
  } else {
    ui.alert("Hoja nueva creada, pero no se encontró 'Parametros' para actualizar el Mes Activo — " +
             "actualízalo a mano en Parametros!G2.");
  }

  ui.alert("Nueva hoja '" + nuevoNombre + "' creada con fórmulas precargadas y Nombre/Teléfono protegidos. " +
           "El Dashboard ya apunta a este mes (Parametros!G2). " +
           "Recuerda: los despachos del nuevo mes se completan a mano, apoyándote en la " +
           "columna 'Alerta Tratamiento' del mes anterior." + avisoProteccion);
}
