/**
 * ======================================================================
 * CONFIG-NEW
 * ======================================================================
 * Centralised settings — edit here, nowhere else.
 */

var CONFIG = {
  TIMEZONE:             'America/Argentina/Buenos_Aires',
  START_HOUR:           9,
  END_HOUR:             18,
  TEXT_COLUMN:          2, // Column B ("Ingresó")
  TIMESTAMP_COLUMN:     1, // Column A ("Fecha operación")
  AUDIT_SHEET:          '_AUDIT_LOG',
  MASS_EDIT_THRESHOLD:  5,
  BACKUP_FOLDER_ID:     '1uAlMjC9KbNibWzHL4FSbO65Ln6PzIkz4',
  AUDIT_MAX_ROWS:       10000,
  AUDIT_ROTATE_DELETE:  1000,
};


/**
 * ======================================================================
 * ON OPEN
 * ======================================================================
 */

function onOpen() {
  var ss = SpreadsheetApp.getActive();

  prepararHojasRegOpsV1_(ss);

  // El Mayor se actualiza primero. Las tareas de formato y validación pueden
  // tardar sobre planillas grandes y antes impedían que esta llamada se ejecutara.
  actualizarMayorSiCorrespondeV1_();

  configurarSeparadoresNumericosV1_(ss);
  asegurarFormatoOperacionesCompensadas_(ss);
  asegurarModeloRelacionalCuentasV1_(ss);

  var diario = ss.getSheetByName('Diario');
  if (diario) ss.setActiveSheet(diario);
}

/**
 * Usa formato internacional tipo calculadora en todo el archivo:
 * coma para miles y punto para decimales.
 */
function configurarSeparadoresNumericosV1_(ss) {
  ss = ss || SpreadsheetApp.getActive();

  if (ss.getSpreadsheetLocale() !== 'en_US') {
    ss.setSpreadsheetLocale('en_US');
  }

  var diario = ss.getSheetByName('Diario');
  if (diario && diario.getMaxRows() >= 6) {
    var cantidadFilas = diario.getMaxRows() - 5;
    diario.getRange(6, 4, cantidadFilas, 1)
      .setNumberFormat('#,##0.##;[Red](#,##0.##);-');
    diario.getRange(6, 6, cantidadFilas, 1)
      .setNumberFormat('#,##0.##;[Red](#,##0.##);-');
  }
}

/**
 * Renombra la hoja operativa y crea el Mayor sin sobrescribir hojas existentes.
 * Puede ejecutarse más de una vez sin duplicar hojas.
 */
function prepararHojasRegOpsV1_(ss) {
  ss = ss || SpreadsheetApp.getActive();

  if (ss.getName() !== 'RegOps v1') {
    ss.rename('RegOps v1');
  }

  var diario = ss.getSheetByName('Diario');
  var regOps = ss.getSheetByName('RegOps');

  if (!diario && regOps) {
    regOps.setName('Diario');
    diario = regOps;
  }

  if (!diario) {
    throw new Error('No se encontró la hoja "RegOps" ni la hoja "Diario".');
  }

  if (!ss.getSheetByName('Mayor')) {
    ss.insertSheet('Mayor');
  }
}


/**
 * Baja la intensidad visual de las operaciones compensadas:
 * D y F tienen igual importe con signo contrario.
 * La regla se agrega al final para no tapar alertas existentes.
 */
function asegurarFormatoOperacionesCompensadas_(ss) {
  ss = ss || SpreadsheetApp.getActive();
  var diario = ss.getSheetByName('Diario');
  if (!diario) return;

  var formula = '=AND($D6<>"",$F6<>"",ABS($D6+$F6)<0.01)';
  var formulaNegativoD = '=AND($D6<>"",$F6<>"",ABS($D6+$F6)<0.01,D6<0)';
  var formulaNegativoF = '=AND($D6<>"",$F6<>"",ABS($D6+$F6)<0.01,F6<0)';
  var formulasRegOps = {};
  formulasRegOps[formula] = true;
  formulasRegOps[formulaNegativoD] = true;
  formulasRegOps[formulaNegativoF] = true;

  var reglas = diario.getConditionalFormatRules().filter(function(regla) {
    try {
      var condicion = regla.getBooleanCondition();
      if (!condicion) return true;

      var valores = condicion.getCriteriaValues();
      if (!valores || !valores.length) return true;

      var formulaExistente = String(valores[0] || '').replace(/\s/g, '');
      return formulaExistente.indexOf('ABS($D6+$F6)') === -1;
    } catch (err) {
      return true;
    }
  });

  var cantidadFilas = diario.getMaxRows() - 5;
  var rango = diario.getRange(6, 1, cantidadFilas, 7);
  var reglaCompensada = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(formula)
    .setFontColor('#8a8a8a')
    .setBackground('#f7f7f7')
    .setRanges([rango])
    .build();

  var reglaNegativoD = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(formulaNegativoD)
    .setFontColor('#c27a7a')
    .setBackground('#f7f7f7')
    .setRanges([diario.getRange(6, 4, cantidadFilas, 1)])
    .build();

  var reglaNegativoF = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(formulaNegativoF)
    .setFontColor('#c27a7a')
    .setBackground('#f7f7f7')
    .setRanges([diario.getRange(6, 6, cantidadFilas, 1)])
    .build();

  // Sheets aplica primero la regla ubicada más arriba.
  // Los negativos deben evaluarse antes que el gris general de la fila.
  reglas = [reglaNegativoD, reglaNegativoF, reglaCompensada].concat(reglas);
  diario.setConditionalFormatRules(reglas);

  // Todo el efecto visual queda resuelto por estas reglas; no se recorren filas.
}

/**
 * Cancela cualquier limpieza anterior y deja solamente las reglas
 * de formato condicional. No recorre las filas del Diario.
 */
function limpiarFormatoDiarioV1() {
  detenerLimpiezaDiarioV1_();

  var ss = SpreadsheetApp.getActive();
  asegurarFormatoOperacionesCompensadas_(ss);
  SpreadsheetApp.flush();
  ss.toast('Formato condicional del Diario instalado.', 'RegOps', 4);
}

function detenerLimpiezaDiarioV1_() {
  PropertiesService.getDocumentProperties()
    .deleteProperty('REGOPS_LIMPIEZA_PROXIMA_FILA');

  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'limpiarFormatoDiarioV1') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

function instalarFormatoOperacionesCompensadas() {
  asegurarFormatoOperacionesCompensadas_(SpreadsheetApp.getActive());
}


/**
 * ======================================================================
 * ON EDIT (Disparador Automático Único)
 * ======================================================================
 */

function onEdit(e) {
  if (!e || !e.range) return;

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return;

  try {
    var range      = e.range;
    var sheet      = range.getSheet();
    var sheetName  = sheet.getName();

    // Ignorar hoja de auditoría
    if (sheetName === CONFIG.AUDIT_SHEET) return;

    var row        = range.getRow();
    var col        = range.getColumn();
    var numCells   = range.getNumRows() * range.getNumColumns();
    var isMassEdit = numCells > CONFIG.MASS_EDIT_THRESHOLD;

    // 1. SI SE EDITA A2 EN HOJA DE REPORTE -> Recalcular líneas de mes
    if (range.getA1Notation() === 'A2' && 
       (sheetName.toLowerCase().includes('report') || sheetName.toLowerCase().includes('reporte'))) {
      Utilities.sleep(1200); // Esperar que QUERY traiga datos
      procesarHoja_(sheet);
      return;
    }

    // C1 de Mayor funciona como botón de actualización.
    if (sheetName === 'Mayor' && range.getA1Notation() === 'C1') {
      if (String(e.value || '').toUpperCase() === 'TRUE') {
        range.setValue(false);
        actualizarMayorV1();
      }
      return;
    }

    if (row <= 1) return;

    var ss        = SpreadsheetApp.getActive();
    var userEmail = getUserEmail_();
    var isOwner   = isOwner_(ss, userEmail);

    // 2. Control de horario comercial
    if (!isOwner && !isWithinBusinessHours_()) {
      var reason = isMassEdit
        ? 'BLOCKED — mass edit outside hours'
        : 'BLOCKED — single edit outside hours';

      logAudit_(ss, e, userEmail, reason);
      revertEdit_(e);
      notify_(ss, 'Editing is only allowed Mon–Fri 09:00–18:00 (Argentina time).');
      return;
    }

    if (isMassEdit) {
      logAudit_(ss, e, userEmail, 'ALLOWED — mass edit');
    }

    // 3. Mantener el plan de cuentas como tabla maestra.
    var primeraColumnaEditada = range.getColumn();
    var ultimaColumnaEditada = range.getLastColumn();
    var tocaEstructuraCuentas =
      sheetName === 'CtasDefinicion' &&
      range.getLastRow() >= 3 &&
      primeraColumnaEditada <= 5 &&
      ultimaColumnaEditada >= 1;

    if (tocaEstructuraCuentas) {
      procesarEdicionPlanCuentasV1_(ss);
      marcarMayorDesactualizadoV1_();
      return;
    }

    // 4. En Diario, C y E muestran nombres pero guardan ID ocultos.
    if (sheetName === 'Diario') {
      actualizarIdsCuentasEditadasV1_(sheet, range);

      // Una edición de los movimientos sólo marca el Mayor como pendiente.
      // El cálculo pesado se ejecuta al abrir por el propietario o con el botón.
      var tocaMovimientos =
        range.getLastRow() >= 6 &&
        primeraColumnaEditada <= 6 &&
        ultimaColumnaEditada >= 1;
      if (tocaMovimientos) marcarMayorDesactualizadoV1_();
    }

    // 5. Timestamp automático en Diario.
    // El disparador simple cubre al propietario; el instalable permite
    // escribir en A aunque la edición la haga otro usuario.
    actualizarTimestampIngresoV1_(sheet, range);

    if (isMassEdit) return;

  } catch (err) {
    Logger.log('onEdit error: ' + err);
    try {
      SpreadsheetApp.getActive().toast(String(err.message || err), 'RegOps v1', 8);
    } catch (_) {}
  } finally {
    lock.releaseLock();
  }
}


/**
 * Escribe o borra el timestamp de A cuando se modifica "Ingresó" en B.
 * Es idempotente: puede ser llamado por el disparador simple e instalable.
 */
function actualizarTimestampIngresoV1_(sheet, range) {
  if (!sheet || sheet.getName() !== 'Diario') return;

  var primeraColumna = range.getColumn();
  var ultimaColumna = range.getLastColumn();
  var incluyeColumnaIngreso =
    primeraColumna <= CONFIG.TEXT_COLUMN &&
    ultimaColumna >= CONFIG.TEXT_COLUMN;

  if (!incluyeColumnaIngreso) return;

  var primeraFila = Math.max(range.getRow(), 6);
  var ultimaFila = range.getLastRow();
  if (ultimaFila < primeraFila) return;

  var cantidadFilas = ultimaFila - primeraFila + 1;
  var nombres = sheet
    .getRange(primeraFila, CONFIG.TEXT_COLUMN, cantidadFilas, 1)
    .getDisplayValues();
  var rangoTimestamps = sheet
    .getRange(primeraFila, CONFIG.TIMESTAMP_COLUMN, cantidadFilas, 1);
  var timestampsActuales = rangoTimestamps.getValues();
  var ahora = new Date();

  var timestampsNuevos = nombres.map(function(fila, indice) {
    var nombre = String(fila[0] || '').trim();

    if (nombre === '') return [''];
    if (timestampsActuales[indice][0]) return [timestampsActuales[indice][0]];
    return [ahora];
  });

  rangoTimestamps
    .setValues(timestampsNuevos)
    .setNumberFormat('yyyy-MM-dd HH:mm');
}

/**
 * Handler del disparador instalable. Se ejecuta con los permisos
 * del propietario que lo instala.
 */
function timestampOnEditRegOpsV1(e) {
  if (!e || !e.range) return;
  actualizarTimestampIngresoV1_(e.range.getSheet(), e.range);
}

/**
 * Instala una sola vez el disparador autorizado para todos los editores.
 */
function instalarTimestampParaTodosV1() {
  var ss = SpreadsheetApp.getActive();

  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'timestampOnEditRegOpsV1') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('timestampOnEditRegOpsV1')
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  ss.toast('Timestamp activado para todos los usuarios.', 'RegOps', 5);
}


/**
 * ======================================================================
 * DIBUJAR LÍNEAS HORIZONTALES DE MES
 * ======================================================================
 */

function aplicarLineasDeMesEnTodasLasHojas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();

  sheets.forEach(function(sheet) {
    if (sheet.getName().startsWith('_')) return;
    procesarHoja_(sheet);
  });
}

function procesarHoja_(sheet) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  if (lastRow < 4 || lastCol === 0) return;

  var rawName = sheet.getName().toLowerCase().trim();
  var isReport = rawName.includes('report') || rawName.includes('reporte');
  var isDiario = rawName === 'diario';

  var colFecha = isReport ? 2 : 1;
  var startRow = 4;

  // 1. Limpia todos los bordes de la tabla antes de re-dibujar
  var totalRange = sheet.getRange(startRow, 1, lastRow - startRow + 1, lastCol);
  totalRange.setBorder(false, false, false, false, false, false);

  var fechasValues = sheet.getRange(startRow, colFecha, lastRow - startRow + 1, 1).getDisplayValues();

  // 2. Recorre fila por fila buscando los cierres de mes
  for (var i = 0; i < fechasValues.length - 1; i++) {
    var str1 = fechasValues[i][0];
    var str2 = fechasValues[i + 1][0];

    var clave1 = obtenerClaveMes_(str1);
    var clave2 = obtenerClaveMes_(str2);

    if (clave1 && clave2 && clave1 !== clave2) {
      var rowNum = startRow + i;

      if (isDiario) {
        // En Diario dibuja la línea SOLAMENTE desde la Columna A (1) a la J (10)
        sheet.getRange(rowNum, 1, 1, 10)
             .setBorder(null, null, true, null, null, null, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      } else {
        // En reportes dibuja la línea en todo el ancho de la tabla
        sheet.getRange(rowNum, 1, 1, lastCol)
             .setBorder(null, null, true, null, null, null, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      }
    }
  }
}

/**
 * Obtiene la clave "YYYY-MM" para comparar meses
 */
function obtenerClaveMes_(valStr) {
  if (!valStr || valStr.trim() === '') return null;

  var d = new Date(valStr);
  if (!isNaN(d.getTime())) {
    return d.getFullYear() + '-' + (d.getMonth() + 1);
  }

  // Comprobación por patrones Regex si es un string formateado
  var matchISO = valStr.match(/(\d{4})[-/](\d{1,2})/);
  if (matchISO) return matchISO[1] + '-' + parseInt(matchISO[2], 10);

  var matchLAT = valStr.match(/(\d{1,2})[-/](\d{1,2})[-/](\d{4})/);
  if (matchLAT) return matchLAT[3] + '-' + parseInt(matchLAT[2], 10);

  return null;
}


/**
 * ======================================================================
 * FUNCIONES AUXILIARES (Auditoría, Horarios, Backups)
 * ======================================================================
 */

function isWithinBusinessHours_() {
  var now  = new Date();
  var hour = Number(Utilities.formatDate(now, CONFIG.TIMEZONE, 'H'));
  var day  = Number(Utilities.formatDate(now, CONFIG.TIMEZONE, 'u'));

  return day >= 1 && day <= 5 &&
         hour >= CONFIG.START_HOUR && hour < CONFIG.END_HOUR;
}

function getUserEmail_() {
  try {
    return Session.getActiveUser().getEmail() ||
           Session.getEffectiveUser().getEmail() ||
           'unknown';
  } catch (_) {
    return 'unknown';
  }
}

function isOwner_(ss, userEmail) {
  try {
    return ss.getOwner().getEmail() === userEmail;
  } catch (_) {
    return false;
  }
}

function revertEdit_(e) {
  try {
    var range = e.range;
    if (range.getNumRows() === 1 && range.getNumColumns() === 1 && typeof e.oldValue !== 'undefined') {
      range.setValue(e.oldValue);
    }
  } catch (err) {
    Logger.log('revertEdit_ error: ' + err);
  }
}

function logAudit_(ss, e, userEmail, status) {
  try {
    var auditSheet = ss.getSheetByName(CONFIG.AUDIT_SHEET);

    if (!auditSheet) {
      auditSheet = ss.insertSheet(CONFIG.AUDIT_SHEET);
      auditSheet.appendRow([
        'Timestamp', 'User', 'Sheet', 'Range',
        'Cells Edited', 'Old Value (single only)',
        'New Value (single only)', 'Status'
      ]);
      auditSheet.hideSheet();
    }

    var timestamp = Utilities.formatDate(
      new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss'
    );

    auditSheet.appendRow([
      timestamp,
      userEmail,
      e.range.getSheet().getName(),
      e.range.getA1Notation(),
      e.range.getNumRows() * e.range.getNumColumns(),
      typeof e.oldValue !== 'undefined' ? e.oldValue : '',
      typeof e.value    !== 'undefined' ? e.value    : '',
      status
    ]);

  } catch (err) {
    Logger.log('logAudit_ error: ' + err);
  }
}

function backupRegOps() {
  var ss          = SpreadsheetApp.getActive();
  var file        = DriveApp.getFileById(ss.getId());
  var timestamp   = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd_HH-mm');
  var backupName  = ss.getName() + '_BACKUP_' + timestamp;

  if (CONFIG.BACKUP_FOLDER_ID) {
    var folder = DriveApp.getFolderById(CONFIG.BACKUP_FOLDER_ID);
    file.makeCopy(backupName, folder);
    cleanOldBackups_(folder, 30);
  } else {
    file.makeCopy(backupName);
  }
}

function cleanOldBackups_(folder, days) {
  var cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - days);
  var files = folder.getFiles();

  while (files.hasNext()) {
    var file = files.next();
    if (file.getName().includes('_BACKUP_') && file.getDateCreated() < cutoff) {
      file.setTrashed(true);
    }
  }
}

function cleanAuditDaily() {
  var ss         = SpreadsheetApp.getActive();
  var auditSheet = ss.getSheetByName(CONFIG.AUDIT_SHEET);

  if (!auditSheet) return;

  var lastRow = auditSheet.getLastRow();

  if (lastRow > CONFIG.AUDIT_MAX_ROWS) {
    auditSheet.deleteRows(2, CONFIG.AUDIT_ROTATE_DELETE);
    Logger.log('Mantenimiento nocturno: Log de auditoría rotado.');
  }
}

function notify_(ss, message) {
  ss.toast(message, 'Access Restricted', 5);
}