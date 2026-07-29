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
  var ss    = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('RegOps') || ss.getSheets()[0];
  if (sheet) ss.setActiveSheet(sheet);
}


/**
 * ======================================================================
 * ON EDIT (Disparador Automático Único)
 * ======================================================================
 */

function onEdit(e) {
  if (!e || !e.range) return;

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(3000)) return;

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
      return;
    }

    // 3. Datestamp automático en RegOps
    if (sheetName === 'RegOps' &&
        range.getNumRows() === 1 && range.getNumColumns() === 1 &&
        col === CONFIG.TEXT_COLUMN) {

      var tsCell = sheet.getRange(row, CONFIG.TIMESTAMP_COLUMN);
      tsCell.setValue(new Date()).setNumberFormat('yyyy-MM-dd HH:mm');
    }

  } catch (err) {
    Logger.log('onEdit error: ' + err);
  } finally {
    lock.releaseLock();
  }
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
  var isRegOps = rawName.includes('regops');

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

      if (isRegOps) {
        // En RegOps dibuja la línea SOLAMENTE desde la Columna A (1) a la J (10)
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