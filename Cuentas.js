/**
 * Completa en CtasDefinicion de v1 la estructura del plan de cuentas de v2.
 * Solo actualiza ID, Grupo, Estado y Orden cuando el nombre coincide exactamente.
 * Las cuentas exclusivas de v1 conservan sus datos actuales.
 */
function actualizarEstructuraCuentasDesdeV2() {
  var ID_REGOPS_V2 = '18doMyPGnRoUh6uIPsXy-MkYS7X1477gING3DtHFGn_A';
  var NOMBRE_HOJA = 'CtasDefinicion';

  var ssV1 = SpreadsheetApp.getActive();
  var hojaV1 = ssV1.getSheetByName(NOMBRE_HOJA);
  if (!hojaV1) throw new Error('No se encontró CtasDefinicion en RegOps v1.');

  var ssV2 = SpreadsheetApp.openById(ID_REGOPS_V2);
  var hojaV2 = ssV2.getSheetByName(NOMBRE_HOJA);
  if (!hojaV2) throw new Error('No se encontró CtasDefinicion en RegOps v2.');

  var ultimaFilaV2 = hojaV2.getLastRow();
  if (ultimaFilaV2 < 2) throw new Error('El plan de cuentas de RegOps v2 está vacío.');

  var datosV2 = hojaV2.getRange(2, 1, ultimaFilaV2 - 1, 5).getValues();
  var metadataPorCuenta = new Map();

  datosV2.forEach(function(fila) {
    var cuenta = String(fila[0] || '').trim();
    if (!cuenta) return;
    metadataPorCuenta.set(cuenta, [fila[1], fila[2], fila[3], fila[4]]);
  });

  var ultimaFilaV1 = hojaV1
    .getRange(hojaV1.getMaxRows(), 1)
    .getNextDataCell(SpreadsheetApp.Direction.UP)
    .getRow();

  if (ultimaFilaV1 < 3) throw new Error('El plan de cuentas de RegOps v1 está vacío.');

  var cantidadFilas = ultimaFilaV1 - 2;
  var datosV1 = hojaV1.getRange(3, 1, cantidadFilas, 5).getValues();
  var coincidencias = 0;
  var sinCoincidencia = [];

  var salida = datosV1.map(function(fila) {
    var cuenta = String(fila[0] || '').trim();
    if (!cuenta) return fila;

    var metadata = metadataPorCuenta.get(cuenta);
    if (!metadata) {
      sinCoincidencia.push(cuenta);
      return fila;
    }

    coincidencias++;
    return [cuenta, metadata[0], metadata[1], metadata[2], metadata[3]];
  });

  hojaV1.getRange('A2:E2').setValues([[
    'Cuenta', 'ID', 'Grupo', 'Estado (1=Activa, 0=Inactiva)', 'Orden'
  ]]);
  hojaV1.getRange(3, 1, cantidadFilas, 5).setValues(salida);

  hojaV1.getRange('A2:E2')
    .setBackground('#424242')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  hojaV1.getRange(3, 2, cantidadFilas, 1).setNumberFormat('0');
  hojaV1.getRange(3, 4, cantidadFilas, 2).setNumberFormat('0');
  hojaV1.setFrozenRows(2);
  hojaV1.setColumnWidth(1, 310);
  hojaV1.setColumnWidth(2, 70);
  hojaV1.setColumnWidth(3, 110);
  hojaV1.setColumnWidth(4, 190);
  hojaV1.setColumnWidth(5, 80);

  SpreadsheetApp.flush();

  var mensaje = coincidencias + ' cuentas actualizadas desde v2';
  if (sinCoincidencia.length) {
    mensaje += '; ' + sinCoincidencia.length + ' sin coincidencia exacta';
  }

  ssV1.toast(mensaje, 'RegOps v1', 8);
  Logger.log('Cuentas sin coincidencia exacta: ' + JSON.stringify(sinCoincidencia));
}
