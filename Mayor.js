function actualizarMayorV1() {
  var ss = SpreadsheetApp.getActive();
  var diario = ss.getSheetByName('Diario');
  var mayor = ss.getSheetByName('Mayor');
  var definicion = ss.getSheetByName('CtasDefinicion');

  if (!diario || !mayor || !definicion) {
    throw new Error('Se requieren las hojas Diario, Mayor y CtasDefinicion.');
  }

  var inicio = new Date();
  var mesesFuente = obtenerMesesMayorV1_(diario);
  var meses = mesesFuente.slice().reverse();
  var cantidadMeses = meses.length;
  var ultimaColumna = 3 + cantidadMeses;
  var tasaArsActual = Number(diario.getRange('L2').getValue()) || 1;
  var tasaEuroUsd = Number(diario.getRange('L1').getValue()) / tasaArsActual || 1;
  var tasaBrlUsd = 0.20;

  var nombresPlan = definicion
    .getRange(3, 1, Math.max(definicion.getLastRow() - 2, 1), 1)
    .getDisplayValues()
    .map(function(fila) { return String(fila[0] || '').trim(); })
    .filter(function(nombre) { return nombre !== ''; });

  var resumenFuente = diario.getRange(6, 11, 8, 2 + cantidadMeses).getValues();
  var cuentasFuente = diario
    .getRange(15, 11, nombresPlan.length, 2 + cantidadMeses)
    .getValues();

  var filas = [];
  var tipos = [];

  filas.push(nuevaFilaMayorV1_(ultimaColumna));
  tipos.push('libre');

  var filaMeses = nuevaFilaMayorV1_(ultimaColumna);
  filaMeses[0] = 'T/C Euro/USD';
  filaMeses[2] = tasaEuroUsd;
  filas.push(filaMeses);
  tipos.push('meses');

  var filaBrl = nuevaFilaMayorV1_(ultimaColumna);
  filaBrl[0] = 'T/C BRL/USD';
  filaBrl[2] = tasaBrlUsd;
  filas.push(filaBrl);
  tipos.push('tasa');

  var encabezado = nuevaFilaMayorV1_(ultimaColumna);
  encabezado[0] = 'Periodo';
  encabezado[1] = 'Acum (USD)';
  encabezado[2] = 'Acum histórico';
  meses.forEach(function(mes, indice) {
    var fechaMes = mes.valor instanceof Date ? mes.valor : new Date(mes.valor);
    encabezado[3 + indice] = isNaN(fechaMes.getTime()) ? mes.valor : fechaMes;
  });
  filas.push(encabezado);
  tipos.push('encabezado');

  var tasasArs = nuevaFilaMayorV1_(ultimaColumna);
  tasasArs[0] = 'T/C ARS/USD';
  tasasArs[2] = tasaArsActual;
  meses.forEach(function(mes, indice) {
    tasasArs[3 + indice] = mes.tasa;
  });
  filas.push(tasasArs);
  tipos.push('tasa');

  var filaPatrimonio = nuevaFilaMayorV1_(ultimaColumna);
  filaPatrimonio[0] = 'Patrimonio Neto (USD)';
  filaPatrimonio[1] = Number(diario.getRange('L4').getValue()) || 0;
  filaPatrimonio[2] = 'Resultados (USD)';
  meses.forEach(function(mes, indice) {
    filaPatrimonio[3 + indice] = Number(mes.patrimonio) || 0;
  });
  filas.push(filaPatrimonio);
  tipos.push('patrimonio');

  filas.push(filaTituloMayorV1_('CUENTAS AGRUPADAS', ultimaColumna));
  tipos.push('titulo');

  var resumenCero = [];
  resumenFuente.forEach(function(origen) {
    var nombre = normalizarNombreResumenMayorV1_(origen[0]);
    if (!nombre || /^Retiros /i.test(nombre)) return;

    var fila = convertirFilaFuenteMayorV1_(nombre, origen, cantidadMeses, ultimaColumna, true);
    if (Math.abs(Number(fila[1]) || 0) < 0.5) {
      resumenCero.push(fila);
    } else {
      filas.push(fila);
      tipos.push('agrupada');
    }
  });

  filas.push(nuevaFilaMayorV1_(ultimaColumna));
  tipos.push('libre');

  var retiroFuente = resumenFuente[7];
  if (retiroFuente && retiroFuente[0]) {
    var filaRetiro = convertirFilaFuenteMayorV1_(
      'Retiros ARS (USD) en ' + cantidadMeses + ' meses',
      retiroFuente,
      cantidadMeses,
      ultimaColumna,
      true
    );
    filas.push(filaRetiro);
    tipos.push('retiro');
  }

  var cuentas = cuentasFuente.map(function(origen, indice) {
    var nombre = String(origen[0] || nombresPlan[indice] || '').trim();
    return {
      nombre: nombre,
      grupo: grupoCuentaMayorV1_(nombre),
      fila: convertirCuentaMayorV1_(
        nombre,
        origen,
        meses,
        ultimaColumna,
        tasaArsActual,
        tasaEuroUsd,
        tasaBrlUsd
      )
    };
  }).filter(function(cuenta) {
    return cuenta.nombre !== '';
  });

  var cuentasCero = cuentas.filter(function(cuenta) {
    return Math.abs(Number(cuenta.fila[1]) || 0) < 0.5;
  });

  ['ACTIVO', 'PATRIMONIO', 'RESULTADO'].forEach(function(grupo) {
    var grupoCuentas = cuentas.filter(function(cuenta) {
      return cuenta.grupo === grupo && Math.abs(Number(cuenta.fila[1]) || 0) >= 0.5;
    });

    if (!grupoCuentas.length) return;

    filas.push(nuevaFilaMayorV1_(ultimaColumna));
    tipos.push('libre');
    filas.push(filaTituloMayorV1_(grupo, ultimaColumna));
    tipos.push('titulo');

    grupoCuentas.forEach(function(cuenta) {
      cuenta.fila[0] = '\u00a0\u00a0' + cuenta.fila[0];
      filas.push(cuenta.fila);
      tipos.push('cuenta');
    });
  });

  filas.push(nuevaFilaMayorV1_(ultimaColumna));
  tipos.push('libre');
  filas.push(filaTituloMayorV1_('CUENTAS EN CERO', ultimaColumna));
  tipos.push('ceroTitulo');

  resumenCero.concat(cuentasCero.map(function(cuenta) {
    cuenta.fila[0] = '\u00a0\u00a0' + cuenta.fila[0];
    return cuenta.fila;
  })).forEach(function(fila) {
    filas.push(fila);
    tipos.push('cero');
  });

  mayor.clear();
  mayor.getRange(1, 1, filas.length, ultimaColumna).setValues(filas);
  aplicarFormatoMayorV1_(mayor, filas, tipos, ultimaColumna, cantidadMeses);
  asegurarBotonActualizarMayorV1_(ss);
  SpreadsheetApp.flush();

  ss.toast(
    'Mayor actualizado en ' + ((new Date() - inicio) / 1000).toFixed(1) + ' s',
    'RegOps v1',
    4
  );
}

function obtenerMesesMayorV1_(diario) {
  var ancho = Math.max(diario.getLastColumn() - 12, 1);
  var fechas = diario.getRange(5, 13, 1, ancho).getValues()[0];
  var tasas = diario.getRange(2, 13, 1, ancho).getValues()[0];
  var patrimonio = diario.getRange(4, 13, 1, ancho).getValues()[0];
  var meses = [];

  for (var i = 0; i < fechas.length; i++) {
    if (fechas[i] === '' || fechas[i] === null) break;
    meses.push({
      valor: fechas[i],
      tasa: tasas[i],
      patrimonio: patrimonio[i]
    });
  }

  return meses;
}

function convertirFilaFuenteMayorV1_(nombre, origen, cantidadMeses, ultimaColumna, dividirMil) {
  var fila = nuevaFilaMayorV1_(ultimaColumna);
  fila[0] = nombre;
  fila[1] = Number(origen[1]) || 0;

  for (var i = 0; i < cantidadMeses; i++) {
    var valor = Number(origen[2 + cantidadMeses - 1 - i]) || 0;
    fila[3 + i] = dividirMil ? valor / 1000 : valor;
  }

  return fila;
}

function convertirCuentaMayorV1_(nombre, origen, meses, ultimaColumna, tasaArsActual, tasaEuroUsd, tasaBrlUsd) {
  var fila = nuevaFilaMayorV1_(ultimaColumna);
  fila[0] = nombre;
  fila[1] = convertirMonedaMayorV1_(
    Number(origen[1]) || 0,
    nombre,
    tasaArsActual,
    tasaEuroUsd,
    tasaBrlUsd
  );

  for (var i = 0; i < meses.length; i++) {
    var indiceFuente = 2 + meses.length - 1 - i;
    fila[3 + i] = convertirMonedaMayorV1_(
      Number(origen[indiceFuente]) || 0,
      nombre,
      Number(meses[i].tasa) || tasaArsActual,
      tasaEuroUsd,
      tasaBrlUsd
    );
  }

  return fila;
}

function convertirMonedaMayorV1_(importe, nombre, tasaArsUsd, tasaEuroUsd, tasaBrlUsd) {
  var cuenta = String(nombre || '').toUpperCase();

  if (/\bARS\b/.test(cuenta)) return importe / tasaArsUsd / 1000;
  if (/\bEURO\b|\bEUR\b/.test(cuenta)) return importe * tasaEuroUsd / 1000;
  if (/\bBRL\b|\bREAL(?:ES)?\b/.test(cuenta)) return importe * tasaBrlUsd / 1000;
  return importe / 1000;
}

function nuevaFilaMayorV1_(cantidadColumnas) {
  return new Array(cantidadColumnas).fill('');
}

function filaTituloMayorV1_(titulo, cantidadColumnas) {
  var fila = nuevaFilaMayorV1_(cantidadColumnas);
  fila[0] = titulo;
  return fila;
}

function grupoCuentaMayorV1_(nombre) {
  if (/^(CJ|CC)\b/i.test(nombre)) return 'ACTIVO';
  if (/^INV/i.test(nombre)) return 'PATRIMONIO';
  return 'RESULTADO';
}

function normalizarNombreResumenMayorV1_(nombre) {
  return String(nombre || '')
    .replace(/\[USD\]/g, '(USD)')
    .replace(/\s*\(no suma en PN\)\s*/i, '')
    .trim();
}

function formatearMesMayorV1_(valor) {
  var fecha = valor instanceof Date ? valor : new Date(valor);
  if (isNaN(fecha.getTime())) return String(valor || '');

  var nombres = [
    'Enero', 'Feb', 'Marzo', 'Abril', 'Mayo', 'Junio',
    'Julio', 'Agosto', 'Sept', 'Oct', 'Nov', 'Dic'
  ];
  return fecha.getFullYear() + '-' + nombres[fecha.getMonth()];
}

function obtenerAnioMayorV1_(valor) {
  if (valor instanceof Date && !isNaN(valor.getTime())) return String(valor.getFullYear());
  var texto = String(valor || '');
  var coincidencia = texto.match(/\d{4}/);
  return coincidencia ? coincidencia[0] : '';
}

function aplicarFormatoMayorV1_(mayor, filas, tipos, ultimaColumna, cantidadMeses) {
  var rango = mayor.getRange(1, 1, filas.length, ultimaColumna);
  rango.setFontFamily('Nunito').setFontSize(8).setVerticalAlignment('middle');
  mayor.setFrozenRows(6);
  mayor.setFrozenColumns(3);
  mayor.setHiddenGridlines(false);

  mayor.setColumnWidth(1, 270);
  mayor.setColumnWidth(2, 105);
  mayor.setColumnWidth(3, 115);
  if (cantidadMeses > 0) {
    mayor.setColumnWidths(4, cantidadMeses, 88);
  }

  rango.setNumberFormat('#,##0;[Red](#,##0);-');
  mayor.getRange(2, 1, 2, 1).setNumberFormat('@');
  mayor.getRange(2, 3, 2, 1).setNumberFormat('0.00');
  if (cantidadMeses > 0) {
    mayor.getRange(4, 4, 1, cantidadMeses).setNumberFormat('yyyy-mm');
  }
  if (filas.length >= 5) {
    mayor.getRange(5, 2, filas.length - 4, 1).setFontSize(10);
  }

  var fondoAlterno = '#f7f7f7';
  for (var fila = 1; fila <= filas.length; fila++) {
    if (fila % 2 === 0 && tipos[fila - 1] !== 'titulo' && tipos[fila - 1] !== 'ceroTitulo') {
      mayor.getRange(fila, 1, 1, ultimaColumna).setBackground(fondoAlterno);
    }
  }

  tipos.forEach(function(tipo, indice) {
    var fila = indice + 1;
    var rangoFila = mayor.getRange(fila, 1, 1, ultimaColumna);

    if (tipo === 'encabezado') {
      rangoFila.setBackground('#424242').setFontColor('#ffffff').setFontWeight('bold');
    } else if (tipo === 'patrimonio') {
      rangoFila.setFontWeight('bold').setFontSize(10)
        .setBorder(null, null, true, null, null, null, '#b7b7b7', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    } else if (tipo === 'titulo') {
      rangoFila.setFontWeight('bold').setBackground('#ffffff');
    } else if (tipo === 'retiro') {
      rangoFila.setBackground('#cfe2f3');
    } else if (tipo === 'ceroTitulo' || tipo === 'cero') {
      rangoFila.setBackground('#d9ead3');
      if (tipo === 'ceroTitulo') rangoFila.setFontWeight('bold');
    }
  });

  mayor.getRange(1, 2, filas.length, ultimaColumna - 1).setHorizontalAlignment('right');
  mayor.getRange(1, 1, filas.length, 1).setHorizontalAlignment('left');
  mayor.getRange(4, 1, 1, ultimaColumna).setHorizontalAlignment('center');

  for (var col = 4; col <= ultimaColumna; col++) {
    if (col < ultimaColumna) {
      var actual = obtenerAnioMayorV1_(filas[3][col - 1]);
      var siguiente = obtenerAnioMayorV1_(filas[3][col]);
      if (actual && siguiente && actual !== siguiente) {
        mayor.getRange(1, col, filas.length, 1)
          .setBorder(null, null, null, true, null, null, '#666666', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      }
    }
  }
}


/**
 * Crea o corrige el botón de actualización del Mayor.
 * Conserva una sola imagen y la vincula con actualizarMayorV1.
 */
function asegurarBotonActualizarMayorV1_(ss) {
  ss = ss || SpreadsheetApp.getActive();
  var hoja = ss.getSheetByName('Mayor');
  if (!hoja) return;

  var titulo = 'RegOps v1 - Actualizar Mayor';
  var existentes = hoja.getImages().filter(function(imagen) {
    var tituloActual = '';
    var scriptActual = '';

    try { tituloActual = imagen.getAltTextTitle(); } catch (err) {}
    try { scriptActual = imagen.getScript(); } catch (err) {}

    return tituloActual === titulo || scriptActual === 'actualizarMayorV1';
  });

  if (existentes.length > 0) {
    var botonExistente = existentes[0];
    existentes.slice(1).forEach(function(imagen) { imagen.remove(); });

    botonExistente
      .setAltTextTitle(titulo)
      .setAltTextDescription('Actualiza el Libro Mayor de RegOps v1.')
      .setAnchorCell(hoja.getRange('A1'))
      .setAnchorCellXOffset(0)
      .setAnchorCellYOffset(0)
      .setWidth(125)
      .setHeight(27)
      .assignScript('actualizarMayorV1');

    return;
  }

  var pngBase64 = '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';
  var blob = Utilities.newBlob(
    Utilities.base64Decode(pngBase64),
    'image/png',
    'Actualizar-Mayor-v1.png'
  );

  hoja.insertImage(blob, 1, 1, 4, 2)
    .setAltTextTitle(titulo)
    .setAltTextDescription('Actualiza el Libro Mayor de RegOps v1.')
    .setAnchorCell(hoja.getRange('A1'))
    .setAnchorCellXOffset(0)
    .setAnchorCellYOffset(0)
    .setWidth(125)
    .setHeight(27)
    .assignScript('actualizarMayorV1');
}

function instalarBotonActualizarMayorV1() {
  asegurarBotonActualizarMayorV1_(SpreadsheetApp.getActive());
}
