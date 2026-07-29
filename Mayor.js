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
  meses.forEach(function(mes, indice) {
    filaMeses[3 + indice] = formatearMesMayorV1_(mes.valor);
  });
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

function aplicarFormatoMayorV1_(mayor, filas, tipos, ultimaColumna, cantidadMeses) {
  var rango = mayor.getRange(1, 1, filas.length, ultimaColumna);
  rango.setFontFamily('Nunito').setFontSize(8).setVerticalAlignment('middle');
  mayor.setFrozenRows(6);
  mayor.setFrozenColumns(3);
  mayor.setHiddenGridlines(true);

  mayor.setColumnWidth(1, 270);
  mayor.setColumnWidth(2, 105);
  mayor.setColumnWidth(3, 115);
  if (cantidadMeses > 0) {
    mayor.setColumnWidths(4, cantidadMeses, 88);
  }

  rango.setNumberFormat('#,##0;[Red](#,##0);-');
  mayor.getRange(2, 1, 2, 1).setNumberFormat('@');
  mayor.getRange(2, 3, 2, 1).setNumberFormat('0.00');

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
      var actual = String(filas[1][col - 1] || '').substring(0, 4);
      var siguiente = String(filas[1][col] || '').substring(0, 4);
      if (actual && siguiente && actual !== siguiente) {
        mayor.getRange(1, col, filas.length, 1)
          .setBorder(null, null, null, true, null, null, '#666666', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      }
    }
  }
}
