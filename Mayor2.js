/**
 * Mayor2: misma presentación que Mayor, con cálculo independiente.
 * Los movimientos se obtienen exclusivamente de Diario!A:F mediante SUMIFS.
 * Mayor se usa sólo como plantilla visual.
 */
function crearMayor2V1() {
  var ss = SpreadsheetApp.getActive();
  var diario = ss.getSheetByName('Diario');
  var mayor = ss.getSheetByName('Mayor');
  var definicion = ss.getSheetByName('CtasDefinicion');
  if (!diario || !definicion) {
    throw new Error('Se requieren las hojas Diario y CtasDefinicion.');
  }

  var mayor2 = ss.getSheetByName('Mayor2');
  if (!mayor2) {
    if (!mayor) {
      throw new Error('No existe Mayor2 ni Mayor como plantilla inicial.');
    }
    mayor2 = ss.insertSheet('Mayor2', mayor.getIndex() + 1);
    var filasPlantilla = mayor.getLastRow();
    var columnasPlantilla = mayor.getLastColumn();
    if (mayor2.getMaxRows() < filasPlantilla) {
      mayor2.insertRowsAfter(
        mayor2.getMaxRows(), filasPlantilla - mayor2.getMaxRows()
      );
    }
    if (mayor2.getMaxColumns() < columnasPlantilla) {
      mayor2.insertColumnsAfter(
        mayor2.getMaxColumns(), columnasPlantilla - mayor2.getMaxColumns()
      );
    }
    mayor.getRange(1, 1, filasPlantilla, columnasPlantilla)
      .copyTo(mayor2.getRange(1, 1, filasPlantilla, columnasPlantilla));
  }

  var filas = mayor2.getLastRow();
  var columnas = mayor2.getLastColumn();
  mayor2.getRange(1, 1, filas, columnas).breakApart();

  var meses = obtenerMesesMayor2V1_(diario);
  var cantidadMeses = Math.min(meses.length, Math.max(columnas - 3, 0));
  meses = meses.slice(0, cantidadMeses);
  var datos = mayor2.getRange(1, 1, filas, columnas).getValues();

  var tasaArs = numeroMayor2V1_(diario.getRange('L2').getValue(), 1);
  var tasaEuroArs = numeroMayor2V1_(diario.getRange('L1').getValue(), tasaArs);
  var tasaEuroUsd = tasaArs ? tasaEuroArs / tasaArs : 1;
  var tasaBrlUsd = 0.20;

  datos[0][1] = tasaEuroUsd;
  datos[1][1] = tasaBrlUsd;
  datos[2][1] = tasaArs;
  meses.forEach(function(mes, i) {
    datos[2][3 + i] = mes.tasa || tasaArs;
    datos[3][3 + i] = mes.anio;
    datos[4][3 + i] = abreviarMesMayor2V1_(mes.mes);
  });

  var filasDef = definicion.getRange(
    3, 1, Math.max(definicion.getLastRow() - 2, 1), 4
  ).getValues();
  var cuentas = [];
  filasDef.forEach(function(fila) {
    var nombre = String(fila[0] || '').trim();
    if (!nombre) return;
    cuentas.push({
      nombre: nombre,
      nombreVisible: nombreMostrarMayor2V1_(nombre),
      grupo: String(fila[2] || '').trim().toUpperCase(),
      estado: Number(fila[3]) || 0,
      filasMayor2: []
    });
  });

  var cuentaPorVisible = {};
  cuentas.forEach(function(cuenta) {
    cuentaPorVisible[normalizarNombreMayor2V1_(cuenta.nombreVisible)] = cuenta;
  });
  datos.forEach(function(fila, indice) {
    var cuenta = cuentaPorVisible[normalizarNombreMayor2V1_(fila[0])];
    if (cuenta) cuenta.filasMayor2.push(indice + 1);
  });
  cuentas.forEach(function(cuenta) {
    cuenta.filaCanonica = cuenta.filasMayor2.length ? cuenta.filasMayor2[0] : 0;
  });

  // Cuentas: cada período suma directamente Diario!D y Diario!F.
  cuentas.forEach(function(cuenta) {
    cuenta.filasMayor2.forEach(function(numeroFila) {
      var indiceFila = numeroFila - 1;
      meses.forEach(function(mes, indiceMes) {
        datos[indiceFila][3 + indiceMes] =
          formulaSumifsCuentaMayor2V1_(cuenta.nombre, mes);
      });
      datos[indiceFila][2] = cantidadMeses
        ? '=SUM(D' + numeroFila + ':' +
          columnaA1Mayor2V1_(3 + cantidadMeses) + numeroFila + ')'
        : '=0';
      datos[indiceFila][1] =
        formulaConversionActualMayor2V1_(cuenta.nombre, numeroFila);
    });
  });

  var resumenes = definicionesResumenMayor2V1_();
  var filaResumen = {};
  resumenes.forEach(function(resumen) {
    var fila = buscarFilaEtiquetaMayor2V1_(datos, resumen.nombre);
    if (!fila) return;
    filaResumen[resumen.nombre] = fila;
    var filasCuenta = cuentas.filter(function(cuenta) {
      return cuenta.filaCanonica && resumen.coincide(cuenta.nombre);
    }).map(function(cuenta) {
      return cuenta.filaCanonica;
    });
    datos[fila - 1][1] = formulaSumaCeldasMayor2V1_(filasCuenta, 2);
    datos[fila - 1][2] = formulaSumaCeldasMayor2V1_(filasCuenta, 3);
    meses.forEach(function(_, indiceMes) {
      datos[fila - 1][3 + indiceMes] =
        formulaSumaCeldasMayor2V1_(filasCuenta, 4 + indiceMes);
    });
  });

  // Patrimonio actual e histórico calculado desde las filas de cuentas.
  var filaPatrimonio = buscarFilaEtiquetaMayor2V1_(datos, 'PATRIMONIO NETO');
  var cuentasPatrimonio = cuentas.filter(function(cuenta) {
    return cuenta.filaCanonica &&
      /^(CJ|CC|INV|INVP)\b/i.test(cuenta.nombre) &&
      !/^Retiro\b/i.test(cuenta.nombre);
  });
  if (filaPatrimonio) {
    datos[filaPatrimonio - 1][1] = formulaSumaCeldasMayor2V1_(
      cuentasPatrimonio.map(function(cuenta) { return cuenta.filaCanonica; }), 2
    );
    meses.forEach(function(_, indiceMes) {
      datos[filaPatrimonio - 1][3 + indiceMes] =
        formulaPatrimonioPeriodoMayor2V1_(cuentasPatrimonio, 4 + indiceMes);
    });
  }

  // Retiro total: movimientos mensuales directos, convertidos con la tasa de cada mes.
  var filaRetiro = buscarFilaPrefijoMayor2V1_(datos, 'Retiro total');
  var retiros = cuentas.filter(function(cuenta) {
    return /^Retiro\b/i.test(cuenta.nombre) && cuenta.filaCanonica;
  });
  if (filaRetiro) {
    var terminosRetiro = [];
    retiros.forEach(function(cuenta) {
      meses.forEach(function(_, indiceMes) {
        var columna = columnaA1Mayor2V1_(4 + indiceMes);
        terminosRetiro.push(
          terminoConversionPeriodoMayor2V1_(
            cuenta.nombre, columna + cuenta.filaCanonica, columna + '$3'
          )
        );
      });
    });
    datos[filaRetiro - 1][1] = terminosRetiro.length
      ? '=SUM(' + terminosRetiro.join(',') + ')'
      : '=0';
    datos[filaRetiro - 1][0] =
      '="Retiro total USDk "&TEXT(B' + filaRetiro +
      ',"#,##0")&" y promedio USDk "&TEXT(B' + filaRetiro +
      '/' + Math.max(cantidadMeses, 1) +
      '/3,"#,##0")&"/mes/persona"';
  }

  // Morosos y vínculo al detalle inferior.
  var filasMorosos = buscarTodasFilasEtiquetaMayor2V1_(datos, 'MOROSOS');
  var filasEstado2 = cuentas.filter(function(cuenta) {
    return cuenta.estado === 2 && cuenta.filaCanonica;
  }).map(function(cuenta) {
    return cuenta.filaCanonica;
  });
  var formulaMorosos = formulaSumaCeldasMayor2V1_(filasEstado2, 2);
  if (filasMorosos.length) {
    datos[filasMorosos[0] - 1][1] = filasMorosos.length > 1
      ? '=HYPERLINK("#gid=' + mayor2.getSheetId() +
        '&range=A' + filasMorosos[filasMorosos.length - 1] +
        '",TEXT(' + formulaMorosos.substring(1) + ',"#,##0"))'
      : formulaMorosos;
  }
  if (filasMorosos.length > 1) {
    datos[filasMorosos[filasMorosos.length - 1] - 1][1] = formulaMorosos;
  }

  var filaCobertura = buscarFilaAlternativasMayor2V1_(
    datos, ['A CUBRIR', 'EXCEDENTE']
  );
  if (filaCobertura && filaPatrimonio) {
    var filaInv = filaResumen['INV USD (terceros)'];
    var filaInvp = filaResumen['INVP USD (TT)'];
    if (filaInv && filaInvp) {
      datos[filaCobertura - 1][1] =
        '=B' + filaInv + '-B' + filaInvp + '+B' + filaPatrimonio;
      datos[filaCobertura - 1][0] =
        '=IF(B' + filaCobertura + '>0,"EXCEDENTE","A CUBRIR")';
    }
  }

  mayor2.getRange(1, 1, filas, columnas).setValues(datos);
  combinarAniosMayor2V1_(mayor2, meses);
  mayor2.setFrozenRows(6);
  mayor2.setFrozenColumns(3);
  mayor2.setHiddenGridlines(false);
  mayor2.setColumnWidth(1, 240);
  mayor2.setColumnWidth(2, 85);
  mayor2.setColumnWidth(3, 100);
  if (cantidadMeses) mayor2.setColumnWidths(4, cantidadMeses, 58);
  aplicarReglasCoberturaMayor2V1_(mayor2, filas, columnas);

  SpreadsheetApp.flush();
  ss.setActiveSheet(mayor2);
  ss.toast(
    'Mayor2 actualizado directamente desde Diario con SUMIFS.',
    'RegOps v1',
    6
  );
}

function actualizarMayor2V1() {
  crearMayor2V1();
}

function obtenerMesesMayor2V1_(diario) {
  var ultimaFila = diario.getLastRow();
  var mapa = {};
  var tamanoBloque = 3000;

  // El Diario puede tener decenas de miles de movimientos. Leer la columna
  // completa en una sola llamada supera el límite de tamaño de Apps Script.
  for (var filaInicial = 6;
    filaInicial <= ultimaFila;
    filaInicial += tamanoBloque) {
    var cantidad = Math.min(tamanoBloque, ultimaFila - filaInicial + 1);
    var fechas = diario.getRange(filaInicial, 1, cantidad, 1).getValues();
    fechas.forEach(function(fila) {
      var fecha = fila[0] instanceof Date ? fila[0] : new Date(fila[0]);
      if (isNaN(fecha.getTime())) return;
      var anio = fecha.getFullYear();
      var mes = fecha.getMonth() + 1;
      var clave = anio + '-' + (mes < 10 ? '0' : '') + mes;
      mapa[clave] = {clave: clave, anio: anio, mes: mes};
    });
  }

  // Recupera sólo los valores históricos de tasa existentes como valores iniciales.
  // Ninguna fórmula de Mayor2 queda vinculada a la tabla auxiliar.
  var tasas = {};
  var ancho = Math.max(diario.getLastColumn() - 12, 0);
  if (ancho) {
    var encabezados = diario.getRange(5, 13, 1, ancho).getValues()[0];
    var valoresTasa = diario.getRange(2, 13, 1, ancho).getValues()[0];
    encabezados.forEach(function(valor, indice) {
      var fecha = valor instanceof Date ? valor : new Date(valor);
      if (isNaN(fecha.getTime())) return;
      var mes = fecha.getMonth() + 1;
      var clave = fecha.getFullYear() + '-' + (mes < 10 ? '0' : '') + mes;
      tasas[clave] = Number(valoresTasa[indice]) || 0;
    });
  }

  return Object.keys(mapa).sort().reverse().map(function(clave) {
    mapa[clave].tasa = tasas[clave] || 0;
    return mapa[clave];
  });
}

function formulaSumifsCuentaMayor2V1_(cuenta, mes) {
  var nombre = String(cuenta).replace(/"/g, '""');
  var inicio = 'DATE(' + mes.anio + ',' + mes.mes + ',1)';
  return '=SUMIFS(Diario!$D$6:$D,Diario!$A$6:$A,">="&' + inicio +
    ',Diario!$A$6:$A,"<"&EDATE(' + inicio +
    ',1),Diario!$C$6:$C,"' + nombre +
    '")+SUMIFS(Diario!$F$6:$F,Diario!$A$6:$A,">="&' + inicio +
    ',Diario!$A$6:$A,"<"&EDATE(' + inicio +
    ',1),Diario!$E$6:$E,"' + nombre + '")';
}

function formulaConversionActualMayor2V1_(nombre, fila) {
  if (/\bARS\b/i.test(nombre)) return '=C' + fila + '/$B$3/1000';
  if (/\b(EURO|EUR)\b/i.test(nombre)) return '=C' + fila + '*$B$1/1000';
  if (/\b(BRL|REAL|REALES)\b/i.test(nombre)) return '=C' + fila + '*$B$2/1000';
  return '=C' + fila + '/1000';
}

function terminoConversionPeriodoMayor2V1_(nombre, celda, celdaTasa) {
  if (/\bARS\b/i.test(nombre)) return celda + '/' + celdaTasa + '/1000';
  if (/\b(EURO|EUR)\b/i.test(nombre)) return celda + '*$B$1/1000';
  if (/\b(BRL|REAL|REALES)\b/i.test(nombre)) return celda + '*$B$2/1000';
  return celda + '/1000';
}

function formulaPatrimonioPeriodoMayor2V1_(cuentas, columna) {
  var letra = columnaA1Mayor2V1_(columna);
  var terminos = cuentas.map(function(cuenta) {
    return terminoConversionPeriodoMayor2V1_(
      cuenta.nombre, letra + cuenta.filaCanonica, letra + '$3'
    );
  });
  return terminos.length ? '=SUM(' + terminos.join(',') + ')' : '=0';
}

function definicionesResumenMayor2V1_() {
  return [
    {nombre: 'CJ ARS (USD)', coincide: function(n) {
      return /^CJ\b/i.test(n) && /\bARS\b/i.test(n);
    }},
    {nombre: 'CJ USD', coincide: function(n) {
      return /^CJ\b/i.test(n) && /\bUSD\b/i.test(n) && !/\bUSDT\b/i.test(n);
    }},
    {nombre: 'CJ EURO (USD)', coincide: function(n) {
      return /^CJ\b/i.test(n) && /\b(EURO|EUR)\b/i.test(n);
    }},
    {nombre: 'CJ BRL (USD)', coincide: function(n) {
      return /^CJ\b/i.test(n) && /\b(BRL|REAL|REALES)\b/i.test(n);
    }},
    {nombre: 'CC ARS (USD)', coincide: function(n) {
      return /^CC\b/i.test(n) && /\bARS\b/i.test(n);
    }},
    {nombre: 'CC USD', coincide: function(n) {
      return /^CC\b/i.test(n) && /\bUSD\b/i.test(n);
    }},
    {nombre: 'CC BRL (USD)', coincide: function(n) {
      return /^CC\b/i.test(n) && /\b(BRL|REAL|REALES)\b/i.test(n);
    }},
    {nombre: 'INV USD (terceros)', coincide: function(n) {
      return /^INV\s/i.test(n) && !/^INVP\b/i.test(n);
    }},
    {nombre: 'INVP USD (TT)', coincide: function(n) {
      return /^INVP\b/i.test(n);
    }}
  ];
}

function aplicarReglasCoberturaMayor2V1_(hoja, filas, columnas) {
  var reglas = hoja.getConditionalFormatRules().filter(function(regla) {
    try {
      var condicion = regla.getBooleanCondition();
      if (!condicion) return true;
      var valores = condicion.getCriteriaValues();
      var formula = valores && valores.length ? String(valores[0] || '') : '';
      return formula.indexOf('$A1="EXCEDENTE"') === -1 &&
        formula.indexOf('$A1="A CUBRIR"') === -1;
    } catch (err) {
      return true;
    }
  });
  var rango = hoja.getRange(1, 1, filas, columnas);
  reglas.unshift(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$A1="EXCEDENTE"')
      .setBackground('#d9ead3').setRanges([rango]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$A1="A CUBRIR"')
      .setBackground('#f4cccc').setRanges([rango]).build()
  );
  hoja.setConditionalFormatRules(reglas);
}

function numeroMayor2V1_(valor, alternativa) {
  var numero = Number(valor);
  return isFinite(numero) && numero !== 0 ? numero : alternativa;
}

function abreviarMesMayor2V1_(mes) {
  return ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun',
    'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'][mes - 1] || '';
}

function columnaA1Mayor2V1_(numeroColumna) {
  var resultado = '';
  while (numeroColumna > 0) {
    var resto = (numeroColumna - 1) % 26;
    resultado = String.fromCharCode(65 + resto) + resultado;
    numeroColumna = Math.floor((numeroColumna - 1) / 26);
  }
  return resultado;
}

function nombreMostrarMayor2V1_(nombre) {
  var texto = String(nombre || '').trim();
  return /^Retiro\b/i.test(texto)
    ? texto.replace(/\bARS\b/gi, '').replace(/\s+/g, ' ').trim()
    : texto;
}

function normalizarNombreMayor2V1_(nombre) {
  return String(nombre || '').replace(/\u00a0/g, ' ')
    .replace(/\s+/g, ' ').trim();
}

function buscarFilaEtiquetaMayor2V1_(datos, etiqueta) {
  var buscada = normalizarNombreMayor2V1_(etiqueta);
  for (var i = 0; i < datos.length; i++) {
    if (normalizarNombreMayor2V1_(datos[i][0]) === buscada) return i + 1;
  }
  return 0;
}

function buscarTodasFilasEtiquetaMayor2V1_(datos, etiqueta) {
  var resultado = [];
  var buscada = normalizarNombreMayor2V1_(etiqueta);
  datos.forEach(function(fila, i) {
    if (normalizarNombreMayor2V1_(fila[0]) === buscada) resultado.push(i + 1);
  });
  return resultado;
}

function buscarFilaPrefijoMayor2V1_(datos, prefijo) {
  var buscado = normalizarNombreMayor2V1_(prefijo).toLowerCase();
  for (var i = 0; i < datos.length; i++) {
    if (normalizarNombreMayor2V1_(datos[i][0]).toLowerCase()
      .indexOf(buscado) === 0) return i + 1;
  }
  return 0;
}

function buscarFilaAlternativasMayor2V1_(datos, alternativas) {
  for (var i = 0; i < alternativas.length; i++) {
    var fila = buscarFilaEtiquetaMayor2V1_(datos, alternativas[i]);
    if (fila) return fila;
  }
  return 0;
}

function formulaSumaCeldasMayor2V1_(filas, columna) {
  if (!filas.length) return '=0';
  var letra = columnaA1Mayor2V1_(columna);
  return '=SUM(' + filas.map(function(fila) {
    return letra + fila;
  }).join(',') + ')';
}

function combinarAniosMayor2V1_(hoja, meses) {
  if (!meses.length) return;
  var inicio = 0;
  while (inicio < meses.length) {
    var fin = inicio;
    while (fin + 1 < meses.length &&
      meses[fin + 1].anio === meses[inicio].anio) fin++;
    var rango = hoja.getRange(4, 4 + inicio, 1, fin - inicio + 1);
    if (fin > inicio) rango.merge();
    rango.setValue(meses[inicio].anio).setHorizontalAlignment('center');
    inicio = fin + 1;
  }
}
