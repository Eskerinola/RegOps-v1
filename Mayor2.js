/**
 * Crea una versión experimental del Mayor alimentada con fórmulas nativas.
 * La estructura se toma del Mayor actual, pero los importes se vinculan
 * directamente con la tabla rápida ubicada en Diario!K:...
 */
function crearMayor2V1() {
  var ss = SpreadsheetApp.getActive();
  var diario = ss.getSheetByName('Diario');
  var mayor = ss.getSheetByName('Mayor');
  var definicion = ss.getSheetByName('CtasDefinicion');

  if (!diario || !mayor || !definicion) {
    throw new Error('Se requieren las hojas Diario, Mayor y CtasDefinicion.');
  }

  var mayor2 = ss.getSheetByName('Mayor2');
  if (!mayor2) {
    mayor2 = ss.insertSheet('Mayor2', mayor.getIndex() + 1);
  }

  var ultimaFila = mayor.getLastRow();
  var ultimaColumna = mayor.getLastColumn();
  var rangoDestinoCompleto = mayor2.getRange(
    1, 1, mayor2.getMaxRows(), mayor2.getMaxColumns()
  );

  rangoDestinoCompleto.breakApart();
  mayor2.clear();

  if (mayor2.getMaxRows() < ultimaFila) {
    mayor2.insertRowsAfter(
      mayor2.getMaxRows(), ultimaFila - mayor2.getMaxRows()
    );
  }
  if (mayor2.getMaxColumns() < ultimaColumna) {
    mayor2.insertColumnsAfter(
      mayor2.getMaxColumns(), ultimaColumna - mayor2.getMaxColumns()
    );
  }

  mayor.getRange(1, 1, ultimaFila, ultimaColumna)
    .copyTo(mayor2.getRange(1, 1, ultimaFila, ultimaColumna));

  mayor2.getRange(1, 1, ultimaFila, ultimaColumna).breakApart();

  var meses = obtenerMesesMayor2V1_(diario);
  var cantidadMeses = meses.length;
  var datos = mayor2.getRange(1, 1, ultimaFila, ultimaColumna).getValues();

  var filasDefinicion = definicion
    .getRange(3, 1, Math.max(definicion.getLastRow() - 2, 1), 4)
    .getValues();

  var cuentas = [];
  filasDefinicion.forEach(function(fila, indice) {
    var nombre = String(fila[0] || '').trim();
    if (!nombre) return;

    cuentas.push({
      nombre: nombre,
      nombreVisible: nombreMostrarMayor2V1_(nombre),
      grupo: String(fila[2] || '').trim().toUpperCase(),
      estado: Number(fila[3]) || 0,
      filaDiario: 15 + indice,
      filasMayor2: []
    });
  });

  var cuentaPorNombreVisible = {};
  cuentas.forEach(function(cuenta) {
    cuentaPorNombreVisible[
      normalizarNombreMayor2V1_(cuenta.nombreVisible)
    ] = cuenta;
  });

  // Localiza todas las apariciones de cuentas, incluida la sección MOROSOS.
  for (var indiceFila = 0; indiceFila < datos.length; indiceFila++) {
    var etiqueta = normalizarNombreMayor2V1_(datos[indiceFila][0]);
    var cuenta = cuentaPorNombreVisible[etiqueta];
    if (cuenta) cuenta.filasMayor2.push(indiceFila + 1);
  }

  // La primera aparición es la fila canónica usada para los totales.
  cuentas.forEach(function(cuenta) {
    cuenta.filaCanonica = cuenta.filasMayor2.length
      ? cuenta.filasMayor2[0]
      : 0;
  });

  // Tasas actuales.
  datos[0][1] = '=Diario!L1/Diario!L2';
  datos[1][1] = '=0.20';
  datos[2][1] = '=Diario!L2';

  // Cada cuenta se vincula con su fila rápida del Diario.
  cuentas.forEach(function(cuenta) {
    cuenta.filasMayor2.forEach(function(filaMayor2) {
      var fila = filaMayor2 - 1;
      datos[fila][2] = '=Diario!L' + cuenta.filaDiario;

      if (/\bARS\b/i.test(cuenta.nombre)) {
        datos[fila][1] = '=C' + filaMayor2 + '/$B$3/1000';
      } else if (/\b(EURO|EUR)\b/i.test(cuenta.nombre)) {
        datos[fila][1] = '=C' + filaMayor2 + '*$B$1/1000';
      } else if (/\b(BRL|REAL|REALES)\b/i.test(cuenta.nombre)) {
        datos[fila][1] = '=C' + filaMayor2 + '*$B$2/1000';
      } else {
        datos[fila][1] = '=C' + filaMayor2 + '/1000';
      }

      meses.forEach(function(mes, indiceMes) {
        datos[fila][3 + indiceMes] =
          '=Diario!' + mes.columna + cuenta.filaDiario;
      });
    });
  });

  var definicionesResumen = [
    {
      nombre: 'CJ ARS (USD)',
      coincide: function(n) {
        return /^CJ\b/i.test(n) && /\bARS\b/i.test(n);
      }
    },
    {
      nombre: 'CJ USD',
      coincide: function(n) {
        return /^CJ\b/i.test(n) &&
          /\bUSD\b/i.test(n) &&
          !/\bUSDT\b/i.test(n);
      }
    },
    {
      nombre: 'CJ EURO (USD)',
      coincide: function(n) {
        return /^CJ\b/i.test(n) && /\b(EURO|EUR)\b/i.test(n);
      }
    },
    {
      nombre: 'CJ BRL (USD)',
      coincide: function(n) {
        return /^CJ\b/i.test(n) && /\b(BRL|REAL|REALES)\b/i.test(n);
      }
    },
    {
      nombre: 'CC ARS (USD)',
      coincide: function(n) {
        return /^CC\b/i.test(n) && /\bARS\b/i.test(n);
      }
    },
    {
      nombre: 'CC USD',
      coincide: function(n) {
        return /^CC\b/i.test(n) && /\bUSD\b/i.test(n);
      }
    },
    {
      nombre: 'CC BRL (USD)',
      coincide: function(n) {
        return /^CC\b/i.test(n) && /\b(BRL|REAL|REALES)\b/i.test(n);
      }
    },
    {
      nombre: 'INV USD (terceros)',
      coincide: function(n) {
        return /^INV\s/i.test(n) && !/^INVP\b/i.test(n);
      }
    },
    {
      nombre: 'INVP USD (TT)',
      coincide: function(n) {
        return /^INVP\b/i.test(n);
      }
    }
  ];

  var filaResumenPorNombre = {};
  definicionesResumen.forEach(function(resumen) {
    var filaResumen = buscarFilaEtiquetaMayor2V1_(datos, resumen.nombre);
    if (!filaResumen) return;

    filaResumenPorNombre[resumen.nombre] = filaResumen;
    var filasCuentas = cuentas
      .filter(function(cuenta) {
        return cuenta.filaCanonica && resumen.coincide(cuenta.nombre);
      })
      .map(function(cuenta) { return cuenta.filaCanonica; });

    datos[filaResumen - 1][1] =
      formulaSumaCeldasMayor2V1_(filasCuentas, 2);
    datos[filaResumen - 1][2] =
      formulaSumaCeldasMayor2V1_(filasCuentas, 3);

    meses.forEach(function(_, indiceMes) {
      datos[filaResumen - 1][3 + indiceMes] =
        formulaSumaCeldasMayor2V1_(filasCuentas, 4 + indiceMes);
    });
  });

  // Patrimonio neto y evolución histórica.
  var filaPatrimonio = buscarFilaEtiquetaMayor2V1_(
    datos, 'PATRIMONIO NETO'
  );
  if (filaPatrimonio) {
    var filasPatrimonio = cuentas
      .filter(function(cuenta) {
        return cuenta.filaCanonica &&
          /^(CJ|CC|INV)\b/i.test(cuenta.nombre) &&
          !/^Retiro\b/i.test(cuenta.nombre);
      })
      .map(function(cuenta) { return cuenta.filaCanonica; });

    datos[filaPatrimonio - 1][1] =
      formulaSumaCeldasMayor2V1_(filasPatrimonio, 2);

    meses.forEach(function(mes, indiceMes) {
      datos[filaPatrimonio - 1][3 + indiceMes] =
        '=Diario!' + mes.columna + '4';
    });
  }

  // Retiro total: cada mes se convierte con su propia tasa histórica.
  var filaRetiro = buscarFilaPrefijoMayor2V1_(datos, 'Retiro total');
  if (filaRetiro) {
    var retiros = cuentas.filter(function(cuenta) {
      return /^Retiro\b/i.test(cuenta.nombre);
    });
    var terminosRetiro = [];

    retiros.forEach(function(cuenta) {
      meses.forEach(function(mes) {
        terminosRetiro.push(
          'Diario!' + mes.columna + cuenta.filaDiario +
          '/Diario!' + mes.columna + '$2/1000'
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

  // MOROSOS: suma cuentas Estado 2 y conserva el enlace al detalle.
  var filasMorosos = buscarTodasFilasEtiquetaMayor2V1_(datos, 'MOROSOS');
  var filasEstado2 = cuentas
    .filter(function(cuenta) {
      return cuenta.estado === 2 && cuenta.filaCanonica;
    })
    .map(function(cuenta) { return cuenta.filaCanonica; });
  var formulaMorosos = formulaSumaCeldasMayor2V1_(filasEstado2, 2);

  if (filasMorosos.length) {
    datos[filasMorosos[0] - 1][1] = filasMorosos.length > 1
      ? '=HYPERLINK("#gid=' + mayor2.getSheetId() +
        '&range=A' + filasMorosos[filasMorosos.length - 1] +
        '",TEXT(' + formulaMorosos.substring(1) + ',"#,##0"))'
      : formulaMorosos;
  }
  if (filasMorosos.length > 1) {
    datos[filasMorosos[filasMorosos.length - 1] - 1][1] =
      formulaMorosos;
  }

  // A CUBRIR / EXCEDENTE.
  var filaCobertura = buscarFilaAlternativasMayor2V1_(
    datos, ['A CUBRIR', 'EXCEDENTE']
  );
  if (filaCobertura && filaPatrimonio) {
    var filaInv = filaResumenPorNombre['INV USD (terceros)'];
    var filaInvp = filaResumenPorNombre['INVP USD (TT)'];
    if (filaInv && filaInvp) {
      datos[filaCobertura - 1][1] =
        '=B' + filaInv + '-B' + filaInvp + '+B' + filaPatrimonio;
      datos[filaCobertura - 1][0] =
        '=IF(B' + filaCobertura +
        '>0,"EXCEDENTE","A CUBRIR")';
    }
  }

  mayor2.getRange(1, 1, ultimaFila, ultimaColumna).setValues(datos);

  // Restablece los encabezados de años combinados.
  combinarAniosMayor2V1_(mayor2, meses);

  mayor2.setName('Mayor2');
  mayor2.setFrozenRows(6);
  mayor2.setFrozenColumns(3);
  mayor2.setHiddenGridlines(false);
  mayor2.setColumnWidth(1, 240);
  mayor2.setColumnWidth(2, 85);
  mayor2.setColumnWidth(3, 100);
  if (cantidadMeses) {
    mayor2.setColumnWidths(4, cantidadMeses, 58);
  }

  // El color de A CUBRIR / EXCEDENTE cambia automáticamente.
  var reglas = mayor2.getConditionalFormatRules().filter(function(regla) {
    try {
      var condicion = regla.getBooleanCondition();
      if (!condicion) return true;
      var valores = condicion.getCriteriaValues();
      var formula = valores && valores.length
        ? String(valores[0] || '')
        : '';
      return formula.indexOf('$A1="EXCEDENTE"') === -1 &&
        formula.indexOf('$A1="A CUBRIR"') === -1;
    } catch (err) {
      return true;
    }
  });

  var rangoCuerpo = mayor2.getRange(
    1, 1, ultimaFila, ultimaColumna
  );
  reglas.unshift(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$A1="EXCEDENTE"')
      .setBackground('#d9ead3')
      .setRanges([rangoCuerpo])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$A1="A CUBRIR"')
      .setBackground('#f4cccc')
      .setRanges([rangoCuerpo])
      .build()
  );
  mayor2.setConditionalFormatRules(reglas);

  SpreadsheetApp.flush();
  ss.setActiveSheet(mayor2);
  ss.toast(
    'Mayor2 creado. Desde ahora sus importes se actualizan con fórmulas.',
    'RegOps v1',
    6
  );
}

function actualizarMayor2V1() {
  crearMayor2V1();
}

function obtenerMesesMayor2V1_(diario) {
  var ancho = Math.max(diario.getLastColumn() - 12, 1);
  var valores = diario.getRange(5, 13, 1, ancho).getValues()[0];
  var meses = [];

  for (var indice = 0; indice < valores.length; indice++) {
    if (valores[indice] === '' || valores[indice] === null) break;
    meses.push({
      valor: valores[indice],
      columna: columnaA1Mayor2V1_(13 + indice)
    });
  }

  return meses.reverse();
}

function columnaA1Mayor2V1_(numeroColumna) {
  var resultado = '';
  var numero = numeroColumna;

  while (numero > 0) {
    var resto = (numero - 1) % 26;
    resultado = String.fromCharCode(65 + resto) + resultado;
    numero = Math.floor((numero - 1) / 26);
  }

  return resultado;
}

function nombreMostrarMayor2V1_(nombre) {
  var texto = String(nombre || '').trim();
  if (/^Retiro\b/i.test(texto)) {
    return texto.replace(/\bARS\b/gi, '').replace(/\s+/g, ' ').trim();
  }
  return texto;
}

function normalizarNombreMayor2V1_(nombre) {
  return String(nombre || '')
    .replace(/\u00a0/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function buscarFilaEtiquetaMayor2V1_(datos, etiqueta) {
  var buscada = normalizarNombreMayor2V1_(etiqueta);
  for (var indice = 0; indice < datos.length; indice++) {
    if (normalizarNombreMayor2V1_(datos[indice][0]) === buscada) {
      return indice + 1;
    }
  }
  return 0;
}

function buscarTodasFilasEtiquetaMayor2V1_(datos, etiqueta) {
  var resultado = [];
  var buscada = normalizarNombreMayor2V1_(etiqueta);

  datos.forEach(function(fila, indice) {
    if (normalizarNombreMayor2V1_(fila[0]) === buscada) {
      resultado.push(indice + 1);
    }
  });

  return resultado;
}

function buscarFilaPrefijoMayor2V1_(datos, prefijo) {
  var buscado = normalizarNombreMayor2V1_(prefijo).toLowerCase();
  for (var indice = 0; indice < datos.length; indice++) {
    var valor = normalizarNombreMayor2V1_(datos[indice][0]).toLowerCase();
    if (valor.indexOf(buscado) === 0) return indice + 1;
  }
  return 0;
}

function buscarFilaAlternativasMayor2V1_(datos, alternativas) {
  for (var indice = 0; indice < alternativas.length; indice++) {
    var fila = buscarFilaEtiquetaMayor2V1_(datos, alternativas[indice]);
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

  var anios = meses.map(function(mes) {
    var fecha = mes.valor instanceof Date
      ? mes.valor
      : new Date(mes.valor);
    return isNaN(fecha.getTime())
      ? String(mes.valor || '').match(/\d{4}/)[0]
      : String(fecha.getFullYear());
  });

  var inicio = 0;
  while (inicio < anios.length) {
    var fin = inicio;
    while (fin + 1 < anios.length && anios[fin + 1] === anios[inicio]) {
      fin++;
    }

    var rango = hoja.getRange(4, 4 + inicio, 1, fin - inicio + 1);
    if (fin > inicio) rango.merge();
    rango.setValue(anios[inicio]).setHorizontalAlignment('center');
    inicio = fin + 1;
  }
}
