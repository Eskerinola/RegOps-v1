/**
 * Mayor eficiente de RegOps v1.
 *
 * Lee Diario!A:F por bloques, agrupa en memoria y escribe valores terminados.
 * No depende de la antigua tabla auxiliar ubicada a la derecha del Diario.
 */

var REGOPS_MAYOR_V1 = {
  HOJA: 'Mayor',
  DIARIO: 'Diario',
  CUENTAS: 'CtasDefinicion',
  PRIMERA_FILA_DIARIO: 6,
  TAMANO_BLOQUE: 3000,
  PROPIEDAD_SUCIO: 'REGOPS_MAYOR_DESACTUALIZADO',
  PROPIEDAD_ACTUALIZADO: 'REGOPS_MAYOR_ULTIMA_ACTUALIZACION'
};


function actualizarMayorV1() {
  var ss = SpreadsheetApp.getActive();
  var diario = ss.getSheetByName(REGOPS_MAYOR_V1.DIARIO);
  var mayor = ss.getSheetByName(REGOPS_MAYOR_V1.HOJA);
  var definicion = ss.getSheetByName(REGOPS_MAYOR_V1.CUENTAS);

  if (!diario || !mayor || !definicion) {
    throw new Error('Se requieren las hojas Diario, Mayor y CtasDefinicion.');
  }

  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(1000)) return;

  try {
    var tasas = obtenerTasasMayorEficienteV1_(mayor);
    var plan = obtenerPlanMayorEficienteV1_(definicion);
    var movimientos = acumularDiarioMayorEficienteV1_(diario);
    var meses = Object.keys(movimientos.meses).sort().reverse();
    asegurarColumnasMayorEficienteV1_(mayor, 3 + meses.length);

    var filas = mayor.getLastRow();
    var columnas = mayor.getLastColumn();
    // Sólo la fila de años contiene combinaciones. No tocar toda la hoja:
    // hacerlo provocaba recalculaciones innecesarias en las fórmulas del Diario.
    if (columnas > 3) mayor.getRange(4, 4, 1, columnas - 3).breakApart();
    var datos = mayor.getRange(1, 1, filas, columnas).getValues();

    var tasaArsActual = tasas.actual.ars || 1;
    datos[0][0] = 'T/C Euro/USD';
    datos[0][1] = tasas.actual.euro;
    datos[1][0] = 'T/C BRL/USD';
    datos[1][1] = tasas.actual.brl;
    datos[2][0] = 'T/C ARS/USD';
    datos[2][1] = tasaArsActual;

    // Limpia encabezados y resultados mensuales anteriores.
    for (var c = 3; c < columnas; c++) {
      datos[2][c] = '';
      datos[3][c] = '';
      datos[4][c] = '';
    }

    meses.forEach(function(clave, indice) {
      var partes = clave.split('-');
      var anio = Number(partes[0]);
      var mes = Number(partes[1]);
      var columna = 3 + indice;
      datos[2][columna] = tasas.historicas[clave] || tasaArsActual;
      datos[3][columna] = anio;
      datos[4][columna] = abreviarMesMayorEficienteV1_(mes);
    });

    var cuentaPorVisible = {};
    plan.cuentas.forEach(function(cuenta) {
      cuenta.filas = [];
      cuentaPorVisible[
        normalizarMayorEficienteV1_(nombreVisibleMayorEficienteV1_(cuenta.nombre))
      ] = cuenta;
    });

    datos.forEach(function(fila, indice) {
      var cuenta = cuentaPorVisible[normalizarMayorEficienteV1_(fila[0])];
      if (cuenta) cuenta.filas.push(indice + 1);
    });

    plan.cuentas.forEach(function(cuenta) {
      cuenta.filaCanonica = cuenta.filas.length ? cuenta.filas[0] : 0;
      var porMes = movimientos.porCuenta[cuenta.nombre] || {};
      var acumuladoMoneda = 0;
      meses.forEach(function(clave) {
        acumuladoMoneda += Number(porMes[clave]) || 0;
      });
      cuenta.acumuladoMoneda = acumuladoMoneda;
      cuenta.acumuladoUsdK = convertirActualMayorEficienteV1_(
        cuenta.nombre,
        acumuladoMoneda,
        tasas.actual
      );

      cuenta.filas.forEach(function(numeroFila) {
        var indiceFila = numeroFila - 1;
        datos[indiceFila][1] = cuenta.acumuladoUsdK;
        datos[indiceFila][2] = acumuladoMoneda;

        for (var col = 3; col < columnas; col++) datos[indiceFila][col] = '';
        meses.forEach(function(clave, indiceMes) {
          datos[indiceFila][3 + indiceMes] = Number(porMes[clave]) || 0;
        });
      });
    });

    var resumenes = definicionesResumenMayorEficienteV1_();
    var filaResumen = {};
    resumenes.forEach(function(resumen) {
      var fila = buscarFilaMayorEficienteV1_(datos, resumen.nombre);
      if (!fila) return;

      filaResumen[resumen.nombre] = fila;
      var incluidas = plan.cuentas.filter(function(cuenta) {
        return cuenta.filaCanonica && resumen.coincide(cuenta.nombre);
      });

      datos[fila - 1][1] = sumarMayorEficienteV1_(incluidas, 'acumuladoUsdK');
      datos[fila - 1][2] = sumarMayorEficienteV1_(incluidas, 'acumuladoMoneda');

      for (var col = 3; col < columnas; col++) datos[fila - 1][col] = '';
      meses.forEach(function(clave, indiceMes) {
        datos[fila - 1][3 + indiceMes] = incluidas.reduce(function(total, cuenta) {
          var porMes = movimientos.porCuenta[cuenta.nombre] || {};
          return total + (Number(porMes[clave]) || 0);
        }, 0);
      });
    });

    var cuentasPatrimonio = plan.cuentas.filter(function(cuenta) {
      return cuenta.filaCanonica &&
        /^(CJ|CC|INV|INVP)\b/i.test(cuenta.nombre) &&
        !/^Retiro\b/i.test(cuenta.nombre);
    });
    var filaPatrimonio = buscarFilaMayorEficienteV1_(datos, 'PATRIMONIO NETO');
    var patrimonio = cuentasPatrimonio.reduce(function(total, cuenta) {
      return total + cuenta.acumuladoUsdK;
    }, 0);

    if (filaPatrimonio) {
      datos[filaPatrimonio - 1][1] = patrimonio;
      for (var colPn = 3; colPn < columnas; colPn++) {
        datos[filaPatrimonio - 1][colPn] = '';
      }
      meses.forEach(function(clave, indiceMes) {
        var tasaMes = {
          ars: tasas.historicas[clave] || tasaArsActual,
          euro: tasas.actual.euro,
          brl: tasas.actual.brl
        };
        datos[filaPatrimonio - 1][3 + indiceMes] =
          cuentasPatrimonio.reduce(function(total, cuenta) {
            var valor = Number(
              (movimientos.porCuenta[cuenta.nombre] || {})[clave]
            ) || 0;
            return total + convertirActualMayorEficienteV1_(
              cuenta.nombre, valor, tasaMes
            );
          }, 0);
      });
    }

    var retiros = plan.cuentas.filter(function(cuenta) {
      return cuenta.filaCanonica && /^Retiro\b/i.test(cuenta.nombre);
    });
    var retiroTotal = 0;
    meses.forEach(function(clave) {
      var tasaMes = {
        ars: tasas.historicas[clave] || tasaArsActual,
        euro: tasas.actual.euro,
        brl: tasas.actual.brl
      };
      retiros.forEach(function(cuenta) {
        var valor = Number(
          (movimientos.porCuenta[cuenta.nombre] || {})[clave]
        ) || 0;
        retiroTotal += convertirActualMayorEficienteV1_(
          cuenta.nombre, valor, tasaMes
        );
      });
    });

    var filaRetiro = buscarFilaPrefijoMayorEficienteV1_(datos, 'Retiro total');
    if (filaRetiro) {
      datos[filaRetiro - 1][0] =
        'Retiro total USDk ' + redondearMayorEficienteV1_(retiroTotal) +
        ' y promedio USDk ' +
        redondearMayorEficienteV1_(
          retiroTotal / Math.max(meses.length, 1) / 3
        ) +
        '/mes/persona';
      datos[filaRetiro - 1][1] = retiroTotal;
    }

    var morosos = plan.cuentas.filter(function(cuenta) {
      return cuenta.estado === 2 && cuenta.filaCanonica;
    });
    var totalMorosos = sumarMayorEficienteV1_(morosos, 'acumuladoUsdK');
    var filasMorosos = buscarTodasFilasMayorEficienteV1_(datos, 'MOROSOS');
    filasMorosos.forEach(function(fila) {
      datos[fila - 1][1] = totalMorosos;
    });

    var filaCobertura = buscarFilaAlternativasMayorEficienteV1_(
      datos, ['A CUBRIR', 'EXCEDENTE']
    );
    var filaInv = filaResumen['INV USD (terceros)'];
    var filaInvp = filaResumen['INVP USD (TT)'];
    if (filaCobertura && filaInv && filaInvp) {
      var cobertura =
        Number(datos[filaInv - 1][1] || 0) -
        Number(datos[filaInvp - 1][1] || 0) +
        patrimonio;
      datos[filaCobertura - 1][0] =
        cobertura > 0 ? 'EXCEDENTE' : 'A CUBRIR';
      datos[filaCobertura - 1][1] = cobertura;
    }

    // B1:B3 son entradas manuales de Tasas de Cambio. Se leen, pero no se
    // reescriben durante el cálculo. Así el Mayor del Diario no recalcula
    // todos sus SUMIFS cada vez que se pulsa Actualizar.
    if (filas > 3) {
      mayor.getRange(4, 1, filas - 3, columnas).setValues(datos.slice(3));
    }
    combinarAniosMayorEficienteV1_(mayor, meses);
    aplicarFormatoMayorEficienteV1_(mayor, filas, columnas, meses.length);

    var propiedades = PropertiesService.getDocumentProperties();
    propiedades.setProperty(REGOPS_MAYOR_V1.PROPIEDAD_SUCIO, '0');
    propiedades.setProperty(
      REGOPS_MAYOR_V1.PROPIEDAD_ACTUALIZADO,
      new Date().toISOString()
    );

    SpreadsheetApp.flush();
    ss.toast('Mayor actualizado.', 'RegOps v1', 4);
  } finally {
    lock.releaseLock();
  }
}


function actualizarMayorSiCorrespondeV1_() {
  var ss = SpreadsheetApp.getActive();
  var email = '';
  try {
    email = Session.getEffectiveUser().getEmail();
    if (ss.getOwner().getEmail() !== email) return;
  } catch (_) {
    return;
  }

  var propiedades = PropertiesService.getDocumentProperties();
  var sucio = propiedades.getProperty(REGOPS_MAYOR_V1.PROPIEDAD_SUCIO);
  var actualizado = propiedades.getProperty(
    REGOPS_MAYOR_V1.PROPIEDAD_ACTUALIZADO
  );
  if (sucio !== '0' || !actualizado) actualizarMayorV1();
}


function marcarMayorDesactualizadoV1_() {
  PropertiesService.getDocumentProperties()
    .setProperty(REGOPS_MAYOR_V1.PROPIEDAD_SUCIO, '1');
}


function obtenerPlanMayorEficienteV1_(hoja) {
  var ultimaFila = hoja.getRange(hoja.getMaxRows(), 1)
    .getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
  if (ultimaFila < 3) return {cuentas: []};

  var valores = hoja.getRange(3, 1, ultimaFila - 2, 5).getValues();
  return {
    cuentas: valores.map(function(fila) {
      return {
        nombre: String(fila[0] || '').trim(),
        id: Number(fila[1]) || 0,
        grupo: String(fila[2] || '').trim(),
        estado: Number(fila[3]) || 0,
        orden: Number(fila[4]) || 0
      };
    }).filter(function(cuenta) {
      return cuenta.nombre !== '';
    }).sort(function(a, b) {
      return a.orden - b.orden;
    })
  };
}


function acumularDiarioMayorEficienteV1_(diario) {
  var ultimaFila = Math.max(
    diario.getRange(diario.getMaxRows(), 1)
      .getNextDataCell(SpreadsheetApp.Direction.UP).getRow(),
    diario.getRange(diario.getMaxRows(), 3)
      .getNextDataCell(SpreadsheetApp.Direction.UP).getRow(),
    diario.getRange(diario.getMaxRows(), 5)
      .getNextDataCell(SpreadsheetApp.Direction.UP).getRow()
  );
  var porCuenta = {};
  var meses = {};

  if (ultimaFila < REGOPS_MAYOR_V1.PRIMERA_FILA_DIARIO) {
    return {porCuenta: porCuenta, meses: meses};
  }

  // El Diario actual entra cómodamente en una sola lectura. Evitar múltiples
  // viajes getRange/getValues reduce de forma importante el tiempo del botón.
  var cantidad = ultimaFila - REGOPS_MAYOR_V1.PRIMERA_FILA_DIARIO + 1;
  var valores = diario.getRange(
    REGOPS_MAYOR_V1.PRIMERA_FILA_DIARIO, 1, cantidad, 6
  ).getValues();

  valores.forEach(function(fila) {
    var fecha = fila[0] instanceof Date ? fila[0] : new Date(fila[0]);
    if (isNaN(fecha.getTime())) return;
    var mes = fecha.getMonth() + 1;
    var clave = fecha.getFullYear() + '-' + (mes < 10 ? '0' : '') + mes;
    meses[clave] = true;
    acumularCuentaMesMayorEficienteV1_(porCuenta, fila[2], clave, fila[3]);
    acumularCuentaMesMayorEficienteV1_(porCuenta, fila[4], clave, fila[5]);
  });

  return {porCuenta: porCuenta, meses: meses};
}

function acumularCuentaMesMayorEficienteV1_(mapa, nombre, mes, importe) {
  nombre = String(nombre || '').trim();
  importe = Number(importe);
  if (!nombre || !isFinite(importe)) return;
  if (!mapa[nombre]) mapa[nombre] = {};
  mapa[nombre][mes] = (mapa[nombre][mes] || 0) + importe;
}


function obtenerTasasMayorEficienteV1_(mayor) {
  var actual = {
    euro: Number(mayor.getRange('B1').getValue()) || 1,
    brl: Number(mayor.getRange('B2').getValue()) || 0.20,
    ars: Number(mayor.getRange('B3').getValue()) || 1
  };
  var ultimaColumna = Math.max(mayor.getLastColumn(), 4);
  var encabezados = mayor.getRange(3, 4, 3, ultimaColumna - 3).getValues();
  var historicas = {};
  var anioActual = 0;

  for (var i = 0; i < encabezados[0].length; i++) {
    if (encabezados[1][i] !== '' && encabezados[1][i] !== null) {
      anioActual = Number(encabezados[1][i]) || anioActual;
    }
    var numeroMes = numeroMesMayorEficienteV1_(encabezados[2][i]);
    if (!anioActual || !numeroMes) continue;
    var clave = anioActual + '-' + (numeroMes < 10 ? '0' : '') + numeroMes;
    historicas[clave] = Number(encabezados[0][i]) || actual.ars;
  }

  return {actual: actual, historicas: historicas};
}


function convertirActualMayorEficienteV1_(nombre, valor, tasas) {
  if (/\bARS\b/i.test(nombre)) return valor / (tasas.ars || 1) / 1000;
  if (/\b(EURO|EUR)\b/i.test(nombre)) return valor * tasas.euro / 1000;
  if (/\b(BRL|REAL|REALES)\b/i.test(nombre)) return valor * tasas.brl / 1000;
  return valor / 1000;
}


function definicionesResumenMayorEficienteV1_() {
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


function aplicarFormatoMayorEficienteV1_(hoja, filas, columnas, cantidadMeses) {
  hoja.getRange(1, 1, filas, columnas).setFontFamily('Nunito');
  hoja.getRange(1, 2, filas, 1)
    .setNumberFormat('#,##0;[Red](#,##0);-')
    .setFontSize(10);
  hoja.getRange(1, 3, filas, 1)
    .setNumberFormat('#,##0;[Red](#,##0);-')
    .setFontSize(8);
  if (cantidadMeses) {
    hoja.getRange(3, 4, filas - 2, cantidadMeses)
      .setNumberFormat('#,##0;[Red](#,##0);-')
      .setFontSize(8);
  }
  hoja.setFrozenRows(6);
  hoja.setFrozenColumns(3);
  hoja.setHiddenGridlines(false);
}


function asegurarColumnasMayorEficienteV1_(hoja, necesarias) {
  if (hoja.getMaxColumns() < necesarias) {
    hoja.insertColumnsAfter(
      hoja.getMaxColumns(),
      necesarias - hoja.getMaxColumns()
    );
  }
}


function combinarAniosMayorEficienteV1_(hoja, meses) {
  if (!meses.length) return;
  var inicio = 0;
  while (inicio < meses.length) {
    var anio = meses[inicio].substring(0, 4);
    var fin = inicio;
    while (fin + 1 < meses.length &&
      meses[fin + 1].substring(0, 4) === anio) fin++;
    var rango = hoja.getRange(4, 4 + inicio, 1, fin - inicio + 1);
    if (fin > inicio) rango.merge();
    rango.setValue(Number(anio)).setHorizontalAlignment('center');
    inicio = fin + 1;
  }
}


function nombreVisibleMayorEficienteV1_(nombre) {
  var texto = String(nombre || '').trim();
  return /^Retiro\b/i.test(texto)
    ? texto.replace(/\bARS\b/gi, '').replace(/\s+/g, ' ').trim()
    : texto;
}


function normalizarMayorEficienteV1_(texto) {
  return String(texto || '').replace(/\u00a0/g, ' ')
    .replace(/\s+/g, ' ').trim();
}


function buscarFilaMayorEficienteV1_(datos, etiqueta) {
  var buscada = normalizarMayorEficienteV1_(etiqueta);
  for (var i = 0; i < datos.length; i++) {
    if (normalizarMayorEficienteV1_(datos[i][0]) === buscada) return i + 1;
  }
  return 0;
}


function buscarTodasFilasMayorEficienteV1_(datos, etiqueta) {
  var filas = [];
  var buscada = normalizarMayorEficienteV1_(etiqueta);
  datos.forEach(function(fila, indice) {
    if (normalizarMayorEficienteV1_(fila[0]) === buscada) filas.push(indice + 1);
  });
  return filas;
}


function buscarFilaPrefijoMayorEficienteV1_(datos, prefijo) {
  var buscado = normalizarMayorEficienteV1_(prefijo).toLowerCase();
  for (var i = 0; i < datos.length; i++) {
    var valor = normalizarMayorEficienteV1_(datos[i][0]).toLowerCase();
    if (valor.indexOf(buscado) === 0) return i + 1;
  }
  return 0;
}


function buscarFilaAlternativasMayorEficienteV1_(datos, alternativas) {
  for (var i = 0; i < alternativas.length; i++) {
    var fila = buscarFilaMayorEficienteV1_(datos, alternativas[i]);
    if (fila) return fila;
  }
  return 0;
}


function sumarMayorEficienteV1_(cuentas, propiedad) {
  return cuentas.reduce(function(total, cuenta) {
    return total + (Number(cuenta[propiedad]) || 0);
  }, 0);
}


function abreviarMesMayorEficienteV1_(mes) {
  return ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun',
    'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'][mes - 1] || '';
}


function numeroMesMayorEficienteV1_(valor) {
  if (valor instanceof Date && !isNaN(valor.getTime())) {
    return valor.getMonth() + 1;
  }
  var texto = String(valor || '').trim().toLowerCase().substring(0, 3);
  var meses = {
    ene: 1, jan: 1, feb: 2, mar: 3, abr: 4, apr: 4,
    may: 5, jun: 6, jul: 7, ago: 8, aug: 8,
    sep: 9, oct: 10, nov: 11, dic: 12, dec: 12
  };
  return meses[texto] || Number(valor) || 0;
}


function redondearMayorEficienteV1_(numero) {
  return Math.round(Number(numero) || 0);
}
