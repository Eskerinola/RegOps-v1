/**
 * Modelo relacional de cuentas para RegOps v1.
 *
 * El Diario conserva los nombres visibles en C y E.
 * Los ID, que son la identidad estable de cada cuenta, se guardan en
 * dos columnas técnicas ocultas al final del Diario.
 */
var REGOPS_CUENTAS_V1 = {
  HOJA_CUENTAS: 'CtasDefinicion',
  HOJA_DIARIO: 'Diario',
  FILA_ENCABEZADO_CUENTAS: 2,
  PRIMERA_FILA_CUENTAS: 3,
  FILA_ENCABEZADO_DIARIO: 5,
  PRIMERA_FILA_DIARIO: 6,
  COLUMNA_CUENTA_1: 3,
  COLUMNA_CUENTA_2: 5,
  ENCABEZADO_ID_1: 'Cuenta 1 ID',
  ENCABEZADO_ID_2: 'Cuenta 2 ID',
  VERSION_MODELO: '2',
  PROPIEDAD_VERSION: 'REGOPS_V1_MODELO_CUENTAS',
  PROPIEDAD_SIGUIENTE_ID: 'REGOPS_V1_SIGUIENTE_ID'
};


/**
 * Completa en CtasDefinicion de v1 la estructura del plan de cuentas de v2.
 * Solo toma metadatos cuando el nombre coincide exactamente.
 * Si una cuenta ya tiene ID en v1, lo conserva: un ID nunca se modifica.
 */
function actualizarEstructuraCuentasDesdeV2() {
  var ID_REGOPS_V2 = '18doMvPGnRoUh6uIPsXx-MkYS7X1477qlNG3DtHFGn_A';
  var NOMBRE_HOJA = REGOPS_CUENTAS_V1.HOJA_CUENTAS;

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

  var ultimaFilaV1 = ultimaFilaConCuentaV1_(hojaV1);
  if (ultimaFilaV1 < REGOPS_CUENTAS_V1.PRIMERA_FILA_CUENTAS) {
    throw new Error('El plan de cuentas de RegOps v1 está vacío.');
  }

  var cantidadFilas = ultimaFilaV1 - REGOPS_CUENTAS_V1.PRIMERA_FILA_CUENTAS + 1;
  var datosV1 = hojaV1
    .getRange(REGOPS_CUENTAS_V1.PRIMERA_FILA_CUENTAS, 1, cantidadFilas, 5)
    .getValues();

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
    return [
      cuenta,
      fila[1] !== '' && fila[1] !== null ? fila[1] : metadata[0],
      metadata[1],
      metadata[2],
      metadata[3]
    ];
  });

  hojaV1
    .getRange(REGOPS_CUENTAS_V1.PRIMERA_FILA_CUENTAS, 1, cantidadFilas, 5)
    .setValues(salida);

  instalarModeloRelacionalCuentasV1_(ssV1);

  var mensaje = coincidencias + ' cuentas actualizadas desde v2';
  if (sinCoincidencia.length) {
    mensaje += '; ' + sinCoincidencia.length + ' sin coincidencia exacta';
  }

  ssV1.toast(mensaje, 'RegOps v1', 8);
  Logger.log('Cuentas sin coincidencia exacta: ' + JSON.stringify(sinCoincidencia));
}


/**
 * Instalación manual, disponible desde el menú RegOps v1.
 * Migra los movimientos existentes, configura los selectores y oculta los ID.
 */
function instalarModeloRelacionalCuentasV1() {
  instalarModeloRelacionalCuentasV1_(SpreadsheetApp.getActive());
}


function instalarModeloRelacionalCuentasV1_(ss) {
  ss = ss || SpreadsheetApp.getActive();

  var hojaCuentas = ss.getSheetByName(REGOPS_CUENTAS_V1.HOJA_CUENTAS);
  var diario = ss.getSheetByName(REGOPS_CUENTAS_V1.HOJA_DIARIO);
  if (!hojaCuentas || !diario) {
    throw new Error('Se requieren las hojas Diario y CtasDefinicion.');
  }

  normalizarPlanCuentasV1_(hojaCuentas);
  aplicarFormatoEstadosCuentasV1_(hojaCuentas);

  var mapa = obtenerMapaCuentasV1_(hojaCuentas);
  var columnasId = asegurarColumnasIdDiarioV1_(diario);
  configurarSelectoresDiarioV1_(diario, mapa);
  var migracion = sincronizarIdsDiarioV1_(diario, mapa, columnasId);

  PropertiesService.getDocumentProperties().setProperty(
    REGOPS_CUENTAS_V1.PROPIEDAD_VERSION,
    REGOPS_CUENTAS_V1.VERSION_MODELO
  );

  SpreadsheetApp.flush();

  var mensaje = migracion.filas + ' movimientos vinculados por ID';
  if (migracion.sinCuenta.length) {
    mensaje += '; ' + migracion.sinCuenta.length + ' nombres sin correspondencia';
  }
  ss.toast(mensaje, 'Modelo de cuentas instalado', 8);
  Logger.log('Nombres sin ID: ' + JSON.stringify(migracion.sinCuenta));
}


/**
 * En cada apertura solo comprueba las columnas técnicas.
 * La migración completa se ejecuta una sola vez desde el menú.
 */
function asegurarModeloRelacionalCuentasV1_(ss) {
  ss = ss || SpreadsheetApp.getActive();
  var diario = ss.getSheetByName(REGOPS_CUENTAS_V1.HOJA_DIARIO);
  if (!diario) return;

  asegurarColumnasIdDiarioV1_(diario);

  var version = PropertiesService.getDocumentProperties().getProperty(
    REGOPS_CUENTAS_V1.PROPIEDAD_VERSION
  );

  if (version !== REGOPS_CUENTAS_V1.VERSION_MODELO) {
    ss.toast(
      'Ejecutá "Instalar modelo de cuentas por ID" una sola vez.',
      'RegOps v1',
      8
    );
  }
}


/**
 * Se ejecuta cuando se modifica CtasDefinicion.
 * Asigna ID y valores predeterminados, ordena y renumera Orden.
 */
function procesarEdicionPlanCuentasV1_(ss) {
  ss = ss || SpreadsheetApp.getActive();
  var hoja = ss.getSheetByName(REGOPS_CUENTAS_V1.HOJA_CUENTAS);
  var diario = ss.getSheetByName(REGOPS_CUENTAS_V1.HOJA_DIARIO);
  if (!hoja || !diario) return;

  normalizarPlanCuentasV1_(hoja);
  aplicarFormatoEstadosCuentasV1_(hoja);

  var mapa = obtenerMapaCuentasV1_(hoja);
  configurarSelectoresDiarioV1_(diario, mapa);
  asegurarColumnasIdDiarioV1_(diario);

  SpreadsheetApp.flush();
  ss.toast('Plan de cuentas ordenado y selectores actualizados.', 'RegOps v1', 5);
}


/**
 * Actualiza los ID ocultos cuando se elige o borra una cuenta en C o E.
 * Admite una edición individual o el pegado de varias filas.
 */
function actualizarIdsCuentasEditadasV1_(diario, rangoEditado) {
  var primeraColumna = rangoEditado.getColumn();
  var ultimaColumna = rangoEditado.getLastColumn();
  var tocaCuenta1 =
    primeraColumna <= REGOPS_CUENTAS_V1.COLUMNA_CUENTA_1 &&
    ultimaColumna >= REGOPS_CUENTAS_V1.COLUMNA_CUENTA_1;
  var tocaCuenta2 =
    primeraColumna <= REGOPS_CUENTAS_V1.COLUMNA_CUENTA_2 &&
    ultimaColumna >= REGOPS_CUENTAS_V1.COLUMNA_CUENTA_2;

  if (!tocaCuenta1 && !tocaCuenta2) return;

  var primeraFila = Math.max(
    rangoEditado.getRow(),
    REGOPS_CUENTAS_V1.PRIMERA_FILA_DIARIO
  );
  var ultimaFila = rangoEditado.getLastRow();
  if (ultimaFila < primeraFila) return;

  var ss = diario.getParent();
  var hojaCuentas = ss.getSheetByName(REGOPS_CUENTAS_V1.HOJA_CUENTAS);
  if (!hojaCuentas) throw new Error('No se encontró CtasDefinicion.');

  var mapa = obtenerMapaCuentasV1_(hojaCuentas);
  var columnasId = asegurarColumnasIdDiarioV1_(diario);
  var cantidadFilas = ultimaFila - primeraFila + 1;

  if (tocaCuenta1) {
    escribirIdsDesdeNombresV1_(
      diario,
      primeraFila,
      cantidadFilas,
      REGOPS_CUENTAS_V1.COLUMNA_CUENTA_1,
      columnasId.cuenta1,
      mapa
    );
  }

  if (tocaCuenta2) {
    escribirIdsDesdeNombresV1_(
      diario,
      primeraFila,
      cantidadFilas,
      REGOPS_CUENTAS_V1.COLUMNA_CUENTA_2,
      columnasId.cuenta2,
      mapa
    );
  }
}


function escribirIdsDesdeNombresV1_(
  diario,
  primeraFila,
  cantidadFilas,
  columnaNombre,
  columnaId,
  mapa
) {
  var nombres = diario
    .getRange(primeraFila, columnaNombre, cantidadFilas, 1)
    .getDisplayValues();

  var ids = nombres.map(function(fila) {
    var nombre = String(fila[0] || '').trim();
    if (!nombre) return [''];

    var cuenta = mapa.porNombre.get(nombre);
    return [cuenta ? cuenta.id : ''];
  });

  diario
    .getRange(primeraFila, columnaId, cantidadFilas, 1)
    .setValues(ids)
    .setNumberFormat('0');
}


function sincronizarIdsDiarioV1_(diario, mapa, columnasId) {
  var ultimaFila = Math.max(
    diario.getRange(diario.getMaxRows(), 1)
      .getNextDataCell(SpreadsheetApp.Direction.UP)
      .getRow(),
    diario.getRange(diario.getMaxRows(), REGOPS_CUENTAS_V1.COLUMNA_CUENTA_1)
      .getNextDataCell(SpreadsheetApp.Direction.UP)
      .getRow(),
    diario.getRange(diario.getMaxRows(), REGOPS_CUENTAS_V1.COLUMNA_CUENTA_2)
      .getNextDataCell(SpreadsheetApp.Direction.UP)
      .getRow()
  );

  if (ultimaFila < REGOPS_CUENTAS_V1.PRIMERA_FILA_DIARIO) {
    return { filas: 0, sinCuenta: [] };
  }

  var cantidadFilas = ultimaFila - REGOPS_CUENTAS_V1.PRIMERA_FILA_DIARIO + 1;
  var nombres1 = diario
    .getRange(
      REGOPS_CUENTAS_V1.PRIMERA_FILA_DIARIO,
      REGOPS_CUENTAS_V1.COLUMNA_CUENTA_1,
      cantidadFilas,
      1
    )
    .getDisplayValues();
  var nombres2 = diario
    .getRange(
      REGOPS_CUENTAS_V1.PRIMERA_FILA_DIARIO,
      REGOPS_CUENTAS_V1.COLUMNA_CUENTA_2,
      cantidadFilas,
      1
    )
    .getDisplayValues();

  var sinCuenta = new Set();
  var ids1 = [];
  var ids2 = [];

  for (var i = 0; i < cantidadFilas; i++) {
    var nombre1 = String(nombres1[i][0] || '').trim();
    var nombre2 = String(nombres2[i][0] || '').trim();
    var cuenta1 = nombre1 ? mapa.porNombre.get(nombre1) : null;
    var cuenta2 = nombre2 ? mapa.porNombre.get(nombre2) : null;

    if (nombre1 && !cuenta1) sinCuenta.add(nombre1);
    if (nombre2 && !cuenta2) sinCuenta.add(nombre2);

    ids1.push([cuenta1 ? cuenta1.id : '']);
    ids2.push([cuenta2 ? cuenta2.id : '']);
  }

  diario
    .getRange(
      REGOPS_CUENTAS_V1.PRIMERA_FILA_DIARIO,
      columnasId.cuenta1,
      cantidadFilas,
      1
    )
    .setValues(ids1)
    .setNumberFormat('0');

  diario
    .getRange(
      REGOPS_CUENTAS_V1.PRIMERA_FILA_DIARIO,
      columnasId.cuenta2,
      cantidadFilas,
      1
    )
    .setValues(ids2)
    .setNumberFormat('0');

  diario.hideColumns(columnasId.cuenta1);
  diario.hideColumns(columnasId.cuenta2);

  return {
    filas: cantidadFilas,
    sinCuenta: Array.from(sinCuenta)
  };
}


function configurarSelectoresDiarioV1_(diario, mapa) {
  var cuentasActivas = mapa.cuentas
    .filter(function(cuenta) { return cuenta.estado !== 0; })
    .sort(function(a, b) {
      return a.orden - b.orden ||
        a.nombre.localeCompare(b.nombre, 'es', { sensitivity: 'base' });
    })
    .map(function(cuenta) { return cuenta.nombre; });

  if (!cuentasActivas.length) {
    throw new Error('No hay cuentas activas para configurar los selectores.');
  }

  var regla = SpreadsheetApp.newDataValidation()
    .requireValueInList(cuentasActivas, true)
    .setAllowInvalid(false)
    .setHelpText('Seleccione una cuenta activa. La relación interna se guarda por ID.')
    .build();

  var cantidadFilas = diario.getMaxRows() - REGOPS_CUENTAS_V1.PRIMERA_FILA_DIARIO + 1;
  diario
    .getRange(
      REGOPS_CUENTAS_V1.PRIMERA_FILA_DIARIO,
      REGOPS_CUENTAS_V1.COLUMNA_CUENTA_1,
      cantidadFilas,
      1
    )
    .setDataValidation(regla);
  diario
    .getRange(
      REGOPS_CUENTAS_V1.PRIMERA_FILA_DIARIO,
      REGOPS_CUENTAS_V1.COLUMNA_CUENTA_2,
      cantidadFilas,
      1
    )
    .setDataValidation(regla);
}


function asegurarColumnasIdDiarioV1_(diario) {
  var filaEncabezado = REGOPS_CUENTAS_V1.FILA_ENCABEZADO_DIARIO;
  var ultimaColumna = diario.getMaxColumns();
  var encabezados = diario
    .getRange(filaEncabezado, 1, 1, ultimaColumna)
    .getDisplayValues()[0];

  var columna1 = encabezados.indexOf(REGOPS_CUENTAS_V1.ENCABEZADO_ID_1) + 1;
  var columna2 = encabezados.indexOf(REGOPS_CUENTAS_V1.ENCABEZADO_ID_2) + 1;

  var faltantes = (columna1 ? 0 : 1) + (columna2 ? 0 : 1);
  if (faltantes > 0) {
    diario.insertColumnsAfter(diario.getMaxColumns(), faltantes);
    var primeraNueva = diario.getMaxColumns() - faltantes + 1;

    if (!columna1) {
      columna1 = primeraNueva;
      diario.getRange(filaEncabezado, columna1)
        .setValue(REGOPS_CUENTAS_V1.ENCABEZADO_ID_1);
      primeraNueva++;
    }

    if (!columna2) {
      columna2 = primeraNueva;
      diario.getRange(filaEncabezado, columna2)
        .setValue(REGOPS_CUENTAS_V1.ENCABEZADO_ID_2);
    }
  }

  diario.getRange(filaEncabezado, columna1, 1, 1)
    .setValue(REGOPS_CUENTAS_V1.ENCABEZADO_ID_1);
  diario.getRange(filaEncabezado, columna2, 1, 1)
    .setValue(REGOPS_CUENTAS_V1.ENCABEZADO_ID_2);

  diario.hideColumns(columna1);
  diario.hideColumns(columna2);

  return { cuenta1: columna1, cuenta2: columna2 };
}


function normalizarPlanCuentasV1_(hoja) {
  var primeraFila = REGOPS_CUENTAS_V1.PRIMERA_FILA_CUENTAS;
  var ultimaFilaAnterior = ultimaFilaConCuentaV1_(hoja);
  if (ultimaFilaAnterior < primeraFila) return;

  var cantidadAnterior = ultimaFilaAnterior - primeraFila + 1;
  var datos = hoja.getRange(primeraFila, 1, cantidadAnterior, 5).getValues();
  var cuentas = datos.filter(function(fila) {
    return String(fila[0] || '').trim() !== '';
  });

  var nombres = new Set();
  var ids = new Set();
  var maxId = 0;

  cuentas.forEach(function(fila) {
    var nombre = String(fila[0] || '').trim();
    var claveNombre = nombre.toLocaleLowerCase('es');

    if (nombres.has(claveNombre)) {
      throw new Error('Nombre de cuenta duplicado: ' + nombre);
    }
    nombres.add(claveNombre);

    var id = Number(fila[1]);
    if (id > 0) {
      if (ids.has(id)) throw new Error('ID de cuenta duplicado: ' + id);
      ids.add(id);
      maxId = Math.max(maxId, id);
    }
  });

  var propiedades = PropertiesService.getDocumentProperties();
  var siguienteIdGuardado = Number(
    propiedades.getProperty(REGOPS_CUENTAS_V1.PROPIEDAD_SIGUIENTE_ID)
  ) || 1;
  var siguienteId = Math.max(maxId + 1, siguienteIdGuardado);

  cuentas = cuentas.map(function(fila) {
    var nombre = String(fila[0] || '').trim();
    var id = Number(fila[1]);

    if (!(id > 0)) {
      while (ids.has(siguienteId)) siguienteId++;
      id = siguienteId++;
      ids.add(id);
    }

    var grupo = normalizarGrupoCuentaV1_(fila[2], nombre);
    var estado =
      fila[3] === '' || fila[3] === null
        ? 1
        : Number(fila[3]);

    return {
      nombre: nombre,
      id: id,
      grupo: grupo,
      estado: estado,
      orden: 0
    };
  });

  var prioridadGrupo = { Activo: 1, Patrimonio: 2, Resultado: 3 };
  cuentas.sort(function(a, b) {
    return (prioridadGrupo[a.grupo] || 99) - (prioridadGrupo[b.grupo] || 99) ||
      a.nombre.localeCompare(b.nombre, 'es', { sensitivity: 'base' });
  });

  var baseOrden = { Activo: 1000, Patrimonio: 4010, Resultado: 5010 };
  var posicionGrupo = { Activo: 0, Patrimonio: 0, Resultado: 0 };

  cuentas.forEach(function(cuenta) {
    var posicion = posicionGrupo[cuenta.grupo] || 0;
    cuenta.orden = (baseOrden[cuenta.grupo] || 9000) + posicion * 10;
    posicionGrupo[cuenta.grupo] = posicion + 1;
  });

  hoja.getRange('A2:E2').setValues([[
    'Cuenta',
    'ID',
    'Grupo',
    'Estado (1=Activa, 0=Inactiva)',
    'Orden'
  ]]);

  hoja.getRange(primeraFila, 1, cantidadAnterior, 5).clearContent();

  if (cuentas.length) {
    hoja.getRange(primeraFila, 1, cuentas.length, 5).setValues(
      cuentas.map(function(cuenta) {
        return [
          cuenta.nombre,
          cuenta.id,
          cuenta.grupo,
          cuenta.estado,
          cuenta.orden
        ];
      })
    );
  }

  propiedades.setProperty(
    REGOPS_CUENTAS_V1.PROPIEDAD_SIGUIENTE_ID,
    String(siguienteId)
  );

  aplicarFormatoPlanCuentasV1_(hoja, Math.max(cuentas.length, 1));
}


function normalizarGrupoCuentaV1_(grupo, nombre) {
  var texto = String(grupo || '').trim().toLowerCase();

  if (texto === 'activo') return 'Activo';
  if (texto === 'patrimonio') return 'Patrimonio';
  if (texto === 'resultado') return 'Resultado';

  if (/^(CJ|CC)\b/i.test(nombre)) return 'Activo';
  if (/^INV/i.test(nombre)) return 'Patrimonio';
  return 'Resultado';
}


function aplicarFormatoPlanCuentasV1_(hoja, cantidadFilas) {
  hoja.getRange('A2:E2')
    .setBackground('#424242')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  hoja.getRange(3, 1, cantidadFilas, 5)
    .setFontFamily('Nunito')
    .setFontSize(9);

  hoja.getRange(3, 2, cantidadFilas, 1).setNumberFormat('0');
  hoja.getRange(3, 4, cantidadFilas, 2).setNumberFormat('0');
  hoja.setFrozenRows(2);
  hoja.setColumnWidth(1, 310);
  hoja.setColumnWidth(2, 70);
  hoja.setColumnWidth(3, 110);
  hoja.setColumnWidth(4, 220);
  hoja.setColumnWidth(5, 80);
}


function aplicarFormatoEstadosCuentasV1_(hoja) {
  var formula = '=$D3=2';
  var reglas = hoja.getConditionalFormatRules().filter(function(regla) {
    try {
      var condicion = regla.getBooleanCondition();
      if (!condicion) return true;
      var valores = condicion.getCriteriaValues();
      return !(valores && String(valores[0]) === formula);
    } catch (err) {
      return true;
    }
  });

  var rango = hoja.getRange(3, 1, hoja.getMaxRows() - 2, 5);
  var reglaEstado2 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(formula)
    .setBackground('#f4cccc')
    .setRanges([rango])
    .build();

  reglas.unshift(reglaEstado2);
  hoja.setConditionalFormatRules(reglas);
}


function obtenerMapaCuentasV1_(hoja) {
  var ultimaFila = ultimaFilaConCuentaV1_(hoja);
  var cuentas = [];
  var porNombre = new Map();
  var porId = new Map();

  if (ultimaFila < REGOPS_CUENTAS_V1.PRIMERA_FILA_CUENTAS) {
    return { cuentas: cuentas, porNombre: porNombre, porId: porId };
  }

  var datos = hoja
    .getRange(
      REGOPS_CUENTAS_V1.PRIMERA_FILA_CUENTAS,
      1,
      ultimaFila - REGOPS_CUENTAS_V1.PRIMERA_FILA_CUENTAS + 1,
      5
    )
    .getValues();

  datos.forEach(function(fila) {
    var nombre = String(fila[0] || '').trim();
    var id = Number(fila[1]);
    if (!nombre || !(id > 0)) return;

    var cuenta = {
      nombre: nombre,
      id: id,
      grupo: String(fila[2] || '').trim(),
      estado: Number(fila[3]),
      orden: Number(fila[4]) || 0
    };

    cuentas.push(cuenta);
    porNombre.set(nombre, cuenta);
    porId.set(id, cuenta);
  });

  return { cuentas: cuentas, porNombre: porNombre, porId: porId };
}


function ultimaFilaConCuentaV1_(hoja) {
  return hoja
    .getRange(hoja.getMaxRows(), 1)
    .getNextDataCell(SpreadsheetApp.Direction.UP)
    .getRow();
}
