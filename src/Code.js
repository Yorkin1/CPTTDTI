/**
 * ============================================================
 *  Gestor de Clientes y Citas — Backend (Google Apps Script)
 * ============================================================
 *  Aplicación genérica para administrar clientes, citas (con eventos
 *  en Google Calendar) e historial.
 *
 *  Base de datos: Google Sheets (4 hojas)
 *  Programación:  Google Calendar
 *
 *  Todos los comentarios y textos están en español.
 * ============================================================
 */

// Nombres de las hojas dentro de la hoja de cálculo.
var HOJA_CLIENTES      = 'Clientes';
var HOJA_CITAS         = 'Citas';
var HOJA_HISTORIAL     = 'Historial';
var HOJA_USUARIOS      = 'Usuarios';
var HOJA_SERVICIOS     = 'Servicios';
var HOJA_CONFIGURACION = 'Configuracion';
var HOJA_ACTIVIDAD     = 'Actividad';

// Máximo de registros de actividad a conservar (se eliminan los más antiguos).
var MAX_ACTIVIDAD = 1000;

// Clave usada para guardar el ID de la hoja de cálculo en PropertiesService.
var PROP_ID_HOJA = 'ID_HOJA_CALCULO';

// Clave usada para guardar el ID de la carpeta de Drive de imágenes.
var PROP_CARPETA_IMAGENES = 'CARPETA_IMAGENES_ID';

// Clave usada para guardar las sesiones activas en PropertiesService.
var PROP_SESIONES = 'SESIONES_ACTIVAS';

// Clave usada para guardar en PropertiesService los usuarios que ya fueron
// notificados como "nuevos" al administrador (aviso de UNA sola vez).
var PROP_USUARIOS_AVISADOS = 'USUARIOS_AVISADOS';

// Duración de una sesión de acceso (12 horas por defecto).
var DURACION_SESION_MS = 12 * 60 * 60 * 1000;

// Claves y TTL del caché de rendimiento (CacheService).
var PROP_CONFIG_CACHE = 'CONFIG_CACHE';
// CacheService limita el TTL a 6 horas (21600 s); la sesión dura 12 h, así
// que el caché se usa como acelerador y Properties sigue siendo la fuente.
var CACHE_SESIONES_TTL_SEG = Math.min(Math.floor(DURACION_SESION_MS / 1000), 21600);
var CACHE_CONFIG_TTL_SEG = 60;

// Encabezados de cada hoja de datos.
var ENCABEZADOS = {
  Clientes:  ['ID_Cliente', 'Nombre', 'Apellido', 'Telefono', 'Email', 'Direccion', 'Notas', 'Fecha_Registro', 'Foto'],
  Citas:     ['ID_Cita', 'ID_Cliente', 'Titulo', 'Fecha', 'Hora', 'Duracion_Mins', 'Descripcion', 'ID_Evento_Calendar', 'Estado', 'Servicios', 'Total_Precio', 'Agendado_Por'],
  Historial: ['ID_Registro', 'ID_Cliente', 'ID_Cita', 'Fecha', 'Descripcion', 'Resultado'],
  Usuarios:  ['ID_Usuario', 'Nombre', 'Email', 'Salt', 'Hash', 'Rol', 'Activo', 'Fecha_Registro'],
  Servicios: ['ID_Servicio', 'Nombre', 'Precio', 'Duracion_Mins', 'Descripcion', 'Activo'],
  Actividad: ['ID_Registro', 'Fecha', 'Usuario', 'Email', 'Rol', 'Modulo', 'Accion', 'Detalle']
};

// Encabezados de la hoja de configuración (pares clave/valor).
var ENCABEZADOS_CONFIGURACION = ['Clave', 'Valor'];

// Valores predeterminados de configuración (se crean si faltan).
var CONFIGURACION_PREDETERMINADA = {
  CONFIGURADA:                    'NO',
  NOMBRE_NEGOCIO:                 'Mi Negocio',
  LOGO_URL:                       '',
  MENSAJE_BIENVENIDA:             'Bienvenido a nuestro sistema de gestión',
  ETIQUETA_CITA:                  'Cita',
  COLOR_PRIMARIO:                 '#4285F4',
  COLOR_SECUNDARIO:               '#16a34a',
  TEMA:                           'claro',
  DURACION_CITA_PREDETERMINADA:   '60',
  // Reservas en línea (página pública para que el cliente elija horario).
  HABILITAR_RESERVAS:             'SI',
  // Monitor de actividad exclusivo del dueño (primer administrador).
  HABILITAR_MONITOR:              'SI',
  // Horario de atención: { getDay: {abre, cierra} }. getDay: 0=Domingo...6=Sábado.
  // Sin "abre"/"cierra" el día queda cerrado. Default Lun-Vie 09:00-17:00, Sáb 09:00-13:00.
  HORARIO_ATENCION:               '{"0":{"abre":"","cierra":""},"1":{"abre":"09:00","cierra":"17:00"},"2":{"abre":"09:00","cierra":"17:00"},"3":{"abre":"09:00","cierra":"17:00"},"4":{"abre":"09:00","cierra":"17:00"},"5":{"abre":"09:00","cierra":"17:00"},"6":{"abre":"09:00","cierra":"13:00"}}',
  PASO_RESERVA_MIN:               '30'
};

// Orden en el que se muestran/guardan las claves de configuración.
var CLAVES_CONFIGURACION = [
  'CONFIGURADA',
  'NOMBRE_NEGOCIO',
  'LOGO_URL',
  'MENSAJE_BIENVENIDA',
  'ETIQUETA_CITA',
  'COLOR_PRIMARIO',
  'COLOR_SECUNDARIO',
  'TEMA',
  'DURACION_CITA_PREDETERMINADA',
  'HABILITAR_RESERVAS',
  'HABILITAR_MONITOR',
  'HORARIO_ATENCION',
  'PASO_RESERVA_MIN'
];

/**
 * ============================================================
 *  MENÚ PERSONALIZADO
 * ============================================================
 */

/**
 * Se ejecuta automáticamente al abrir la hoja de cálculo.
 * Crea el menú "Client Manager" con las acciones principales.
 */
/**
 * Se ejecuta automáticamente al abrir la hoja de cálculo.
 * Crea el menú "Client Manager" con las acciones principales.
 */
function onOpen() {
  // NOTA: el vínculo ID_HOJA_CALCULO se crea SOLO al ejecutar el ítem
  // "Iniciar configuración" (configurarPlantilla). No se auto-vincula aquí:
  // así cada copia debe pasar obligatoriamente por "Client Manager →
  // Iniciar configuración", que además limpia los datos que la copia pudo
  // heredar de la plantilla maestra.
  try {
    SpreadsheetApp.getUi()
      .createMenu('Client Manager')
      .addItem('Iniciar configuración', 'configurarPlantilla')
      .addItem('Verificar base de datos', 'verificarBaseDatos')
      .addItem('Limpiar datos', 'limpiarDatos')
      .addSeparator()
      .addItem('Proteger base de datos', 'protegerBaseDatos')
      .addItem('Desproteger base de datos', 'desprotegerBaseDatos')
      .addSeparator()
      .addItem('Abrir dashboard', 'abrirDashboard')
      .addToUi();
  } catch (err) {
    // En algunos contextos (sin interfaz) no existe getUi(); se ignora.
    Logger.log('No se pudo crear el menú: ' + err);
  }
}

/**
 * Limpia todo lo innecesario de la base de datos:
 * - Deja "Configuracion" con solo las claves canónicas (elimina legadas/extra).
 * - Repara las hojas de datos y elimina columnas huérfanas vacías.
 * - Purga restos de "Sujetos" y hojas vacías sobrantes.
 * Ejecutable desde el menú "Client Manager → Limpiar datos".
 * @return {Object} resumen de lo hecho.
 */
function limpiarDatos(token) {
  if (!_esContextoEditor_() && !_tieneHojaActiva_() && !_validarSesion_(token)) {
    return _respuestaSesionExpirada_();
  }
  var resumen = [];
  try {
    var hojaCalculo = obtenerHojaCalculo_();

    // 1) Normalizar Configuracion: conservar solo las claves canónicas.
    var hojaCfg = hojaCalculo.getSheetByName(HOJA_CONFIGURACION);
    if (hojaCfg) {
      var registrosAntes = Math.max(0, hojaCfg.getLastRow() - 1);
      var objetos = _leerConfiguracionComoObjeto_(hojaCfg);
      var nuevas = [];
      CLAVES_CONFIGURACION.forEach(function(clave) {
        nuevas.push([clave, (clave in objetos) ? objetos[clave] : CONFIGURACION_PREDETERMINADA[clave]]);
      });
      hojaCfg.clear({ contentsOnly: true });
      hojaCfg.getRange(1, 1, 1, ENCABEZADOS_CONFIGURACION.length)
        .setValues([ENCABEZADOS_CONFIGURACION]);
      hojaCfg.getRange(2, 1, nuevas.length, 2).setValues(nuevas);
      hojaCfg.getRange(1, 1, 1, 2)
        .setFontWeight('bold').setBackground('#1e40af').setFontColor('#ffffff');
      hojaCfg.setFrozenRows(1);
      if (registrosAntes === nuevas.length) {
        resumen.push('Configuracion: ' + nuevas.length + ' claves (ya eran las canónicas).');
      } else {
        resumen.push('Configuracion: se dejaron ' + nuevas.length + ' claves canónicas (se eliminaron ' +
          (registrosAntes - nuevas.length) + ' innecesarias).');
      }
    } else {
      _asegurarValoresConfiguracionPredeterminados_(hojaCalculo);
      resumen.push('Configuracion: hoja creada con las ' + CLAVES_CONFIGURACION.length + ' claves canónicas.');
    }
    _limpiarCacheConfig_();

    // 2) Reparar hojas de datos (elimina columnas huérfanas vacías).
    Object.keys(ENCABEZADOS).forEach(function(nombre) {
      var hoja = hojaCalculo.getSheetByName(nombre);
      var colsAntes = hoja ? hoja.getLastColumn() : 0;
      _asegurarHojaConEncabezados_(hojaCalculo, nombre, ENCABEZADOS[nombre]);
      if (colsAntes && colsAntes !== ENCABEZADOS[nombre].length) {
        resumen.push(nombre + ': columnas corregidas (' + colsAntes + ' → ' + ENCABEZADOS[nombre].length + ').');
      } else {
        resumen.push(nombre + ': encabezados correctos.');
      }
    });

    // 3) Restos de "Sujetos" (hojas huérfanas y claves legadas).
    _eliminarRestosDeSujetos_(hojaCalculo);

    // 4) Hojas iniciales vacías ("Hoja 1"/"Sheet1").
    _eliminarHojaInicialSiVacia_(hojaCalculo);

    var mensaje = 'Limpieza completada:\n\n' + resumen.join('\n');
    try {
      SpreadsheetApp.getUi().alert('Limpiar datos', mensaje, SpreadsheetApp.getUi().ButtonSet.OK);
    } catch (e) {}
    console.log('Limpiar datos:\n' + mensaje);
    Logger.log('Limpiar datos:\n' + mensaje);
    return { exito: true, mensaje: mensaje };
  } catch (err) {
    var em = 'Error al limpiar la base de datos: ' + err.message;
    try {
      SpreadsheetApp.getUi().alert('Limpiar datos', em, SpreadsheetApp.getUi().ButtonSet.OK);
    } catch (e) {}
    Logger.log(em);
    return { exito: false, mensaje: em };
  }
}

/**
 * Abre el dashboard (aplicación web) en una nueva pestaña del navegador.
 * Muestra un cuadro de diálogo con el enlace a la URL /exec desplegada.
 */
function abrirDashboard() {
  var ui = SpreadsheetApp.getUi();
  var url = obtenerUrlAplicacion();

  if (!url) {
    ui.alert(
      'Abrir dashboard',
      'Todavía no se ha desplegado la aplicación web.\n\n' +
      'Vaya a Implementar → Nueva implementación → Aplicación web, ' +
      'ejecútela como usuario y vuelva a intentarlo.',
      ui.ButtonSet.OK
    );
    return;
  }

  var html = HtmlService.createHtmlOutput(
      '<p style="font-family:Arial,sans-serif;font-size:14px;">' +
      'Haga clic en el siguiente botón para abrir el dashboard:</p>' +
      '<p style="text-align:center;margin-top:16px;">' +
      '<a href="' + url + '" target="_blank" ' +
      'style="background:#4285F4;color:#fff;padding:10px 20px;border-radius:6px;' +
      'text-decoration:none;font-family:Arial,sans-serif;font-weight:bold;">' +
      'Abrir dashboard</a></p>')
    .setWidth(360)
    .setHeight(140);
  ui.showModalDialog(html, 'Abrir dashboard');
}

/**
 * Punto de entrada de la aplicación web.
 * Sirve la página de reservas públicas (Reserva.html) cuando se accede con
 * "?v=reservar"; en caso contrario sirve la pantalla principal (Index.html).
 */
function doGet(e) {
  e = e || {};
  if (String(e.parameter.v || '').trim().toLowerCase() === 'reservar') {
    return HtmlService.createTemplateFromFile('Reserva')
      .evaluate()
      .setTitle('Reservar horario')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Gestor de Clientes y Citas')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Permite incluir archivos HTML dentro de otras plantillas.
 * Uso en HTML: <?!= include('HojaEstilos') ?>
 */
function include(nombreArchivo) {
  return HtmlService.createHtmlOutputFromFile(nombreArchivo).getContent();
}

/**
 * ============================================================
 *  CONFIGURACIÓN DE LA PLANTILLA
 * ============================================================
 */

/**
 * Configura la plantilla en la hoja de cálculo copiada por el usuario.
 *
 * Debe ejecutarse UNA VEZ después de copiar la plantilla maestra
 * (desde el menú "Client Manager" → "Iniciar configuración").
 *
 * - Toma la hoja de cálculo enlazada actual (getActiveSpreadsheet).
 * - Guarda SU PROPIO ID en PropertiesService (ID_HOJA_CALCULO), de modo
 *   que cada copia trabaja sobre su propia base de datos y nunca sobre
 *   la plantilla maestra.
 * - Usa LockService para evitar ejecuciones concurrentes.
 * - Verifica/crea las hojas requeridas y repara los encabezados SIN
 *   borrar, mover ni sobrescribir los registros existentes.
 * - Crea/valida la hoja "Configuracion" y sus valores predeterminados.
 */
function configurarPlantilla(token) {
  if (!_esContextoEditor_() && !_tieneHojaActiva_() && !_validarSesion_(token)) {
    return _respuestaSesionExpirada_();
  }
  var candado = LockService.getScriptLock();
  try {
    // Evita que dos ejecuciones simultáneas modifiquen la estructura.
    candado.waitLock(30000);
  } catch (errLock) {
    var msg = 'Ya hay una configuración en curso. Inténtelo de nuevo en unos segundos.';
    _mostrarMensaje_('Configuración', msg);
    return { exito: false, mensaje: msg };
  }

  try {
    // 1) Obtener la hoja de cálculo enlazada (la copia del usuario).
    var hojaCalculo = SpreadsheetApp.getActiveSpreadsheet();
    if (!hojaCalculo) {
      var msgSinHoja = 'No se encontró una hoja de cálculo activa. Ejecute esta ' +
        'opción desde el menú "Client Manager" dentro de su copia de la plantilla.';
      _mostrarMensaje_('Configuración', msgSinHoja);
      return { exito: false, mensaje: msgSinHoja };
    }

    // 2) Detectar si es la PRIMERA configuración de ESTA copia: cada copia
    // tiene su propio ScriptProperties, por lo que un ID aún vacío significa
    // que la plantilla recién se copió. En ese caso se eliminan los datos que
    // la copia pudo heredar de la maestra. Si el ID ya existía (copia en uso),
    // la re-ejecución es NO destructiva.
    var idPrevio = PropertiesService.getScriptProperties().getProperty(PROP_ID_HOJA);
    var esPrimeraConfiguracion = !idPrevio;

    // 3) Guardar SU PROPIO ID (clave para que cada copia use su base de datos).
    PropertiesService.getScriptProperties()
      .setProperty(PROP_ID_HOJA, hojaCalculo.getId());

    // 4) Verificar/crear hojas de datos y reparar encabezados sin tocar datos.
    Object.keys(ENCABEZADOS).forEach(function(nombreHoja) {
      _asegurarHojaConEncabezados_(hojaCalculo, nombreHoja, ENCABEZADOS[nombreHoja]);
    });

    // 5) Crear/validar la hoja de configuración y sus valores predeterminados.
    _asegurarHojaConfiguracion_(hojaCalculo);
    _asegurarValoresConfiguracionPredeterminados_(hojaCalculo);
    _limpiarCacheConfig_();

    // 6) Primera configuración de una copia recién hecha: vaciar los datos
    //    copiados de la maestra (Clientes, Citas, Historial, Usuarios) y
    //    restablecer Configuracion a sus valores canónicos (CONFIGURADA = NO).
    var limpieza = false;
    if (esPrimeraConfiguracion) {
      limpieza = _limpiarDatosCopiados_(hojaCalculo);
      _limpiarCacheConfig_();
    }

    // 7) Purgar restos de la versión anterior con "Sujetos".
    _eliminarRestosDeSujetos_(hojaCalculo);

    // 8) Eliminar "Hoja 1"/"Sheet1" SOLO si está realmente vacía.
    _eliminarHojaInicialSiVacia_(hojaCalculo);

    Logger.log('Plantilla configurada. ID de la hoja: ' + hojaCalculo.getId() +
      ' | Primera configuración: ' + esPrimeraConfiguracion);

    var exitoMsg = 'Configuración completada correctamente.\n\n' +
      'La base de datos de esta copia quedó lista con las hojas: ' +
      'Clientes, Citas, Historial, Servicios, Actividad y Configuracion.\n\n';
    if (limpieza) {
      exitoMsg += 'Como esta copia fue creada a partir de la plantilla, se eliminaron ' +
        'los datos que traía copiados y la configuración se restableció a los valores ' +
        'predeterminados.\n\n';
    }
    exitoMsg += 'Ahora despliegue la aplicación web (Implementar → Aplicación web) ' +
      'y abra la URL /exec para completar el asistente de personalización.';
    _mostrarMensaje_('Configuración completada', exitoMsg);

    return { exito: true, mensaje: exitoMsg, id: hojaCalculo.getId() };
  } catch (err) {
    Logger.log('Error al configurar la plantilla: ' + err);
    _mostrarMensaje_('Error de configuración', 'Ocurrió un error: ' + err.message);
    return { exito: false, mensaje: 'Error al configurar la plantilla: ' + err.message };
  } finally {
    candado.releaseLock();
  }
}

/**
 * Elimina los datos copiados de la plantilla maestra en la primera
 * configuración de una copia:
 * - Vacía las filas de Clientes, Citas, Historial y Usuarios conservando
 *   los encabezados (fila 1).
 * - Restablece la hoja "Configuracion" a las claves canónicas con sus
 *   valores predeterminados (CONFIGURADA = NO).
 * @param {Spreadsheet} hojaCalculo  Hoja de cálculo de la copia.
 * @return {boolean} true si se realizó alguna limpieza.
 */
function _limpiarDatosCopiados_(hojaCalculo) {
  var limpiado = false;

  // 1) Vaciar las hojas de datos (solo filas de datos, no encabezados).
  Object.keys(ENCABEZADOS).forEach(function(nombreHoja) {
    var hoja = hojaCalculo.getSheetByName(nombreHoja);
    if (!hoja) return;
    var ultimaFila = hoja.getLastRow();
    if (ultimaFila > 1) {
      hoja.getRange(2, 1, ultimaFila - 1, hoja.getLastColumn()).clearContent();
      limpiado = true;
    }
  });

  // 2) Restablecer la configuración a los valores canónicos por defecto.
  var hojaCfg = hojaCalculo.getSheetByName(HOJA_CONFIGURACION);
  if (hojaCfg) {
    var filas = [];
    CLAVES_CONFIGURACION.forEach(function(clave) {
      filas.push([clave, CONFIGURACION_PREDETERMINADA[clave]]);
    });
    hojaCfg.clear({ contentsOnly: true });
    hojaCfg.getRange(1, 1, 1, ENCABEZADOS_CONFIGURACION.length)
      .setValues([ENCABEZADOS_CONFIGURACION]);
    hojaCfg.getRange(2, 1, filas.length, 2).setValues(filas);
    hojaCfg.getRange(1, 1, 1, 2)
      .setFontWeight('bold').setBackground('#1e40af').setFontColor('#ffffff');
    hojaCfg.setFrozenRows(1);
    limpiado = true;
  }

  return limpiado;
}

/**
 * Muestra un mensaje al usuario mediante un cuadro de diálogo (si hay
 * interfaz de hoja de cálculo) o un "toast". En contexto sin interfaz
 * simplemente registra en el log.
 */
function _mostrarMensaje_(titulo, mensaje) {
  try {
    var hojaCalculo = SpreadsheetApp.getActiveSpreadsheet();
    if (hojaCalculo) {
      try {
        SpreadsheetApp.getUi().alert(titulo, mensaje, SpreadsheetApp.getUi().ButtonSet.OK);
        return;
      } catch (errUi) {
        hojaCalculo.toast(mensaje, titulo, 8);
        return;
      }
    }
  } catch (err) {
    // Sin interfaz disponible.
  }
  Logger.log('[' + titulo + '] ' + mensaje);
}

/**
 * Garantiza que una hoja de datos exista y tenga los encabezados
 * correctos SIN borrar, mover ni sobrescribir los registros existentes.
 */
function _asegurarHojaConEncabezados_(hojaCalculo, nombreHoja, encabezados) {
  var hoja = hojaCalculo.getSheetByName(nombreHoja);
  if (!hoja) {
    hoja = hojaCalculo.insertSheet(nombreHoja);
    Logger.log('Hoja creada: ' + nombreHoja);
  }

  var totalCols = encabezados.length;

  if (hoja.getLastRow() === 0) {
    // Hoja vacía: escribir encabezados.
    hoja.getRange(1, 1, 1, totalCols).setValues([encabezados]);
  } else {
    // Reparar solo los encabezados (fila 1) que falten o difieran,
    // sin tocar las filas de datos (fila 2 en adelante).
    var anchoActual = Math.max(hoja.getLastColumn(), totalCols);
    var filaActual = hoja.getRange(1, 1, 1, anchoActual).getValues()[0];
    var nuevaFila = filaActual.slice();
    var cambiar = false;
    for (var c = 0; c < totalCols; c++) {
      if (String(filaActual[c] || '') !== String(encabezados[c])) {
        nuevaFila[c] = encabezados[c];
        cambiar = true;
      }
    }
    if (cambiar) {
      hoja.getRange(1, 1, 1, nuevaFila.length).setValues([nuevaFila]);
      Logger.log('Encabezados reparados en: ' + nombreHoja);
    }

    // Eliminar columnas huérfanas (restos de versiones anteriores, p. ej.
    // "ID_Sujeto") que estén totalmente vacías, dejando la hoja con
    // exactamente el número de columnas del esquema actual.
    var ultimaCol = hoja.getLastColumn();
    if (ultimaCol > totalCols) {
      var nFilasDatos = Math.max(0, hoja.getLastRow() - 1);
      for (var c = ultimaCol; c > totalCols; c--) {
        var vacio = true;
        if (nFilasDatos > 0) {
          var datosCol = hoja.getRange(2, c, nFilasDatos, 1).getValues();
          for (var r = 0; r < datosCol.length; r++) {
            if (String(datosCol[r][0] || '') !== '') { vacio = false; break; }
          }
        }
        if (vacio) {
          hoja.deleteColumn(c);
          Logger.log('Columna huérfana eliminada en ' + nombreHoja + ': columna ' + c);
        }
      }
    }
  }

  // Estilo de encabezados (no afecta datos).
  hoja.getRange(1, 1, 1, totalCols)
    .setFontWeight('bold')
    .setBackground('#1e40af')
    .setFontColor('#ffffff');
  hoja.setFrozenRows(1);
  return hoja;
}

/**
 * Garantiza la existencia de la hoja "Configuracion" con sus encabezados
 * (Clave / Valor), sin sobrescribir valores ya guardados.
 */
function _asegurarHojaConfiguracion_(hojaCalculo) {
  var hoja = hojaCalculo.getSheetByName(HOJA_CONFIGURACION);
  if (!hoja) {
    hoja = hojaCalculo.insertSheet(HOJA_CONFIGURACION);
    Logger.log('Hoja de configuración creada.');
  }
  if (hoja.getLastRow() === 0) {
    hoja.getRange(1, 1, 1, ENCABEZADOS_CONFIGURACION.length)
      .setValues([ENCABEZADOS_CONFIGURACION]);
  } else {
    // Reparar encabezados si difieren.
    var fila = hoja.getRange(1, 1, 1, ENCABEZADOS_CONFIGURACION.length).getValues()[0];
    var reparar = false;
    for (var i = 0; i < ENCABEZADOS_CONFIGURACION.length; i++) {
      if (String(fila[i] || '') !== ENCABEZADOS_CONFIGURACION[i]) { reparar = true; break; }
    }
    if (reparar) {
      hoja.getRange(1, 1, 1, ENCABEZADOS_CONFIGURACION.length)
        .setValues([ENCABEZADOS_CONFIGURACION]);
    }
  }
  hoja.getRange(1, 1, 1, ENCABEZADOS_CONFIGURACION.length)
    .setFontWeight('bold')
    .setBackground('#1e40af')
    .setFontColor('#ffffff');
  hoja.setFrozenRows(1);
  return hoja;
}

/**
 * Agrega los valores de configuración predeterminados que falten,
 * sin sobrescribir los que el usuario ya haya definido.
 */
function _asegurarValoresConfiguracionPredeterminados_(hojaCalculo) {
  var hoja = _asegurarHojaConfiguracion_(hojaCalculo);
  var existentes = _leerConfiguracionComoObjeto_(hoja);

  CLAVES_CONFIGURACION.forEach(function(clave) {
    if (!(clave in existentes)) {
      hoja.appendRow([clave, CONFIGURACION_PREDETERMINADA[clave]]);
      Logger.log('Configuración predeterminada agregada: ' + clave);
    }
  });
}

/**
 * Lee la hoja de configuración como un objeto { clave: valor }.
 */
function _leerConfiguracionComoObjeto_(hoja) {
  var resultado = {};
  var datos = hoja.getDataRange().getValues();
  for (var i = 1; i < datos.length; i++) { // saltar encabezado
    var clave = String(datos[i][0] || '').trim();
    if (clave) {
      resultado[clave] = datos[i][1];
    }
  }
  return resultado;
}

/**
 * Elimina la hoja por defecto "Hoja 1"/"Sheet1" SOLO si está
 * completamente vacía y no es la única hoja del documento.
 */
function _eliminarHojaInicialSiVacia_(hojaCalculo) {
  ['Hoja 1', 'Hoja1', 'Sheet1', 'Sheet 1'].forEach(function(nombre) {
    var h = hojaCalculo.getSheetByName(nombre);
    if (h && hojaCalculo.getSheets().length > 1) {
      var vacia = h.getLastRow() === 0 && h.getLastColumn() === 0;
      if (!vacia) {
        // Verificación adicional: sin contenido real en el rango usado.
        try {
          var valores = h.getDataRange().getValues();
          vacia = valores.every(function(fila) {
            return fila.every(function(v) { return String(v).trim() === ''; });
          });
        } catch (err) {
          vacia = false;
        }
      }
      if (vacia) {
        hojaCalculo.deleteSheet(h);
        Logger.log('Hoja inicial vacía eliminada: ' + nombre);
      }
    }
  });
}

/**
 * Purga cualquier resto de la versión anterior con "Sujetos":
 * elimina la hoja huérfana "Sujetos" y las claves legadas de configuración.
 */
function _eliminarRestosDeSujetos_(hojaCalculo) {
  // 1) Eliminar cualquier hoja cuyo nombre contenga "sujet" (restos de la
  //    versión anterior con "Sujetos"), sea "Sujetos", "Sujeto", etc.
  hojaCalculo.getSheets().forEach(function(h) {
    if (/sujet/i.test(String(h.getName()))) {
      hojaCalculo.deleteSheet(h);
      Logger.log('Hoja huérfana de "Sujetos" eliminada: ' + h.getName());
    }
  });

  try {
    var hojaCfg = hojaCalculo.getSheetByName('Configuracion');
    if (hojaCfg) {
      var legates = ['ETIQUETA_SUJETO', 'ETIQUETA_SUJETOS'];
      var datos = hojaCfg.getDataRange().getValues();
      var conservar = [];
      var cambio = false;
      for (var i = 1; i < datos.length; i++) {
        var clave = String(datos[i][0] || '').trim();
        if (legates.indexOf(clave) !== -1) { cambio = true; continue; }
        conservar.push(datos[i]);
      }
      if (cambio) {
        hojaCfg.clear({ contentsOnly: true });
        if (conservar.length > 0) {
          hojaCfg.getRange(2, 1, conservar.length, conservar[0].length).setValues(conservar);
          hojaCfg.getRange(1, 1, 1, 2).setValues([['Clave', 'Valor']]);
        }
        Logger.log('Claves de configuración legadas de "Sujetos" eliminadas.');
      }
    }
  } catch (err) {
    Logger.log('No se pudieron eliminar claves legadas de "Sujetos": ' + err.message);
  }
}

/**
 * Función pública de limpieza: elimina CUALQUIER resto de "Sujetos"
 * (hoja huérfana, columnas ID_Sujeto y claves de configuración) y devuelve
 * un resumen de lo encontrado/eliminado. Ejecutable desde el editor de
 * Apps Script (botón "Ejecutar") o desde el menú.
 */
function eliminarSujetos() {
  var resumen = [];
  try {
    var hojaCalc;
    try { hojaCalc = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) { hojaCalc = null; }
    if (!hojaCalc) hojaCalc = obtenerHojaCalculo_();

    // 1) Hojas cuyo nombre contenga "sujet".
    var borradas = [];
    hojaCalc.getSheets().forEach(function(h) {
      if (/sujet/i.test(String(h.getName()))) {
        borradas.push(h.getName());
        hojaCalc.deleteSheet(h);
      }
    });
    if (borradas.length) resumen.push('Hojas eliminadas: ' + borradas.join(', '));
    else resumen.push('No se encontró ninguna hoja de "Sujetos".');

    // 2) Columnas "ID_Sujeto" en Citas e Historial (informativas + limpieza).
    ['Citas', 'Historial'].forEach(function(nombreHoja) {
      var hoja = hojaCalc.getSheetByName(nombreHoja);
      if (!hoja) return;
      var cab = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
      var pos = cab.indexOf('ID_Sujeto');
      if (pos === -1) return;
      var esperado = ENCABEZADOS[nombreHoja].length;
      // Si la columna está fuera del esquema (al final) y está vacía, se borra.
      if (pos >= esperado) {
        var datosCol = hoja.getRange(2, pos + 1, Math.max(0, hoja.getLastRow() - 1), 1).getValues();
        var vacio = true;
        for (var r = 0; r < datosCol.length; r++) {
          if (String(datosCol[r][0] || '') !== '') { vacio = false; break; }
        }
        if (vacio) { hoja.deleteColumn(pos + 1); resumen.push('Columna ID_Sujeto eliminada de "' + nombreHoja + '".'); }
        else { resumen.push('Columna ID_Sujeto con datos en "' + nombreHoja + '" (revisar a mano).'); }
      } else {
        // Dentro del esquema: solo se relabeta el encabezado al reparar; no se borra.
        resumen.push('En "' + nombreHoja + '" la columna 3 se relabela como "Titulo" al ejecutar configuración.');
      }
    });

    // 3) Claves de configuración legadas.
    _eliminarRestosDeSujetos_(hojaCalc);

    var mensaje = resumen.join('\n');
    try { SpreadsheetApp.getUi().alert('Limpieza de "Sujetos"', mensaje, SpreadsheetApp.getUi().ButtonSet.OK); } catch (e) {}
    Logger.log(mensaje);
    return resumen;
  } catch (err) {
    var em = 'Error: ' + err.message;
    try { SpreadsheetApp.getUi().alert('Limpieza de "Sujetos"', em, SpreadsheetApp.getUi().ButtonSet.OK); } catch (e) {}
    return em;
  }
}

/**
 * ============================================================
 *  ACCESO A LA HOJA DE CÁLCULO
 * ============================================================
 *//**
 * Obtiene la hoja de cálculo asociada al proyecto.
 *
 * IMPORTANTE: nunca crea una hoja nueva (no usa SpreadsheetApp.create).
 * La plantilla es una hoja de cálculo enlazada; el ID se guarda al ejecutar
 * configurarPlantilla().
 *
 * Estrategia:
 *  1) Abrir por el ID guardado en PropertiesService (funciona también en
 *     el contexto de la aplicación web, donde NO hay hoja activa).
 *  2) Si no hay ID guardado, intentar usar la hoja activa (solo disponible
 *     al ejecutar desde el editor/menú) y guardar su ID.
 */
function obtenerHojaCalculo_() {
  var propiedades = PropertiesService.getScriptProperties();
  var idGuardado = propiedades.getProperty(PROP_ID_HOJA);

  // 1) Abrir por el ID guardado.
  if (idGuardado) {
    try {
      return SpreadsheetApp.openById(idGuardado);
    } catch (err) {
      Logger.log('No se pudo abrir la hoja guardada (' + idGuardado + '): ' + err);
    }
  }

  // 2) Usar la hoja activa si el script está enlazado a una (contexto editor).
  var activa = null;
  try {
    activa = SpreadsheetApp.getActiveSpreadsheet();
  } catch (err) {
    activa = null;
  }
  if (activa) {
    propiedades.setProperty(PROP_ID_HOJA, activa.getId());
    return activa;
  }

  throw new Error('No hay hoja de cálculo configurada. Abra su copia de la plantilla ' +
    'y ejecute "Client Manager → Iniciar configuración".');
}

/**
 * Devuelve una hoja concreta por nombre, garantizando que exista.
 *
 * En el contexto de la aplicación web (doGet) NO hay hoja activa, por lo
 * que este método NUNCA ejecuta configurarPlantilla(). En su lugar crea o
 * repara la hoja solicitada dentro de la hoja de cálculo devuelta por
 * obtenerHojaCalculo_() (patrón SpreadsheetApp.openById).
 */
function obtenerHoja_(nombreHoja) {
  var hojaCalculo = obtenerHojaCalculo_();
  var hoja = hojaCalculo.getSheetByName(nombreHoja);
  if (!hoja) {
    if (nombreHoja === HOJA_CONFIGURACION) {
      hoja = _asegurarHojaConfiguracion_(hojaCalculo);
    } else if (ENCABEZADOS[nombreHoja]) {
      hoja = _asegurarHojaConEncabezados_(hojaCalculo, nombreHoja, ENCABEZADOS[nombreHoja]);
    } else {
      hoja = hojaCalculo.insertSheet(nombreHoja);
    }
  } else if (ENCABEZADOS[nombreHoja] && hoja.getLastColumn() !== ENCABEZADOS[nombreHoja].length) {
    // Hoja existente con columnas de más o de menos: reparar estructura
    // (elimina restos huérfanos como "ID_Sujeto" de versiones anteriores).
    hoja = _asegurarHojaConEncabezados_(hojaCalculo, nombreHoja, ENCABEZADOS[nombreHoja]);
  }
  return hoja;
}

/**
 * ============================================================
 *  API DE CONFIGURACIÓN Y ESTADO (para la aplicación web)
 * ============================================================
 */

/**
 * Devuelve el estado de la aplicación con la configuración
 * relevante para el frontend.
 * @return {Object} { configurada, nombreNegocio, mensajeBienvenida,
 *                     etiquetaCita, colorPrimario, duracionCitaPredeterminada }
 */
function obtenerEstadoAplicacion() {
  try {
    _asegurarColumnaFotoEnHojas_();
    var cfg = obtenerConfiguracion();
    return {
      configurada:                  String(cfg.CONFIGURADA || 'NO').toUpperCase() === 'SI',
      nombreNegocio:                cfg.NOMBRE_NEGOCIO || CONFIGURACION_PREDETERMINADA.NOMBRE_NEGOCIO,
      logoUrl:                      cfg.LOGO_URL || '',
      mensajeBienvenida:            cfg.MENSAJE_BIENVENIDA || CONFIGURACION_PREDETERMINADA.MENSAJE_BIENVENIDA,
      etiquetaCita:                 cfg.ETIQUETA_CITA || CONFIGURACION_PREDETERMINADA.ETIQUETA_CITA,
      colorPrimario:                cfg.COLOR_PRIMARIO || CONFIGURACION_PREDETERMINADA.COLOR_PRIMARIO,
      colorSecundario:              cfg.COLOR_SECUNDARIO || CONFIGURACION_PREDETERMINADA.COLOR_SECUNDARIO,
      tema:                         cfg.TEMA || CONFIGURACION_PREDETERMINADA.TEMA,
      duracionCitaPredeterminada:   cfg.DURACION_CITA_PREDETERMINADA || CONFIGURACION_PREDETERMINADA.DURACION_CITA_PREDETERMINADA,
      habilitarReservas:            String(cfg.HABILITAR_RESERVAS || 'NO').toUpperCase() === 'SI',
      habilitarMonitor:             String(cfg.HABILITAR_MONITOR || 'NO').toUpperCase() === 'SI',
      horarioAtencion:              _parseHorarioAtencion_(cfg.HORARIO_ATENCION),
      pasoReserva:                  parseInt(cfg.PASO_RESERVA_MIN, 10) || 30
    };
  } catch (err) {
    // Sin hoja de cálculo vinculada (no se ejecutó "Iniciar configuración"):
    // se devuelve un estado amigable en vez de lanzar un error.
    Logger.log('obtenerEstadoAplicacion sin hoja vinculada: ' + err);
    return {
      configurada: false,
      sinHoja:     true,
      mensaje:     'No hay hoja de cálculo configurada. Abra su copia de la plantilla ' +
                   'y ejecute "Client Manager → Iniciar configuración".'
    };
  }
}

/**
 * Devuelve en UNA sola llamada todo lo que necesita el dashboard al arrancar:
 * configuración, clientes y citas. Reduce el tiempo de carga porque evita
 * varias invocaciones de google.script.run (cada una paga el arranque del
 * servidor). Exige una sesión válida.
 * @param {string} token  Token de sesión.
 * @return {Object} { exito, config, clientes, citas }
 */
function obtenerDatosIniciales(token) {
  if (!_validarSesion_(token)) return _respuestaSesionExpirada_();
  try {
    _asegurarColumnaFotoEnHojas_();
    return {
      exito: true,
      config: obtenerEstadoAplicacion(),
      clientes: filasAObjetos_(obtenerHoja_(HOJA_CLIENTES)),
      citas: _leerCitas_()
    };
  } catch (err) {
    Logger.log('Error en obtenerDatosIniciales: ' + err);
    return { exito: false, mensaje: 'Error al cargar los datos: ' + err.message };
  }
}

/**
 * Asegura que la hoja de clientes tenga la columna "Foto"
 * (la agrega al final si no existe). Ligero: solo escribe si falta.
 */
function _asegurarColumnaFotoEnHojas_() {
  // Solo revisar la estructura una vez por hora (evita operaciones de hoja
  // en cada carga de la aplicación).
  if (CacheService.getScriptCache().get('CHK_FOTO')) return;
  var hojaCalculo;
  try {
    hojaCalculo = obtenerHojaCalculo_();
  } catch (err) {
    return; // Sin hoja enlazada: no hacer nada.
  }
  var hoja = hojaCalculo.getSheetByName(HOJA_CLIENTES);
  if (!hoja) return;
  var esperado = ENCABEZADOS[HOJA_CLIENTES];
  var ancho = Math.max(hoja.getLastColumn(), 1);
  var filaActual = hoja.getRange(1, 1, 1, ancho).getValues()[0];
  var falta = false;
  for (var c = 0; c < esperado.length; c++) {
    if (String(filaActual[c] || '') !== esperado[c]) { falta = true; break; }
  }
  if (falta) {
    _asegurarHojaConEncabezados_(hojaCalculo, HOJA_CLIENTES, esperado);
  }
  CacheService.getScriptCache().put('CHK_FOTO', '1', 3600);
}

/**
 * Devuelve todos los pares clave/valor de la hoja Configuracion.
 * Rellena con los valores predeterminados las claves que falten
 * (sin escribir en la hoja).
 * @return {Object} { CLAVE: valor, ... }
 */
function obtenerConfiguracion() {
  // Cache de lectura en memoria (la configuración cambia poco y así no se
  // vuelve a leer la hoja en cada carga; se invalida al guardar).
  try {
    var crudo = CacheService.getScriptCache().get(PROP_CONFIG_CACHE);
    if (crudo) return JSON.parse(crudo);
  } catch (err) {
    // Si falla el caché, seguir con la lectura normal.
  }
  var hoja = obtenerHoja_(HOJA_CONFIGURACION);
  var cfg = _leerConfiguracionComoObjeto_(hoja);
  // Completar claves faltantes con predeterminados (solo en memoria).
  CLAVES_CONFIGURACION.forEach(function(clave) {
    if (!(clave in cfg) || cfg[clave] === '' || cfg[clave] === null || cfg[clave] === undefined) {
      cfg[clave] = CONFIGURACION_PREDETERMINADA[clave];
    } else {
      cfg[clave] = String(cfg[clave]);
    }
  });
  try {
    CacheService.getScriptCache().put(PROP_CONFIG_CACHE, JSON.stringify(cfg), CACHE_CONFIG_TTL_SEG);
  } catch (err) {}
  return cfg;
}

/** Invalida el caché de configuración (tras guardar o normalizar). */
function _limpiarCacheConfig_() {
  try {
    CacheService.getScriptCache().remove(PROP_CONFIG_CACHE);
  } catch (err) {}
}

/**
 * Guarda pares clave/valor de configuración en la hoja Configuracion.
 * Marca CONFIGURADA = "SI". Usa LockService para evitar escrituras
 * concurrentes. No borra otras claves existentes.
 * Exige una sesión válida EXCEPTO cuando la plantilla aún no está
 * configurada (el asistente inicial corre antes de existir cuentas).
 * @param {string} token  Token de sesión (vacío en el asistente inicial).
 * @param {Object} datos  { NOMBRE_NEGOCIO, MENSAJE_BIENVENIDA, ETIQUETA_CITA, ... }
 * @return {Object} { exito, mensaje }
 */
function guardarConfiguracion(token, datos) {
  var candado = LockService.getScriptLock();
  try {
    candado.waitLock(30000);
  } catch (errLock) {
    return { exito: false, mensaje: 'Hay otra operación en curso. Inténtelo de nuevo.' };
  }

  try {
    datos = datos || {};
    var hojaCalculo = obtenerHojaCalculo_();
    var hoja = _asegurarHojaConfiguracion_(hojaCalculo);

    // El asistente inicial puede guardar sin sesión; una vez configurada,
    // solo se permite cambiar la configuración con una sesión válida.
    var configActual = _leerConfiguracionComoObjeto_(hoja);
    if (String(configActual.CONFIGURADA || 'NO').toUpperCase() === 'SI') {
      if (!_validarSesion_(token)) return _respuestaSesionExpirada_();
    }

    // Construir el conjunto de valores a persistir.
    var aGuardar = {};
    CLAVES_CONFIGURACION.forEach(function(clave) {
      if (clave === 'CONFIGURADA') return; // se fuerza más abajo
      if (clave === 'LOGO_URL') return;    // se maneja explícitamente (permite limpiarla)
      if (datos[clave] !== undefined && datos[clave] !== null && String(datos[clave]) !== '') {
        aGuardar[clave] = String(datos[clave]);
      }
    });

    // Logo del negocio: si llega una imagen nueva se sube a Drive y se guarda
    // su URL; si se pide eliminar, se limpia el valor.
    if (datos.eliminarLogo) {
      aGuardar.LOGO_URL = '';
    } else if (datos.logoDataUrl) {
      try {
        aGuardar.LOGO_URL = _guardarImagenDesdeBase64_(datos.logoDataUrl, 'logo_' + (configActual.NOMBRE_NEGOCIO || 'negocio'));
      } catch (errLogo) {
        return { exito: false, mensaje: 'No se pudo subir el logo: ' + errLogo.message };
      }
    }

    // Al guardar desde el asistente, la plantilla queda configurada.
    aGuardar.CONFIGURADA = 'SI';

    // Leer estado actual (clave -> número de fila).
    var valores = hoja.getDataRange().getValues();
    var filaPorClave = {};
    for (var i = 1; i < valores.length; i++) {
      var c = String(valores[i][0] || '').trim();
      if (c) filaPorClave[c] = i + 1;
    }

    // Actualizar o insertar cada clave.
    Object.keys(aGuardar).forEach(function(clave) {
      var valor = aGuardar[clave];
      if (filaPorClave[clave]) {
        hoja.getRange(filaPorClave[clave], 2).setValue(valor);
      } else {
        hoja.appendRow([clave, valor]);
      }
    });

    SpreadsheetApp.flush();
    _limpiarCacheConfig_();
    Logger.log('Configuración guardada correctamente.');
    _registrarActividad_(token, 'Configuración', 'Actualizó configuración', '');
    return { exito: true, mensaje: 'Configuración guardada correctamente.' };
  } catch (err) {
    Logger.log('Error al guardar configuración: ' + err);
    return { exito: false, mensaje: 'Error al guardar la configuración: ' + err.message };
  } finally {
    candado.releaseLock();
  }
}
/**
 * ============================================================
 *  VALIDACIÓN DE FORMATO (correo y teléfono)
 * ============================================================
 */

/**
 * Valida el formato de un correo electrónico. Cadena vacía = válido
 * (el campo es opcional).
 */
function _emailValido_(email) {
  email = String(email || '').trim();
  if (!email) return true;
  return /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(email);
}

/**
 * Códigos de área válidos para números dominicanos (fijos y móviles).
 */
var CODIGOS_AREA_RD_ = ['809', '829', '849'];

/**
 * Valida el formato de un número de teléfono dominicano. Cadena vacía = válido.
 * Debe tener exactamente 10 dígitos e iniciar con un código de área
 * válido (809, 829 u 849).
 */
function _telefonoValido_(telefono) {
  telefono = String(telefono || '').trim();
  if (!telefono) return true;
  if (!/^[0-9\s\-\(\)]+$/.test(telefono)) return false;

  var digitos = telefono.replace(/\D/g, '');
  if (digitos.length !== 10) return false;
  return CODIGOS_AREA_RD_.indexOf(digitos.slice(0, 3)) !== -1;
}

/**
 * Verifica que existan todas las hojas requeridas y que sus encabezados
 * sean correctos. Callable desde el menú y desde la aplicación web.
 * NO modifica datos; solo informa.
 * @return {Object} estado detallado de la base de datos.
 */
function verificarBaseDatos(token) {
  if (!_esContextoEditor_() && !_tieneHojaActiva_() && !_validarSesion_(token)) {
    return _respuestaSesionExpirada_();
  }
  var resultado = { exito: true, hojas: {}, mensaje: '' };
  try {
    var hojaCalculo = obtenerHojaCalculo_();

    // Hojas de datos.
    Object.keys(ENCABEZADOS).forEach(function(nombreHoja) {
      var esperado = ENCABEZADOS[nombreHoja];
      var hoja = hojaCalculo.getSheetByName(nombreHoja);
      var info = { existe: !!hoja, encabezadosCorrectos: false, filasDatos: 0 };
      if (hoja) {
        var ancho = Math.max(hoja.getLastColumn(), esperado.length);
        var fila = hoja.getRange(1, 1, 1, ancho).getValues()[0];
        var ok = true;
        for (var k = 0; k < esperado.length; k++) {
          if (String(fila[k] || '') !== esperado[k]) { ok = false; break; }
        }
        info.encabezadosCorrectos = ok;
        info.filasDatos = Math.max(0, hoja.getLastRow() - 1);
      } else {
        resultado.exito = false;
      }
      if (!info.encabezadosCorrectos) resultado.exito = false;
      resultado.hojas[nombreHoja] = info;
    });

    // Hoja de configuración.
    var hojaCfg = hojaCalculo.getSheetByName(HOJA_CONFIGURACION);
    resultado.hojas[HOJA_CONFIGURACION] = {
      existe: !!hojaCfg,
      encabezadosCorrectos: !!hojaCfg,
      filasDatos: hojaCfg ? Math.max(0, hojaCfg.getLastRow() - 1) : 0
    };
    if (!hojaCfg) resultado.exito = false;

    resultado.idHojaCalculo = hojaCalculo.getId();
    resultado.mensaje = resultado.exito
      ? 'La base de datos está completa y con los encabezados correctos.'
      : 'Se detectaron problemas. Ejecute "Iniciar configuración" para repararlos.';
  } catch (err) {
    resultado.exito = false;
    resultado.mensaje = 'Error al verificar la base de datos: ' + err.message;
  }

  // Si se ejecuta desde el menú (hay interfaz), mostrar un resumen.
  _mostrarResumenVerificacion_(resultado);
  return resultado;
}

/**
 * Muestra (si hay interfaz de hoja de cálculo) un resumen de la
 * verificación de la base de datos.
 */
function _mostrarResumenVerificacion_(resultado) {
  var lineas = [resultado.mensaje, ''];
  Object.keys(resultado.hojas).forEach(function(nombre) {
    var h = resultado.hojas[nombre];
    var estado = h.existe ? (h.encabezadosCorrectos ? 'OK' : 'ENCABEZADOS') : 'FALTA';
    lineas.push('• ' + nombre + ': ' + estado + ' (' + (h.filasDatos || 0) + ' registros)');
  });
  var texto = lineas.join('\n');
  // Registrar el reporte con console.log (aparece en Ejecuciones → Registros de Cloud)
  // y con Logger.log (aparece en Ver → Registros).
  console.log('Verificar base de datos:\n' + texto);
  Logger.log('Verificar base de datos:\n' + texto);

  // Mostrar por UI si hay interfaz; si falla, usar "toast"; si tampoco, el
  // Logger ya tiene el reporte completo.
  try {
    SpreadsheetApp.getUi().alert('Verificar base de datos', texto, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  } catch (errUi) {
    try {
      var hojaCalculo = SpreadsheetApp.getActiveSpreadsheet();
      if (hojaCalculo) {
        hojaCalculo.toast(resultado.mensaje, 'Verificar base de datos', 10);
      }
    } catch (errToast) {
      // Sin interfaz: el reporte quedó en el Logger.
    }
  }
}

/**
 * Devuelve la URL de la aplicación web desplegada (/exec) o cadena vacía
 * si aún no se ha publicado.
 * @return {string}
 */
function obtenerUrlAplicacion() {
  try {
    var url = ScriptApp.getService().getUrl();
    return url || '';
  } catch (err) {
    Logger.log('No se pudo obtener la URL de la aplicación: ' + err);
    return '';
  }
}

/**
 * Genera un identificador único basado en la marca de tiempo.
 * @param {string} prefijo  Prefijo opcional (ej. "CLI", "CITA").
 */
function generarId(prefijo) {
  var marca = new Date().getTime();
  var aleatorio = Math.floor(Math.random() * 1000);
  var base = String(marca) + String(aleatorio);
  return prefijo ? (prefijo + '-' + base) : base;
}

/**
 * Convierte las filas de una hoja en un arreglo de objetos
 * usando la primera fila como claves.
 */
function filasAObjetos_(hoja) {
  return _filasAObjetosDesdeValores_(hoja.getDataRange().getValues());
}

/**
 * Convierte un arreglo 2D (con encabezados en la fila 0) a objetos,
 * normalizando fechas a texto legible y anotando la fila real (_fila).
 * @param {Array} datos  Matriz con encabezados en [0].
 * @return {Array} objetos.
 */
function _filasAObjetosDesdeValores_(datos) {
  if (datos.length < 2) {
    return [];
  }
  var encabezados = datos[0];
  var zona = obtenerZonaHoraria_(); // Fuera del bucle: evita llamadas por celda.
  var resultado = [];
  for (var i = 1; i < datos.length; i++) {
    var fila = datos[i];
    // Saltar filas totalmente vacías.
    if (fila.join('') === '') {
      continue;
    }
    var obj = {};
    for (var j = 0; j < encabezados.length; j++) {
      var valor = fila[j];
      // Normalizar fechas a texto legible.
      if (valor instanceof Date) {
        valor = Utilities.formatDate(valor, zona, 'yyyy-MM-dd HH:mm');
      }
      obj[encabezados[j]] = valor;
    }
    obj._fila = i + 1; // Número de fila real en la hoja (1-indexado).
    resultado.push(obj);
  }
  return resultado;
}

/**
 * Busca el número de fila de un registro por su ID.
 * @param {Sheet} hoja        Hoja donde buscar.
 * @param {string} nombreCol  Nombre de la columna del ID.
 * @param {string} id         Valor a buscar.
 * @return {number}           Número de fila (0 si no se encuentra).
 */
function buscarFilaPorId_(hoja, nombreCol, id) {
  var ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) {
    return 0;
  }
  var encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
  var indiceCol = encabezados.indexOf(nombreCol);
  if (indiceCol === -1) {
    return 0;
  }
  // Leer solo la columna del ID (evita traer toda la matriz para hallar una fila).
  var columna = hoja.getRange(2, indiceCol + 1, ultimaFila - 1, 1).getValues();
  for (var i = 0; i < columna.length; i++) {
    if (String(columna[i][0]) === String(id)) {
      return i + 2;
    }
  }
  return 0;
}

function obtenerZonaHoraria_() {
  return Session.getScriptTimeZone() || 'America/Bogota';
}

/**
 * ============================================================
 *  VALIDACIONES
 * ============================================================
 */

/**
 * Valida el formato de un correo electrónico.
 * @param {string} email
 * @return {boolean}
 */
function esEmailValido_(email) {
  email = String(email || '').trim();
  if (!email) return false;
  // Regex estándar, sin permitir espacios ni caracteres inválidos.
  return /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(email);
}

/**
 * Valida un número de teléfono contra el patrón configurado
 * (por defecto: internacional flexible). Si no hay patrón válido,
 * acepta cualquier cadena no vacía.
 * @param {string} telefono
 * @return {boolean}
 */
function esTelefonoValido_(telefono) {
  telefono = String(telefono || '').trim();
  if (!telefono) return false;
  try {
    // Patrón fijo de teléfono (no configurable).
    var regex = new RegExp('^[+0-9 ()-]{7,20}$');
    return regex.test(telefono);
  } catch (err) {
    return telefono.length > 0;
  }
}

/**
 * Busca si un valor ya existe en una columna de una hoja, ignorando
 * la fila indicada (para no chocar con el registro en edición).
 * @param {Sheet} hoja      Hoja donde buscar.
 * @param {string} nombreCol Nombre de la columna.
 * @param {string} valor     Valor a buscar.
 * @param {number} exceptoFila Fila a ignorar (0 = ninguna).
 * @return {boolean} true si ya existe.
 */
function existeDuplicado_(hoja, nombreCol, valor, exceptoFila) {
  valor = String(valor || '').trim().toLowerCase();
  if (!valor) return false;
  var ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return false;
  var encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
  var indiceCol = encabezados.indexOf(nombreCol);
  if (indiceCol === -1) return false;
  // Leer solo la columna consultada (evita traer toda la matriz por cada chequeo).
  var columna = hoja.getRange(2, indiceCol + 1, ultimaFila - 1, 1).getValues();
  for (var i = 0; i < columna.length; i++) {
    if (exceptoFila && (i + 2) === exceptoFila) continue;
    if (String(columna[i][0] || '').trim().toLowerCase() === valor) {
      return true;
    }
  }
  return false;
}

/**
 * ============================================================
 *  CRUD — CLIENTES
 * ============================================================
 */

/**
 * Sube una imagen (data URL base64) a una carpeta de Drive de la
 * aplicación y devuelve una URL pública para mostrarla en <img>.
 * @param {string} dataUrl    Ej. "data:image/png;base64,..."
 * @param {string} nombreBase Nombre sugerido del archivo (opcional).
 * @return {string} URL pública de la imagen.
 */
function _guardarImagenDesdeBase64_(dataUrl, nombreBase) {
  if (!dataUrl) return '';
  var mime = dataUrl.match(/^data:([a-zA-Z0-9/.-]+);base64,(.+)$/);
  if (!mime) {
    throw new Error('Formato de imagen no válido.');
  }
  var tipo = mime[1]; // image/png, image/jpeg, image/gif, image/webp, ...
  var binario = Utilities.base64Decode(mime[2]);
  var extension = 'img';
  if (tipo === 'image/png') extension = 'png';
  else if (tipo === 'image/jpeg') extension = 'jpg';
  else if (tipo === 'image/gif') extension = 'gif';
  else if (tipo === 'image/webp') extension = 'webp';

  var base = String(nombreBase || 'imagen').replace(/[^a-zA-Z0-9-_ ]/g, '').replace(/\s+/g, '_') || 'imagen';
  var nombre = base + '_' + new Date().getTime() + '.' + extension;
  var carpeta = obtenerCarpetaImagenes_();
  var blob = Utilities.newBlob(binario, tipo, nombre);
  var archivo = carpeta.createFile(blob);
  archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return 'https://drive.google.com/thumbnail?id=' + archivo.getId() + '&sz=w1000';
}

/** Devuelve (o crea) la carpeta de Drive donde se guardan las imágenes. */
function obtenerCarpetaImagenes_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty(PROP_CARPETA_IMAGENES);
  if (id) {
    try {
      return DriveApp.getFolderById(id);
    } catch (err) {
      // La carpeta ya no existe: se borra el ID para crearla de nuevo.
      props.deleteProperty(PROP_CARPETA_IMAGENES);
    }
  }
  var carpeta = DriveApp.createFolder('ClientManager_Imagenes');
  props.setProperty(PROP_CARPETA_IMAGENES, carpeta.getId());
  return carpeta;
}

function agregarCliente(token, datos) {
  try {
    if (!_validarSesion_(token)) return _respuestaSesionExpirada_();
    // Validar email si se proporciona.
    if (datos.email && !esEmailValido_(datos.email)) {
      return { exito: false, mensaje: 'El correo electrónico no tiene un formato válido.' };
    }
    // Validar teléfono si se proporciona.
    if (datos.telefono && !esTelefonoValido_(datos.telefono)) {
      return { exito: false, mensaje: 'El teléfono no tiene un formato válido.' };
    }
    var hoja = obtenerHoja_(HOJA_CLIENTES);
    // Detectar duplicados.
    if (datos.email && existeDuplicado_(hoja, 'Email', datos.email, 0)) {
      return { exito: false, mensaje: 'Ya existe un cliente con ese correo electrónico.' };
    }
    if (datos.telefono && existeDuplicado_(hoja, 'Telefono', datos.telefono, 0)) {
      return { exito: false, mensaje: 'Ya existe un cliente con ese teléfono.' };
    }
    // La foto puede venir como data URL (se sube a Drive aquí mismo, en la
    // misma llamada) o como URL ya existente.
    var fotoUrl = String(datos.foto || '');
    if (datos.fotoDataUrl) {
      try {
        fotoUrl = _guardarImagenDesdeBase64_(datos.fotoDataUrl,
          (datos.nombre || '') + ' ' + (datos.apellido || ''));
      } catch (errFoto) {
        return { exito: false, mensaje: 'No se pudo subir la imagen: ' + errFoto.message };
      }
    }
    var id = generarId('CLI');
    var fechaRegistro = Utilities.formatDate(new Date(), obtenerZonaHoraria_(), 'yyyy-MM-dd HH:mm');
    hoja.appendRow([
      id,
      datos.nombre || '',
      datos.apellido || '',
      datos.telefono || '',
      datos.email || '',
      datos.direccion || '',
      datos.notas || '',
      fechaRegistro,
      fotoUrl
    ]);
    Logger.log('Cliente agregado: ' + id);
    _registrarActividad_(token, 'Clientes', 'Agregó cliente', ((datos.nombre || '') + ' ' + (datos.apellido || '')).trim() + (datos.email ? ' (' + datos.email + ')' : ''));
    return { exito: true, mensaje: 'Cliente agregado correctamente.', id: id, foto: fotoUrl, fechaRegistro: fechaRegistro };
  } catch (err) {
    Logger.log('Error al agregar cliente: ' + err);
    return { exito: false, mensaje: 'Error al agregar cliente: ' + err.message };
  }
}

/**
 * Devuelve todos los clientes.
 */
function obtenerClientes(token) {
  if (!_validarSesion_(token)) return _respuestaSesionExpirada_();
  var hoja = obtenerHoja_(HOJA_CLIENTES);
  return filasAObjetos_(hoja);
}

/**
 * Actualiza un cliente existente por su ID.
 */
function actualizarCliente(token, id, datos) {
  try {
    if (!_validarSesion_(token)) return _respuestaSesionExpirada_();
    if (datos.email && !esEmailValido_(datos.email)) {
      return { exito: false, mensaje: 'El correo electrónico no tiene un formato válido.' };
    }
    if (datos.telefono && !esTelefonoValido_(datos.telefono)) {
      return { exito: false, mensaje: 'El teléfono no tiene un formato válido.' };
    }
    var hoja = obtenerHoja_(HOJA_CLIENTES);
    var fila = buscarFilaPorId_(hoja, 'ID_Cliente', id);
    if (!fila) {
      return { exito: false, mensaje: 'No se encontró el cliente indicado.' };
    }
    // Detectar duplicados (ignorando la fila del propio cliente).
    if (datos.email && existeDuplicado_(hoja, 'Email', datos.email, fila)) {
      return { exito: false, mensaje: 'Ya existe otro cliente con ese correo electrónico.' };
    }
    if (datos.telefono && existeDuplicado_(hoja, 'Telefono', datos.telefono, fila)) {
      return { exito: false, mensaje: 'Ya existe otro cliente con ese teléfono.' };
    }
    // La foto puede venir como data URL (se sube a Drive aquí mismo) o como
    // URL ya existente.
    var fotoUrl = String(datos.foto || '');
    if (datos.fotoDataUrl) {
      try {
        fotoUrl = _guardarImagenDesdeBase64_(datos.fotoDataUrl,
          (datos.nombre || '') + ' ' + (datos.apellido || ''));
      } catch (errFoto) {
        return { exito: false, mensaje: 'No se pudo subir la imagen: ' + errFoto.message };
      }
    }
    hoja.getRange(fila, 2, 1, 6).setValues([[
      datos.nombre || '',
      datos.apellido || '',
      datos.telefono || '',
      datos.email || '',
      datos.direccion || '',
      datos.notas || ''
    ]]);
    var colFotoCliente = ENCABEZADOS.Clientes.indexOf('Foto') + 1;
    if (colFotoCliente > 0) {
      hoja.getRange(fila, colFotoCliente).setValue(fotoUrl);
    }
    Logger.log('Cliente actualizado: ' + id);
    _registrarActividad_(token, 'Clientes', 'Editó cliente', 'ID ' + id + (datos.email ? ' (' + datos.email + ')' : ''));
    return { exito: true, mensaje: 'Cliente actualizado correctamente.', foto: fotoUrl };
  } catch (err) {
    Logger.log('Error al actualizar cliente: ' + err);
    return { exito: false, mensaje: 'Error al actualizar cliente: ' + err.message };
  }
}

/**
 * Elimina un cliente por su ID.
 */
function eliminarCliente(token, id) {
  try {
    var usuario = _validarSesion_(token);
    if (!usuario) return _respuestaSesionExpirada_();
    if (!_esAdmin_(usuario)) return { exito: false, mensaje: 'Solo el administrador puede eliminar clientes.' };
    var hoja = obtenerHoja_(HOJA_CLIENTES);
    var fila = buscarFilaPorId_(hoja, 'ID_Cliente', id);
    if (!fila) {
      return { exito: false, mensaje: 'No se encontró el cliente indicado.' };
    }
    hoja.deleteRow(fila);
    Logger.log('Cliente eliminado: ' + id);
    _registrarActividad_(token, 'Clientes', 'Eliminó cliente', 'ID ' + id);
    return { exito: true, mensaje: 'Cliente eliminado correctamente.' };
  } catch (err) {
    Logger.log('Error al eliminar cliente: ' + err);
    return { exito: false, mensaje: 'Error al eliminar cliente: ' + err.message };
  }
}

/**
 * ============================================================
 *  SERVICIOS (catálogo de servicios del negocio)
 * ============================================================
 *  La vista "Servicios" permite al ADMINISTRADOR crear, editar y
 *  eliminar el catálogo. Los servicios se ofrecen en la página
 *  pública de reservas en línea, donde el cliente elige cantidades
 *  y el sistema calcula el total. Solo se exponen públicamente los
 *  servicios con Activo = "SI".
 */

/**
 * Valida los campos de un servicio y los devuelve normalizados.
 * @param {Object} datos  { nombre, precio, duracionMins, descripcion, activo }
 * @return {Object|null}  { nombre, precio, duracionMins, descripcion, activo } o null si es inválido.
 */
function _validarServicio_(datos) {
  datos = datos || {};
  var nombre = String(datos.nombre || '').trim();
  if (!nombre) return null;
  var precio = parseFloat(String(datos.precio || '0').replace(/[^\d.,-]/g, '').replace(',', '.'));
  if (isNaN(precio) || precio < 0) precio = 0;
  var duracion = parseInt(datos.duracionMins, 10);
  if (isNaN(duracion) || duracion < 1) duracion = 60;
  var descripcion = String(datos.descripcion || '').trim();
  var activo = String(datos.activo || '').toUpperCase() === 'NO' ? 'NO' : 'SI';
  return { nombre: nombre, precio: precio, duracionMins: duracion, descripcion: descripcion, activo: activo };
}

/**
 * Devuelve todos los servicios (requiere sesión válida).
 * Los editores pueden consultarlos; solo el admin los gestiona.
 * @param {string} token
 * @return {Array|Object}
 */
function obtenerServicios(token) {
  if (!_validarSesion_(token)) return _respuestaSesionExpirada_();
  try {
    return filasAObjetos_(obtenerHoja_(HOJA_SERVICIOS));
  } catch (err) {
    Logger.log('Error en obtenerServicios: ' + err);
    return { exito: false, mensaje: 'Error al obtener los servicios: ' + err.message };
  }
}

/**
 * Crea un servicio nuevo (solo administrador).
 * @param {string} token
 * @param {Object} datos
 * @return {Object}
 */
function agregarServicio(token, datos) {
  try {
    var usuario = _validarSesion_(token);
    if (!usuario) return _respuestaSesionExpirada_();
    if (!_esAdmin_(usuario)) return { exito: false, mensaje: 'Solo el administrador puede gestionar los servicios.' };
    var s = _validarServicio_(datos);
    if (!s) return { exito: false, mensaje: 'El nombre del servicio es obligatorio.' };

    var hoja = obtenerHoja_(HOJA_SERVICIOS);
    var id = generarId('SER');
    hoja.appendRow([id, s.nombre, s.precio, s.duracionMins, s.descripcion, s.activo]);
    Logger.log('Servicio creado: ' + id);
    _registrarActividad_(token, 'Servicios', 'Agregó servicio', (s.nombre || '') + (s.precio != null ? ' ($' + s.precio + ')' : ''));
    return { exito: true, mensaje: 'Servicio creado correctamente.', id: id };
  } catch (err) {
    Logger.log('Error al crear servicio: ' + err);
    return { exito: false, mensaje: 'Error al crear el servicio: ' + err.message };
  }
}

/**
 * Actualiza un servicio existente (solo administrador).
 * @param {string} token
 * @param {string} id   ID_Servicio.
 * @param {Object} datos
 * @return {Object}
 */
function actualizarServicio(token, id, datos) {
  try {
    var usuario = _validarSesion_(token);
    if (!usuario) return _respuestaSesionExpirada_();
    if (!_esAdmin_(usuario)) return { exito: false, mensaje: 'Solo el administrador puede gestionar los servicios.' };
    var s = _validarServicio_(datos);
    if (!s) return { exito: false, mensaje: 'El nombre del servicio es obligatorio.' };

    var hoja = obtenerHoja_(HOJA_SERVICIOS);
    var fila = buscarFilaPorId_(hoja, 'ID_Servicio', id);
    if (!fila) return { exito: false, mensaje: 'No se encontró el servicio indicado.' };
    hoja.getRange(fila, 2, 1, 5).setValues([[
      s.nombre, s.precio, s.duracionMins, s.descripcion, s.activo
    ]]);
    Logger.log('Servicio actualizado: ' + id);
    _registrarActividad_(token, 'Servicios', 'Editó servicio', 'ID ' + id);
    return { exito: true, mensaje: 'Servicio actualizado correctamente.' };
  } catch (err) {
    Logger.log('Error al actualizar servicio: ' + err);
    return { exito: false, mensaje: 'Error al actualizar el servicio: ' + err.message };
  }
}

/**
 * Elimina un servicio por su ID (solo administrador).
 * @param {string} token
 * @param {string} id   ID_Servicio.
 * @return {Object}
 */
function eliminarServicio(token, id) {
  try {
    var usuario = _validarSesion_(token);
    if (!usuario) return _respuestaSesionExpirada_();
    if (!_esAdmin_(usuario)) return { exito: false, mensaje: 'Solo el administrador puede gestionar los servicios.' };
    var hoja = obtenerHoja_(HOJA_SERVICIOS);
    var fila = buscarFilaPorId_(hoja, 'ID_Servicio', id);
    if (!fila) return { exito: false, mensaje: 'No se encontró el servicio indicado.' };
    hoja.deleteRow(fila);
    Logger.log('Servicio eliminado: ' + id);
    _registrarActividad_(token, 'Servicios', 'Eliminó servicio', 'ID ' + id);
    return { exito: true, mensaje: 'Servicio eliminado correctamente.' };
  } catch (err) {
    Logger.log('Error al eliminar servicio: ' + err);
    return { exito: false, mensaje: 'Error al eliminar el servicio: ' + err.message };
  }
}

/**
 * Devuelve los servicios ACTIVOS para la página pública de reservas.
 * Sin token. No expone campos sensibles.
 * @return {Object} { exito, servicios: [{ id, nombre, precio, duracionMins, descripcion }] }
 */
function obtenerServiciosPublicos() {
  try {
    var hoja = obtenerHoja_(HOJA_SERVICIOS);
    var valores = hoja.getDataRange().getValues();
    var cab = valores[0];
    var colId = cab.indexOf('ID_Servicio');
    var colNombre = cab.indexOf('Nombre');
    var colPrecio = cab.indexOf('Precio');
    var colDur = cab.indexOf('Duracion_Mins');
    var colDesc = cab.indexOf('Descripcion');
    var colActivo = cab.indexOf('Activo');
    if (colId < 0 || colNombre < 0) return { exito: true, servicios: [] };
    var servicios = [];
    for (var i = 1; i < valores.length; i++) {
      var fila = valores[i];
      if (colActivo >= 0 && String(fila[colActivo] || 'SI').toUpperCase() === 'NO') continue;
      if (!String(fila[colNombre] || '').trim()) continue;
      servicios.push({
        id: String(fila[colId] || ''),
        nombre: String(fila[colNombre] || '').trim(),
        precio: parseFloat(fila[colPrecio]) || 0,
        duracionMins: parseInt(fila[colDur], 10) || 60,
        descripcion: colDesc >= 0 ? String(fila[colDesc] || '').trim() : ''
      });
    }
    return { exito: true, servicios: servicios };
  } catch (err) {
    Logger.log('Error en obtenerServiciosPublicos: ' + err);
    return { exito: false, mensaje: 'Error al obtener los servicios: ' + err.message };
  }
}

/**
 * ============================================================
 *  USUARIOS Y ACCESO
 * ============================================================
 *  Sistema de acceso por correo electrónico y contraseña almacenados
 *  en la hoja "Usuarios". Roles: 'administrador' y 'editor'.
 *  Todos pueden registrarse y editar; el PRIMER usuario registrado
 *  queda como administrador.
 */

/** Devuelve el hash SHA-256 de un texto (para contraseñas). */
function _hashTexto_(texto) {
  var bytes = Utilities.newBlob(String(texto || '')).getBytes();
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, bytes);
  return digest.map(function(b) {
    return ('0' + ((b + 256) % 256).toString(16)).slice(-2);
  }).join('');
}

/** Genera un salt aleatorio (16 caracteres hex) para contraseñas. */
function _generarSalt_() {
  var bytes = [];
  for (var i = 0; i < 8; i++) {
    bytes.push(Math.floor(Math.random() * 256));
  }
  return bytes.map(function(b) {
    return ('0' + b.toString(16)).slice(-2);
  }).join('');
}

/** Verifica una contraseña contra Salt+Hash (o hash legado sin salt). */
function _verificarContrasena_(usuario, contrasena) {
  var salt = String(usuario.Salt || '');
  var p = String(contrasena || '');
  if (salt) {
    return _hashTexto_(salt + p) === String(usuario.Hash || '');
  }
  return _hashTexto_(p) === String(usuario.Hash || '');
}

/** Devuelve la fecha/hora actual en formato yyyy-MM-dd HH:mm. */
function _fechaRegistro_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
}

/** Lee los usuarios de la hoja como objetos. */
function _usuariosComoObjetos_() {
  var hoja = obtenerHoja_(HOJA_USUARIOS);
  return filasAObjetos_(hoja);
}

function _correosDestinatariosReserva_() {
  var lista = _usuariosComoObjetos_();
  var correos = [];
  for (var i = 0; i < lista.length; i++) {
    var u = lista[i] || {};
    if (String(u.Activo || 'SI').toUpperCase() !== 'SI') continue;
    var cor = String(u.Email || '').trim();
    if (!cor) continue;
    var clave = cor.toLowerCase();
    var duplicado = false;
    for (var j = 0; j < correos.length; j++) {
      if (String(correos[j]).toLowerCase() === clave) { duplicado = true; break; }
    }
    if (!duplicado) correos.push(cor);
  }
  return correos;
}

/** Devuelve el correo del dueño (propietario) de la hoja de cálculo. */
function _correoDueno_() {
  try {
    var hoja = obtenerHojaCalculo_();
    var dueno = hoja.getOwner();
    return dueno ? (dueno.getEmail() || '') : '';
  } catch (e) {
    return '';
  }
}

/**
 * Devuelve el correo del administrador principal (dueño de la app):
 * el usuario con Principal === 'SI' (o el primer administrador) en la hoja
 * Usuarios. Si no se encuentra, cae al dueño de la hoja de cálculo.
 */
function _correoPrincipalAdmin_() {
  try {
    var lista = _usuariosComoObjetos_();
    for (var i = 0; i < lista.length; i++) {
      if (String(lista[i].Rol || '').toLowerCase() === 'administrador' &&
          String(lista[i].Principal || '').toUpperCase() === 'SI') {
        var c = String(lista[i].Email || '').trim();
        if (c) return c;
      }
    }
    for (var j = 0; j < lista.length; j++) {
      if (String(lista[j].Rol || '').toLowerCase() === 'administrador') {
        var e = String(lista[j].Email || '').trim();
        if (e) return e;
      }
    }
  } catch (e) {}
  return _correoDueno_();
}

/**
 * Destinatarios del aviso de una nueva cita/reserva.
 * REGLA: SOLO el dueño de la hoja y, si es distinto, quien la agendó
 * (usuario interno). NUNCA se notifica a todos los usuarios.
 * @param {string} emailAgendador  Correo de quien agenda (puede ser vacío).
 * @return {Array} lista de correos.
 */
function _correosAvisoCita_(emailAgendador) {
  var lista = [];
  var dueno = _correoPrincipalAdmin_();
  if (dueno) lista.push(dueno);
  if (emailAgendador) {
    var e = String(emailAgendador).trim().toLowerCase();
    if (e && e !== String(dueno).toLowerCase()) lista.push(emailAgendador);
  }
  return lista;
}

/** Busca un usuario por su correo (columna Email o ID_Usuario). */
function _buscarUsuario_(usuario) {
  var lista = _usuariosComoObjetos_();
  var clave = String(usuario || '').toLowerCase();
  for (var i = 0; i < lista.length; i++) {
    var idUsr = String(lista[i].Email || lista[i].ID_Usuario || '').toLowerCase();
    if (idUsr === clave) {
      return lista[i];
    }
  }
  return null;
}

/**
 * Repara la hoja de usuarios:
 *  - Si NO existe ningún administrador, convierte al PRIMER usuario
 *    registrado (fila más antigua) en administrador.
 *  - Los usuarios con rol vacío quedan como 'editor'.
 *  - Los usuarios con estado vacio quedan activos ('SI').
 * Se invoca en el login y al listar usuarios.
 */
function _repararRolesUsuarios_() {
  // Reparación estructural: solo se ejecuta una vez cada 15 minutos. Evita
  // releer/escribir la hoja Usuarios en cada login o listado.
  try {
    if (CacheService.getScriptCache().get('CHK_ROLES')) return;
  } catch (err) {
    // Si el caché falla, se sigue con la verificación normal.
  }
  var lista = _usuariosComoObjetos_();
  if (lista.length === 0) return;
  var hoja = obtenerHoja_(HOJA_USUARIOS);
  var colRol = ENCABEZADOS.Usuarios.indexOf('Rol') + 1;
  var colActivo = ENCABEZADOS.Usuarios.indexOf('Activo') + 1;
  var colFecha = ENCABEZADOS.Usuarios.indexOf('Fecha_Registro') + 1;

  var hayAdmin = false;
  for (var i = 0; i < lista.length; i++) {
    if (String(lista[i].Rol || '').toLowerCase() === 'administrador' &&
        String(lista[i].Activo).toUpperCase() !== 'NO') {
      hayAdmin = true;
      break;
    }
  }
  if (!hayAdmin) {
    hoja.getRange(lista[0]._fila, colRol).setValue('administrador');
    Logger.log('Rol administrador asignado al primer usuario: ' + (lista[0].ID_Usuario || lista[0].Email));
  }
  for (var j = 0; j < lista.length; j++) {
    var rolActual = String(lista[j].Rol || '').trim();
    var rolMin = rolActual.toLowerCase();
    if (rolMin !== 'administrador' && rolMin !== 'editor') {
      hoja.getRange(lista[j]._fila, colRol).setValue('editor');
    } else if (rolActual !== rolMin) {
      hoja.getRange(lista[j]._fila, colRol).setValue(rolMin);
    }
    var activo = String(lista[j].Activo || '').trim().toUpperCase();
    if (activo !== 'SI' && activo !== 'NO') {
      hoja.getRange(lista[j]._fila, colActivo).setValue('SI');
    }
    if (!String(lista[j].Fecha_Registro || '').trim()) {
      hoja.getRange(lista[j]._fila, colFecha).setValue(_fechaRegistro_());
    }
  }
  try {
    CacheService.getScriptCache().put('CHK_ROLES', '1', 900);
  } catch (err) {}
}

/** Devuelve true si el usuario existe y tiene rol administrador. */
function _esAdmin_(usuario) {
  var u = _buscarUsuario_(usuario);
  return !!u && String(u.Rol) === 'administrador' && String(u.Activo).toUpperCase() !== 'NO';
}

/**
 * Devuelve true si el usuario es el administrador principal
 * (el PRIMER administrador registrado: el dueño). Ese usuario
 * no puede eliminarse, quitarse su rol ni desactivarse.
 */
function _esUsuarioPrincipal_(usuario) {
  var lista = _usuariosComoObjetos_();
  var clave = String(usuario || '').toLowerCase();
  for (var i = 0; i < lista.length; i++) {
    var id = String(lista[i].Email || lista[i].ID_Usuario || '').toLowerCase();
    if (String(lista[i].Rol || '').toLowerCase() === 'administrador') {
      return id === clave;
    }
  }
  return false;
}

/**
 * ============================================================
 *  SESIONES (tokens de acceso)
 *  Cada token se guarda en PropertiesService con vencimiento.
 *  Todas las funciones de datos/admin exigen un token válido.
 * ============================================================
 */

/** Lee el mapa de sesiones activas { token: { usuario, expira } }. */
function _leerSesiones_() {
  var cache = CacheService.getScriptCache();
  try {
    var crudo = cache.get(PROP_SESIONES);
    if (!crudo) {
      crudo = PropertiesService.getScriptProperties().getProperty(PROP_SESIONES);
      if (crudo) cache.put(PROP_SESIONES, crudo, CACHE_SESIONES_TTL_SEG);
    }
    return crudo ? JSON.parse(crudo) : {};
  } catch (err) {
    return {};
  }
}

/** Guarda el mapa de sesiones activas (en Properties y en caché). */
function _guardarSesiones_(mapa) {
  var json = JSON.stringify(mapa);
  PropertiesService.getScriptProperties().setProperty(PROP_SESIONES, json);
  CacheService.getScriptCache().put(PROP_SESIONES, json, CACHE_SESIONES_TTL_SEG);
}

/** Crea una sesión para el usuario y devuelve el token. */
function _crearSesion_(usuario) {
  var mapa = _leerSesiones_();
  var token = Utilities.getUuid();
  mapa[token] = { usuario: String(usuario), expira: new Date().getTime() + DURACION_SESION_MS };
  _guardarSesiones_(mapa);
  return token;
}

/**
 * Valida un token de sesión. Devuelve el usuario (correo) si el token
 * es válido y no ha expirado; si expiró, lo elimina y devuelve null.
 * @param {string} token
 * @return {string|null}
 */
function _validarSesion_(token) {
  var t = String(token || '').trim();
  if (!t) return null;
  var mapa = _leerSesiones_();
  if (!mapa[t]) return null;
  var ahora = new Date().getTime();
  if (ahora > mapa[t].expira) {
    delete mapa[t];
    _guardarSesiones_(mapa);
    return null;
  }
  return mapa[t].usuario;
}

/** Elimina una sesión (cierre de sesión). */
function _cerrarSesion_(token) {
  var t = String(token || '').trim();
  if (!t) return;
  var mapa = _leerSesiones_();
  if (mapa[t]) {
    delete mapa[t];
    _guardarSesiones_(mapa);
  }
}

/**
 * Respuesta estándar cuando el token falta, es inválido o expiró.
 * El frontend detecta sesionExpirada y vuelve a mostrar el login.
 * @return {Object}
 */
function _respuestaSesionExpirada_() {
  return {
    exito: false,
    sesionExpirada: true,
    mensaje: 'Su sesión expiró. Vuelva a iniciar sesión.'
  };
}

/**
 * Comprueba si la llamada proviene del contexto del editor de la hoja
 * (menú / ejecución directa del propietario en Apps Script), donde
 * Session.getActiveUser() devuelve el correo del usuario autenticado.
 * En la aplicación web con acceso "cualquier usuario" devuelve vacío.
 * @return {boolean}
 */
function _esContextoEditor_() {
  try {
    var correo = Session.getActiveUser().getEmail();
    return !!correo;
  } catch (err) {
    return false;
  }
}

/**
 * Devuelve true si la ejecución tiene una hoja de cálculo activa
 * (contexto del contenedor: editor de Sheets o ejecución desde el
 * editor de Apps Script de una hoja vinculada). En la aplicación web
 * y en la API no existe hoja activa, por lo que esto distingue
 * claramente el uso "desde la hoja" del uso remoto.
 */
function _tieneHojaActiva_() {
  try {
    return !!SpreadsheetApp.getActiveSpreadsheet();
  } catch (err) {
    return false;
  }
}

/**
 * Función pública para cerrar sesión: invalida el token en el servidor.
 * @param {string} token
 * @return {Object}
 */
function cerrarSesion(token) {
  _cerrarSesion_(token);
  return { exito: true };
}

/**
 * Verifica credenciales y devuelve la sesión. El acceso se hace con
 * el correo electrónico (columna Usuario) y la contraseña.
 * Las cuentas se crean mediante registrarUsuario().
 * @param {string} usuario correo electrónico
 * @param {string} contrasena
 */
function verificarLogin(usuario, contrasena) {
  try {
    _repararRolesUsuarios_();
    var u = String(usuario || '').trim().toLowerCase();
    var p = String(contrasena || '');
    if (!u || !p) {
      return { exito: false, mensaje: 'Ingrese correo y contraseña.' };
    }
    var existentes = _usuariosComoObjetos_();
    if (existentes.length === 0) {
      // Aún no hay cuentas: hay que registrarse (la primera será admin).
      return { exito: false, mensaje: 'No hay cuentas registradas. Use "Registrarse" para crear la primera (será la de administrador).' };
    }
    var encontrado = null;
    for (var i = 0; i < existentes.length; i++) {
      var id = String(existentes[i].Email || existentes[i].ID_Usuario || '').toLowerCase();
      if (id === u) { encontrado = existentes[i]; break; }
    }
    if (!encontrado) {
      return { exito: false, mensaje: 'Usuario o contraseña incorrectos.' };
    }
    if (String(encontrado.Activo).toUpperCase() === 'NO') {
      return { exito: false, mensaje: 'Este usuario está desactivado.' };
    }
    if (!_verificarContrasena_(encontrado, p)) {
      return { exito: false, mensaje: 'Usuario o contraseña incorrectos.' };
    }
    var token = _crearSesion_(encontrado.Email || encontrado.ID_Usuario);
    return {
      exito: true, usuario: encontrado.Email || encontrado.ID_Usuario,
      nombre: encontrado.Nombre || encontrado.Email || encontrado.ID_Usuario,
      rol: encontrado.Rol || 'editor',
      principal: _esUsuarioPrincipal_(encontrado.Email || encontrado.ID_Usuario),
      token: token
    };
  } catch (err) {
    Logger.log('Error en verificarLogin: ' + err);
    return { exito: false, mensaje: 'Error al iniciar sesión: ' + err.message };
  }
}

/**
 * Registro abierto: el usuario crea su cuenta con su correo electrónico.
 * Todos quedan habilitados para ver y editar. El PRIMER usuario registrado
 * queda como 'administrador'; el resto como 'editor'.
 * Devuelve la sesión lista para entrar.
 * @param {string} nombre
 * @param {string} correo
 * @param {string} contrasena
 */
function registrarUsuario(nombre, correo, contrasena) {
  try {
    var nom = String(nombre || '').trim();
    var cor = String(correo || '').trim().toLowerCase();
    var p = String(contrasena || '');
    if (!nom) {
      return { exito: false, mensaje: 'Ingrese su nombre.' };
    }
    if (!esEmailValido_(cor)) {
      return { exito: false, mensaje: 'Ingrese un correo electrónico válido.' };
    }
    if (p.length < 8) {
      return { exito: false, mensaje: 'La contraseña debe tener al menos 8 caracteres.' };
    }
    if (_buscarUsuario_(cor)) {
      return { exito: false, mensaje: 'Ya existe una cuenta con ese correo.' };
    }

    var existentes = _usuariosComoObjetos_();
    var rol = existentes.length === 0 ? 'administrador' : 'editor';
    var primerAcceso = existentes.length === 0;
    var salt = _generarSalt_();
    var idUsuario = generarId('USR');
    obtenerHoja_(HOJA_USUARIOS).appendRow([
      idUsuario, nom, cor, salt, _hashTexto_(salt + p), rol, 'SI', _fechaRegistro_()
    ]);
    Logger.log('Nueva cuenta: ' + cor + ' (' + rol + ')');
    _registrarActividad_(null, 'Usuarios', 'Se registró', nom + ' (' + cor + ')', cor);
    var token = _crearSesion_(cor);
    return {
      exito: true, primerAcceso: primerAcceso, usuario: cor, nombre: nom, rol: rol,
      principal: _esUsuarioPrincipal_(cor),
      token: token,
      mensaje: primerAcceso ? 'Cuenta de administrador creada. Bienvenido.' : 'Cuenta creada. Bienvenido.'
    };
  } catch (err) {
    Logger.log('Error en registrarUsuario: ' + err);
    return { exito: false, mensaje: 'Error al crear la cuenta: ' + err.message };
  }
}

/** Lista los usuarios (sin contraseñas). Solo administrador. */
function listarUsuarios(token) {
  var usuarioSesion = _validarSesion_(token);
  if (!usuarioSesion) return _respuestaSesionExpirada_();
  _repararRolesUsuarios_();
  var lista = _usuariosComoObjetos_();

  // Comprobar permiso de administrador sobre la misma lista (sin releer la hoja).
  var esAdmin = false;
  var claveSesion = String(usuarioSesion || '').toLowerCase();
  for (var i = 0; i < lista.length; i++) {
    if (String(lista[i].Email || lista[i].ID_Usuario || '').toLowerCase() === claveSesion &&
        String(lista[i].Rol).toLowerCase() === 'administrador' &&
        String(lista[i].Activo).toUpperCase() !== 'NO') {
      esAdmin = true;
      break;
    }
  }
  if (!esAdmin) {
    return { exito: false, mensaje: 'No tiene permisos de administrador.' };
  }

  // El propietario es el PRIMER administrador registrado (fila más antigua).
  // Se resuelve en un solo pase, sin releer la hoja por cada usuario.
  var principalId = '';
  for (var j = 0; j < lista.length; j++) {
    if (String(lista[j].Rol || '').toLowerCase() === 'administrador') {
      principalId = String(lista[j].Email || lista[j].ID_Usuario || '').toLowerCase();
      break;
    }
  }

  var limpia = lista.map(function(u) {
    var login = String(u.Email || u.ID_Usuario || '');
    return {
      usuario: u.Email || u.ID_Usuario, idUsuario: u.ID_Usuario || '',
      nombre: u.Nombre || '', rol: u.Rol || 'editor',
      activo: String(u.Activo).toUpperCase() !== 'NO',
      fecha: String(u.Fecha_Registro || ''),
      principal: login.toLowerCase() === principalId
    };
  });
  return { exito: true, usuarios: limpia };
}

/** Lee la lista de usuarios ya avisados como "nuevos". @return {Array} */
function _usuariosYaAvisados_() {
  try {
    var crudo = PropertiesService.getScriptProperties().getProperty(PROP_USUARIOS_AVISADOS);
    var lista = crudo ? JSON.parse(crudo) : [];
    return Array.isArray(lista) ? lista : [];
  } catch (err) {
    return [];
  }
}

/** Guarda la lista de usuarios ya avisados como "nuevos". */
function _guardarUsuariosAvisados_(lista) {
  PropertiesService.getScriptProperties()
    .setProperty(PROP_USUARIOS_AVISADOS, JSON.stringify(lista || []));
}

/**
 * Devuelve los usuarios que aún no han sido avisados al administrador como
 * "nuevos" (para mostrar el aviso UNA sola vez por usuario). Excluye al
 * propio usuario de la sesión (no se avisa a sí mismo).
 * @param {string} token
 * @return {Object} { exito, usuarios: [{ usuario, nombre }] }
 */
function obtenerUsuariosNuevos(token) {
  try {
    var usuarioSesion = _validarSesion_(token);
    if (!usuarioSesion) return _respuestaSesionExpirada_();
    if (!_esAdmin_(usuarioSesion)) {
      return { exito: false, mensaje: 'No tiene permisos de administrador.' };
    }

    var lista = _usuariosComoObjetos_();
    var yaAvisados = _usuariosYaAvisados_();
    var clave = String(usuarioSesion || '').toLowerCase();
    var nuevos = [];
    for (var i = 0; i < lista.length; i++) {
      var login = String(lista[i].Email || lista[i].ID_Usuario || '');
      if (!login) continue;
      if (login.toLowerCase() === clave) continue;
      if (yaAvisados.indexOf(login.toLowerCase()) !== -1) continue;
      nuevos.push({ usuario: login, nombre: lista[i].Nombre || login });
    }
    return { exito: true, usuarios: nuevos };
  } catch (err) {
    Logger.log('Error en obtenerUsuariosNuevos: ' + err);
    return { exito: false, mensaje: 'Error al obtener los usuarios: ' + err.message };
  }
}

/**
 * Registra en PropertiesService los usuarios que ya fueron avisados como
 * "nuevos" (para que el aviso solo ocurra una vez por usuario).
 * @param {string} token
 * @param {Array} ids  Lista de usuario (email o ID) a marcar como avisados.
 * @return {Object}
 */
function marcarUsuariosAvisados(token, ids) {
  try {
    var usuarioSesion = _validarSesion_(token);
    if (!usuarioSesion) return _respuestaSesionExpirada_();
    if (!_esAdmin_(usuarioSesion)) {
      return { exito: false, mensaje: 'No tiene permisos de administrador.' };
    }
    var ya = _usuariosYaAvisados_();
    (ids || []).forEach(function(id) {
      var v = String(id || '').trim().toLowerCase();
      if (v && ya.indexOf(v) === -1) ya.push(v);
    });
    _guardarUsuariosAvisados_(ya);
    return { exito: true };
  } catch (err) {
    Logger.log('Error en marcarUsuariosAvisados: ' + err);
    return { exito: false, mensaje: 'Error al guardar el aviso: ' + err.message };
  }
}

/** Crea un usuario. Solo administrador. */
function crearUsuario(token, datos) {
  var usuarioSesion = _validarSesion_(token);
  if (!usuarioSesion) return _respuestaSesionExpirada_();
  if (!_esAdmin_(usuarioSesion)) {
    return { exito: false, mensaje: 'No tiene permisos de administrador.' };
  }
  var usuario = String(datos.usuario || '').trim().toLowerCase();
  var nombre = String(datos.nombre || '').trim();
  var contrasena = String(datos.contrasena || '');
  var rol = String(datos.rol || 'editor') === 'administrador' ? 'administrador' : 'editor';
  if (!usuario || !contrasena) {
    return { exito: false, mensaje: 'Correo y contraseña son obligatorios.' };
  }
  if (String(contrasena).length < 8) {
    return { exito: false, mensaje: 'La contraseña debe tener al menos 8 caracteres.' };
  }
  if (!esEmailValido_(usuario)) {
    return { exito: false, mensaje: 'Ingrese un correo electrónico válido.' };
  }
  if (_buscarUsuario_(usuario)) {
    return { exito: false, mensaje: 'Ya existe una cuenta con ese correo.' };
  }
  var salt = _generarSalt_();
  var idUsuario = generarId('USR');
  obtenerHoja_(HOJA_USUARIOS).appendRow([
    idUsuario, nombre || usuario, usuario, salt, _hashTexto_(salt + contrasena), rol, 'SI', _fechaRegistro_()
  ]);
    Logger.log('Usuario creado: ' + usuario + ' (' + rol + ')');
    _registrarActividad_(token, 'Usuarios', 'Creó usuario', usuario + ' (' + rol + ')');
    return { exito: true, mensaje: 'Usuario "' + usuario + '" creado correctamente.' };
}

/** Actualiza nombre/rol de un usuario. Solo administrador. */
function actualizarUsuario(token, datos) {
  var usuarioSesion = _validarSesion_(token);
  if (!usuarioSesion) return _respuestaSesionExpirada_();
  if (!_esAdmin_(usuarioSesion)) {
    return { exito: false, mensaje: 'No tiene permisos de administrador.' };
  }
  var usuario = String(datos.usuario || '');
  var u = _buscarUsuario_(usuario);
  if (!u) {
    return { exito: false, mensaje: 'No se encontró el usuario indicado.' };
  }
  var rol = String(datos.rol || u.Rol) === 'administrador' ? 'administrador' : 'editor';
  var activo = String(datos.activo || '').toUpperCase() === 'NO' ? 'NO' : 'SI';
  var esPrincipal = _esUsuarioPrincipal_(usuario);
  if (esPrincipal && String(usuarioSesion).toLowerCase() !== String(usuario).toLowerCase()) {
    return { exito: false, mensaje: 'El administrador principal solo puede modificar su propia información.' };
  }
  var hoja = obtenerHoja_(HOJA_USUARIOS);
  var colRol = ENCABEZADOS.Usuarios.indexOf('Rol') + 1;
  var colNombre = ENCABEZADOS.Usuarios.indexOf('Nombre') + 1;
  var colActivo = ENCABEZADOS.Usuarios.indexOf('Activo') + 1;
  if (String(usuarioSesion).toLowerCase() === String(usuario).toLowerCase()) {
    activo = 'SI';
    if (rol !== 'administrador') {
      return { exito: false, mensaje: 'No puede quitarse el rol de administrador a usted mismo.' };
    }
  }
  hoja.getRange(u._fila, colNombre).setValue(datos.nombre || u.Nombre || usuario);
  hoja.getRange(u._fila, colRol).setValue(rol);
  hoja.getRange(u._fila, colActivo).setValue(activo);
    Logger.log('Usuario actualizado: ' + usuario + ' -> ' + rol);
    _registrarActividad_(token, 'Usuarios', 'Editó usuario', usuario + ' -> ' + rol);
    return { exito: true, mensaje: 'Usuario "' + usuario + '" actualizado.' };
}

/** Cambia la contraseña de la sesión actual (se autentica con la actual). */
function cambiarContrasena(token, contrasenaActual, nueva) {
  var usuario = _validarSesion_(token);
  if (!usuario) return _respuestaSesionExpirada_();
  var u = _buscarUsuario_(usuario);
  if (!u) {
    return { exito: false, mensaje: 'No se encontró el usuario.' };
  }
  if (!_verificarContrasena_(u, contrasenaActual)) {
    return { exito: false, mensaje: 'La contraseña actual no es correcta.' };
  }
  if (!nueva || String(nueva).length < 8) {
    return { exito: false, mensaje: 'La nueva contraseña debe tener al menos 8 caracteres.' };
  }
  var hoja = obtenerHoja_(HOJA_USUARIOS);
  var colSalt = ENCABEZADOS.Usuarios.indexOf('Salt') + 1;
  var colHash = ENCABEZADOS.Usuarios.indexOf('Hash') + 1;
  var nuevoSalt = _generarSalt_();
  hoja.getRange(u._fila, colSalt).setValue(nuevoSalt);
  hoja.getRange(u._fila, colHash).setValue(_hashTexto_(nuevoSalt + nueva));
    Logger.log('Contraseña cambiada para: ' + usuario);
    _registrarActividad_(token, 'Usuarios', 'Cambió contraseña', usuario);
    return { exito: true, mensaje: 'Contraseña actualizada correctamente.' };
}

/** Genera una contraseña temporal aleatoria (8 caracteres). */
function _generarContrasenaTemporal_() {
  var chars = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnpqrstuvwxyz23456789';
  var out = '';
  for (var i = 0; i < 8; i++) {
    out += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return out;
}

/**
 * Recuperación de contraseña por correo: genera una contraseña temporal
 * y la envía por email a la cuenta. Solo si el envío tiene éxito se
 * guarda la nueva contraseña (así un fallo de correo nunca invalida la
 * contraseña real del usuario).
 * No revela si el correo existe (siempre responde éxito).
 * @param {string} correo
 */
function recuperarContrasena(correo) {
  try {
    var cor = String(correo || '').trim().toLowerCase();
    if (!esEmailValido_(cor)) {
      return { exito: false, mensaje: 'Ingrese un correo electrónico válido.' };
    }
    var u = _buscarUsuario_(cor);
    if (!u) {
      return { exito: true, mensaje: 'Si el correo está registrado, recibirá una contraseña temporal por email.' };
    }

    var destino = u.Email || u.ID_Usuario || cor;
    var temporal = _generarContrasenaTemporal_();

    // Enviar el correo ANTES de guardar la contraseña: si el envío falla,
    // la contraseña real queda intacta.
    MailApp.sendEmail(destino, 'Recuperación de contraseña',
      'Hola,\n\nSe ha generado una nueva contraseña temporal para tu cuenta:\n\n  ' + temporal +
      '\n\nEntra con ella y cámbiala desde el botón "Cambiar contraseña" de la barra superior.' +
      '\n\nSi no solicitaste este cambio, ignora este correo.');

    // El correo se envió correctamente: ahora sí se persiste la nueva contraseña.
    var hoja = obtenerHoja_(HOJA_USUARIOS);
    var colSalt = ENCABEZADOS.Usuarios.indexOf('Salt') + 1;
    var colHash = ENCABEZADOS.Usuarios.indexOf('Hash') + 1;
    var nuevoSalt = _generarSalt_();
    hoja.getRange(u._fila, colSalt).setValue(nuevoSalt);
    hoja.getRange(u._fila, colHash).setValue(_hashTexto_(nuevoSalt + temporal));

    Logger.log('Contraseña temporal enviada a: ' + destino);
    return { exito: true, mensaje: 'Si el correo está registrado, recibir\u00e1 una contrase\u00f1a temporal por email.' };
  } catch (err) {
    Logger.log('Error en recuperarContrasena: ' + err);
    return { exito: false, mensaje: 'Error al intentar recuperar la contraseña: ' + err.message };
  }
}

/** Desactiva o elimina un usuario. Solo administrador. */
function eliminarUsuario(token, usuario) {
  var usuarioSesion = _validarSesion_(token);
  if (!usuarioSesion) return _respuestaSesionExpirada_();
  if (!_esAdmin_(usuarioSesion)) {
    return { exito: false, mensaje: 'No tiene permisos de administrador.' };
  }
  if (String(usuarioSesion).toLowerCase() === String(usuario).toLowerCase()) {
    return { exito: false, mensaje: 'No puede eliminar su propio usuario.' };
  }
  var u = _buscarUsuario_(usuario);
  if (!u) {
    return { exito: false, mensaje: 'No se encontró el usuario indicado.' };
  }
  if (_esUsuarioPrincipal_(usuario)) {
    return { exito: false, mensaje: 'El administrador principal no puede eliminarse.' };
  }
  var hoja = obtenerHoja_(HOJA_USUARIOS);
  hoja.deleteRow(u._fila);
    Logger.log('Usuario eliminado: ' + usuario);
    _registrarActividad_(token, 'Usuarios', 'Eliminó usuario', usuario);
    return { exito: true, mensaje: 'Usuario "' + usuario + '" eliminado.' };
}

/**
 * ============================================================
 *  CITAS (con integración a Google Calendar)
 * ============================================================
 */

/**
 * Agenda una cita: crea un evento en Google Calendar y
 * registra la fila correspondiente en la hoja Citas.
 * @param {Object} datos {idCliente, titulo, fecha (yyyy-MM-dd),
 *                         hora (HH:mm), duracionMins, descripcion}
 */
function agendarCita(token, datos) {
  try {
    if (!_validarSesion_(token)) return _respuestaSesionExpirada_();

    // Servicios elegidos: precio y duración SIEMPRE se toman del catálogo.
    // La duración de la cita = suma de duraciones × cantidad (o la predeterminada).
    var selServicios = _serviciosDesdeCatalogo_(datos.servicios);
    var duracion;
    if (selServicios.items.length > 0) {
      duracion = selServicios.duracion;
    } else {
      duracion = parseInt(datos.duracionMins, 10) ||
        parseInt(obtenerConfiguracion().DURACION_CITA_PREDETERMINADA, 10) || 60;
    }
    if (duracion < 5) duracion = 60;

    // Evitar superposiciones con otras citas del mismo día.
    var choque = _citaTieneChoque_(datos.fecha, datos.hora, duracion);
    if (choque && choque.choca) {
      return { exito: false, mensaje: choque.mensaje };
    }

    var hoja = obtenerHoja_(HOJA_CITAS);
    var id = generarId('CITA');

    // Construir fecha/hora de inicio y fin.
    var partesFecha = String(datos.fecha).split('-');
    var partesHora = String(datos.hora).split(':');
    var inicio = new Date(
      parseInt(partesFecha[0], 10),
      parseInt(partesFecha[1], 10) - 1,
      parseInt(partesFecha[2], 10),
      parseInt(partesHora[0], 10),
      parseInt(partesHora[1], 10),
      0
    );
    var fin = new Date(inicio.getTime() + duracion * 60000);

    // Determinar quién agenda y el correo del cliente ANTES de crear el
    // evento, para poder añadirlos como invitados al calendario.
    var agendadoPor = '';
    var emailAgendador = '';
    var sesionUsuario = _validarSesion_(token);
    if (sesionUsuario) {
      var usuarioReg = _buscarUsuario_(sesionUsuario);
      agendadoPor = (usuarioReg && usuarioReg.Nombre) ? String(usuarioReg.Nombre) : sesionUsuario;
      emailAgendador = (usuarioReg && usuarioReg.Email) ? String(usuarioReg.Email) : '';
    }

    var correoCliente = String(datos.emailCliente || '');
    var nombreCliente = String(datos.nombreCliente || '');
    if (!correoCliente && datos.idCliente) {
      var clientesCita = filasAObjetos_(obtenerHoja_(HOJA_CLIENTES));
      for (var kc = 0; kc < clientesCita.length; kc++) {
        if (String(clientesCita[kc].ID_Cliente) === String(datos.idCliente)) {
          correoCliente = clientesCita[kc].Email || '';
          nombreCliente = String(clientesCita[kc].Nombre || '') + ' ' + String(clientesCita[kc].Apellido || '');
          break;
        }
      }
    }

    // Invitados del evento: quien agenda y el cliente. El dueño ya es el
    // organizador porque el evento se crea en su calendario por defecto.
    var invitados = [emailAgendador, correoCliente].filter(function(e) {
      return e && esEmailValido_(e);
    });

    // Crear el evento en el calendario predeterminado (del dueño).
    var idEvento = '';
    var errorCalendario = '';
    try {
      var calendario = CalendarApp.getDefaultCalendar();
      var opcionesEvento = { description: datos.descripcion || '' };
      if (invitados.length > 0) {
        opcionesEvento.guests = invitados.join(',');
      }
      var evento = calendario.createEvent(
        datos.titulo || 'Cita',
        inicio,
        fin,
        opcionesEvento
      );
      idEvento = evento.getId();
      Logger.log('Evento de calendario creado: ' + idEvento +
        (invitados.length ? ' | Invitados: ' + invitados.join(', ') : ''));
    } catch (errCal) {
      errorCalendario = errCal.message || String(errCal);
      Logger.log('Advertencia: no se pudo crear el evento de calendario: ' + errorCalendario);
    }

    hoja.appendRow([
      id,
      datos.idCliente || '',
      datos.titulo || '',
      datos.fecha || '',
      datos.hora || '',
      duracion,
      datos.descripcion || '',
      idEvento,
      'Programada',
      selServicios.items.length > 0 ? JSON.stringify(selServicios.items) : '',
      selServicios.items.length > 0 ? selServicios.total : '',
      agendadoPor
    ]);

    // Forzar Fecha y Hora como TEXTO para que Sheets no las convierta en fecha/hora.
    var filaCita = hoja.getLastRow();
    var colFechaC = ENCABEZADOS.Citas.indexOf('Fecha') + 1;
    var colHoraC = ENCABEZADOS.Citas.indexOf('Hora') + 1;
    hoja.getRange(filaCita, colFechaC).setNumberFormat('@')
      .setValue(String(datos.fecha || ''));
    hoja.getRange(filaCita, colHoraC).setNumberFormat('@')
      .setValue(String(datos.hora || ''));

    var avisoCorreo = '';
    try {
      // correoCliente y nombreCliente ya se obtuvieron antes de crear el evento.
      if (correoCliente) {
        var asunto = 'Cita agendada';
        var cuerpo = 'Estimado/a ' + (nombreCliente || 'cliente') +
          ':\n\nSe ha agendado su cita con los siguientes datos:\n' +
          'Título: ' + (datos.titulo || 'Cita') + '\n' +
          'Fecha: ' + (datos.fecha || '') + '\n' +
          'Hora: ' + _hora12_(datos.hora) + '\n' +
          'Duración: ' + (datos.duracionMins || 60) + ' min\n' +
          (datos.descripcion ? 'Descripción: ' + datos.descripcion + '\n' : '') +
          '\nGracias por preferirnos.';
        MailApp.sendEmail(correoCliente, asunto, cuerpo);
        avisoCorreo = ' Confirmación enviada a ' + correoCliente + '.';
      }
    } catch (errCorreo) {
      avisoCorreo = ' No se pudo enviar la confirmación por correo: ' + (errCorreo.message || String(errCorreo));
      Logger.log('Error al enviar correo de cita: ' + errCorreo);
    }

    // Aviso interno de la cita: SOLO quien la agendó y el dueño de la hoja.
    // Nunca se notifica a todos los usuarios.
    try {
      var destinatariosAviso = _correosAvisoCita_(emailAgendador);
      if (destinatariosAviso.length > 0) {
        var cuerpoAviso =
          'Se agendó una nueva cita.\n\n' +
          'Título: ' + (datos.titulo || 'Sin título') + '\n' +
          'Cliente: ' + (datos.nombreCliente || (datos.idCliente || '')) + '\n' +
          'Fecha: ' + (datos.fecha || '') + '\n' +
          'Hora: ' + _hora12_(datos.hora) + '\n' +
          'Duración: ' + duracion + ' min\n' +
          'Agendada por: ' + agendadoPor + '\n' +
          (datos.descripcion ? 'Descripción: ' + datos.descripcion + '\n' : '') +
          'ID de cita: ' + id + '\n';
        MailApp.sendEmail(destinatariosAviso.join(','), 'Nueva cita agendada — ' + (datos.fecha || ''), cuerpoAviso);
      }
    } catch (errAviso) {
      Logger.log('Error al enviar aviso de cita: ' + errAviso);
    }

    Logger.log('Cita agendada: ' + id);
    _registrarActividad_(token, 'Citas', 'Agendó cita', (datos.titulo || 'Cita') + ' ' + (datos.fecha || '') + ' ' + (datos.hora || ''));
    return {
      exito: true,
      mensaje: (idEvento
        ? 'Cita agendada y evento creado en Google Calendar.'
        : 'Cita agendada, pero NO se pudo crear el evento en el calendario: ' + errorCalendario) + avisoCorreo,
      id: id
    };
  } catch (err) {
    Logger.log('Error al agendar cita: ' + err);
    return { exito: false, mensaje: 'Error al agendar cita: ' + err.message };
  }
}

/**
 * Lee y devuelve todas las citas como objetos con UNA sola lectura de la
 * hoja. De paso, si alguna celda Fecha/Hora quedó como tipo Fecha
 * (p. ej. hora con base 1899-12-30), la repara en memoria y la escribe
 * como texto limpio ("2026-08-05" / "14:44") solo si hubo cambios.
 * @return {Array} citas (objetos).
 */
function _leerCitas_() {
  var hoja = obtenerHoja_(HOJA_CITAS);
  var valores = hoja.getDataRange().getValues();
  if (valores.length < 2) return [];
  var colFecha = ENCABEZADOS.Citas.indexOf('Fecha') + 1;
  var colHora = ENCABEZADOS.Citas.indexOf('Hora') + 1;
  var zona = obtenerZonaHoraria_();
  var cambios = false;
  for (var i = 1; i < valores.length; i++) {
    if (valores[i][colFecha - 1] instanceof Date) {
      valores[i][colFecha - 1] = Utilities.formatDate(valores[i][colFecha - 1], zona, 'yyyy-MM-dd');
      cambios = true;
    }
    if (valores[i][colHora - 1] instanceof Date) {
      valores[i][colHora - 1] = Utilities.formatDate(valores[i][colHora - 1], zona, 'HH:mm');
      cambios = true;
    }
  }
  if (cambios) {
    var filasDatos = valores.length - 1;
    hoja.getRange(2, colFecha, filasDatos, 1).setNumberFormat('@');
    hoja.getRange(2, colHora, filasDatos, 1).setNumberFormat('@');
    hoja.getRange(2, colFecha, filasDatos, 1).setValues(
      valores.slice(1).map(function(r) { return [r[colFecha - 1]]; }));
    hoja.getRange(2, colHora, filasDatos, 1).setValues(
      valores.slice(1).map(function(r) { return [r[colHora - 1]]; }));
    Logger.log('Citas: fechas y horas corregidas a texto.');
  }
  return _filasAObjetosDesdeValores_(valores);
}

/**
 * Devuelve todas las citas.
 */
function obtenerCitas(token) {
  if (!_validarSesion_(token)) return _respuestaSesionExpirada_();
  return _leerCitas_();
}

/**
 * Firma de cambios: valores baratos (última fila + fecha de modificación)
 * de las hojas de datos. El panel la consulta cada pocos segundos para
 * detectar cuándo alguien agregó un cliente o registró una cita (p. ej. por
 * la página pública de reservas) y refrescar la vista abierta solo entonces.
 * No transfiere datos de las filas, solo contadores/fechas.
 * @param {string} token  Token de sesión.
 * @return {Object} { exito, clientes, citas, historial }
 */
function obtenerFirmaCambios(token) {
  if (!_validarSesion_(token)) return _respuestaSesionExpirada_();
  // Se abre la hoja de cálculo UNA sola vez (esta función es la más invocada
  // de la app: el panel la consulta cada 30 segundos).
  var hojaCalculo;
  try {
    hojaCalculo = obtenerHojaCalculo_();
  } catch (err) {
    return { exito: true, clientes: '', citas: '', historial: '' };
  }
  function firmaHoja(nombre) {
    try {
      var hoja = hojaCalculo.getSheetByName(nombre);
      return hoja ? (hoja.getLastRow() + '|' + hoja.getLastModified().getTime()) : '';
    } catch (err) {
      return '';
    }
  }
  return {
    exito: true,
    clientes: firmaHoja(HOJA_CLIENTES),
    citas: firmaHoja(HOJA_CITAS),
    historial: firmaHoja(HOJA_HISTORIAL)
  };
}

/**
 * Cambia el estado de una cita (Programada / Completada / Cancelada).
 * @param {string} id     ID de la cita.
 * @param {string} estado Nuevo estado.
 */
function cambiarEstadoCita(token, id, estado) {
  try {
    if (!_validarSesion_(token)) return _respuestaSesionExpirada_();
    var estadosValidos = ['Programada', 'Completada', 'Cancelada', 'Expirada'];
    estado = String(estado || '');
    if (estadosValidos.indexOf(estado) === -1) {
      return { exito: false, mensaje: 'Estado no válido.' };
    }
    var hoja = obtenerHoja_(HOJA_CITAS);
    var fila = buscarFilaPorId_(hoja, 'ID_Cita', id);
    if (!fila) {
      return { exito: false, mensaje: 'No se encontró la cita indicada.' };
    }
    var colEstado = ENCABEZADOS.Citas.indexOf('Estado') + 1;
    var estadoActual = String(hoja.getRange(fila, colEstado).getValue() || '');
    if (estadoActual === 'Completada' && estado === 'Programada') {
      return { exito: false, mensaje: 'Una cita completada no se puede desmarcar.' };
    }
    hoja.getRange(fila, colEstado).setValue(estado);

    // Al completarse, se genera (o actualiza) el historial con los datos de la
    // cita. Al volver a "Programada", se elimina ese registro.
    if (estado === 'Completada') {
      var valoresCita = hoja.getRange(fila, 1, 1, ENCABEZADOS.Citas.length).getValues()[0];
      _crearHistorialDesdeCita_(id, valoresCita);
    } else if (estado === 'Programada') {
      _eliminarHistorialDeCita_(id);
    }

    Logger.log('Estado de cita actualizado: ' + id + ' -> ' + estado);
    _registrarActividad_(token, 'Citas', 'Cambió estado de cita', 'ID ' + id + ' -> ' + estado);
    return { exito: true, mensaje: 'Estado actualizado a "' + estado + '".' };
  } catch (err) {
    Logger.log('Error al cambiar estado de cita: ' + err);
    return { exito: false, mensaje: 'Error al cambiar el estado de la cita: ' + err.message };
  }
}

/**
 * Actualiza una cita existente y, si tiene evento en Google Calendar,
 * también actualiza la hora del evento.
 * @param {string} id    ID de la cita.
 * @param {Object} datos {idCliente, titulo, fecha, hora, duracionMins, descripcion}
 */
// Construye la descripción del historial: título + nota del empleado + servicios,
// sin la línea de TOTAL (el monto vive en su propia columna) y sin duplicar el
// desglose que ya viene embebido en la descripción de la cita (reservas en línea).
function _descripcionHistorial_(valoresCita) {
  var titulo = valoresCita[2] || '';
  var descripcionCita = valoresCita[6] || '';

  // Quitar el bloque automático (Servicios:/TOTAL:) que ya trae la cita; el
  // desglose de servicios ahora vive en su propia columna del historial.
  var nota = String(descripcionCita || '')
    .replace(/\s*Servicios:[\s\S]*$/i, '')
    .replace(/\s*TOTAL:[\s\S]*$/i, '')
    .trim();

  return (titulo ? titulo : '') + (nota ? ' — ' + nota : '');
}

// Genera o actualiza el registro de historial de una cita con sus datos.
// valoresCita: [ID_Cita, ID_Cliente, Titulo, Fecha, Hora, Duracion_Mins, Descripcion, ID_Evento_Calendar, Estado, Servicios, Total_Precio, Agendado_Por]
function _crearHistorialDesdeCita_(idCita, valoresCita) {
  var hojaHistorial = obtenerHoja_(HOJA_HISTORIAL);
  var cab = ENCABEZADOS[HOJA_HISTORIAL];
  var colIDCita = cab.indexOf('ID_Cita') + 1;
  var colFecha = cab.indexOf('Fecha') + 1;
  var fecha = Utilities.formatDate(new Date(), obtenerZonaHoraria_(), 'yyyy-MM-dd HH:mm');

  var filaExistente = 0;
  var ultimaFila = hojaHistorial.getLastRow();
  if (ultimaFila >= 2) {
    // Leer solo la columna ID_Cita (evita traer todo el historial para buscarla).
    var ids = hojaHistorial.getRange(2, colIDCita, ultimaFila - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(idCita)) { filaExistente = i + 2; break; }
    }
  }

  var idCliente = valoresCita[1] || '';
  var descripcion = _descripcionHistorial_(valoresCita);

  if (filaExistente) {
    hojaHistorial.getRange(filaExistente, colFecha, 1, 3).setValues([[
      fecha, descripcion, 'Completada'
    ]]);
    return;
  }
  hojaHistorial.appendRow([
    generarId('HIS'),
    idCliente,
    idCita,
    fecha,
    descripcion,
    'Completada'
  ]);
}

// Elimina el registro de historial generado para una cita (si existe).
function _eliminarHistorialDeCita_(idCita) {
  var hojaHistorial = obtenerHoja_(HOJA_HISTORIAL);
  var cab = ENCABEZADOS[HOJA_HISTORIAL];
  var colIDCita = cab.indexOf('ID_Cita') + 1;
  var ultimaFila = hojaHistorial.getLastRow();
  if (ultimaFila >= 2) {
    var ids = hojaHistorial.getRange(2, colIDCita, ultimaFila - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(idCita)) {
        hojaHistorial.deleteRow(i + 2);
        return;
      }
    }
  }
}

function actualizarCita(token, id, datos) {
  try {
    if (!_validarSesion_(token)) return _respuestaSesionExpirada_();

    // Evitar superposiciones con otras citas del mismo día (ignorando la propia).
    var choque = _citaTieneChoque_(datos.fecha, datos.hora, datos.duracionMins, id);
    if (choque && choque.choca) {
      return { exito: false, mensaje: choque.mensaje };
    }

    var hoja = obtenerHoja_(HOJA_CITAS);
    var fila = buscarFilaPorId_(hoja, 'ID_Cita', id);
    if (!fila) {
      return { exito: false, mensaje: 'No se encontró la cita indicada.' };
    }
    var selServiciosUpd = _serviciosDesdeCatalogo_(datos.servicios);
    var duracion;
    if (selServiciosUpd.items.length > 0) {
      duracion = selServiciosUpd.duracion;
    } else {
      duracion = parseInt(datos.duracionMins, 10) || 60;
    }
    if (duracion < 5) duracion = 60;
    var colCli = ENCABEZADOS.Citas.indexOf('ID_Cliente') + 1;
    var colFecha = ENCABEZADOS.Citas.indexOf('Fecha') + 1;
    var colHora = ENCABEZADOS.Citas.indexOf('Hora') + 1;
    hoja.getRange(fila, colCli).setValue(datos.idCliente || '');
    hoja.getRange(fila, 3, 1, 5).setValues([[
      datos.titulo || '',
      String(datos.fecha || ''),
      String(datos.hora || ''),
      duracion,
      datos.descripcion || ''
    ]]);
    var colServicios = ENCABEZADOS.Citas.indexOf('Servicios') + 1;
    var colTotal = ENCABEZADOS.Citas.indexOf('Total_Precio') + 1;
    hoja.getRange(fila, colServicios).setValue(selServiciosUpd.items.length > 0 ? JSON.stringify(selServiciosUpd.items) : '');
    hoja.getRange(fila, colTotal).setValue(selServiciosUpd.items.length > 0 ? selServiciosUpd.total : '');
    // Mantener Fecha y Hora como texto.
    hoja.getRange(fila, colFecha).setNumberFormat('@').setValue(String(datos.fecha || ''));
    hoja.getRange(fila, colHora).setNumberFormat('@').setValue(String(datos.hora || ''));

    // Actualizar el evento de Calendar si existe.
    var colEvento = ENCABEZADOS.Citas.indexOf('ID_Evento_Calendar') + 1;
    var idEvento = hoja.getRange(fila, colEvento).getValue();
    if (idEvento) {
      try {
        var partesFecha = String(datos.fecha).split('-');
        var partesHora = String(datos.hora).split(':');
        var inicio = new Date(
          parseInt(partesFecha[0], 10),
          parseInt(partesFecha[1], 10) - 1,
          parseInt(partesFecha[2], 10),
          parseInt(partesHora[0], 10),
          parseInt(partesHora[1], 10),
          0
        );
        var fin = new Date(inicio.getTime() + duracion * 60000);
        var evento = CalendarApp.getDefaultCalendar().getEventById(String(idEvento));
        if (evento) {
          evento.setTitle(datos.titulo || 'Cita');
          evento.setTime(inicio, fin);
          if (datos.descripcion) evento.setDescription(datos.descripcion);
          Logger.log('Evento de calendario actualizado: ' + idEvento);
        }
      } catch (errCal) {
        Logger.log('Advertencia: no se pudo actualizar el evento de calendario: ' + errCal);
      }
    }

    Logger.log('Cita actualizada: ' + id);
    _registrarActividad_(token, 'Citas', 'Editó cita', 'ID ' + id);
    return { exito: true, mensaje: 'Cita actualizada correctamente.' };
  } catch (err) {
    Logger.log('Error al actualizar cita: ' + err);
    return { exito: false, mensaje: 'Error al actualizar la cita: ' + err.message };
  }
}

/**
 * Elimina una cita por su ID y, si existe, también el evento
 * correspondiente en Google Calendar.
 * @param {string} id  ID de la cita.
 */
function eliminarCita(token, id) {
  try {
    var usuario = _validarSesion_(token);
    if (!usuario) return _respuestaSesionExpirada_();
    if (!_esAdmin_(usuario)) return { exito: false, mensaje: 'Solo el administrador puede eliminar citas.' };
    var hoja = obtenerHoja_(HOJA_CITAS);
    var fila = buscarFilaPorId_(hoja, 'ID_Cita', id);
    if (!fila) {
      return { exito: false, mensaje: 'No se encontró la cita indicada.' };
    }
    var colEvento = ENCABEZADOS.Citas.indexOf('ID_Evento_Calendar') + 1;
    var idEvento = hoja.getRange(fila, colEvento).getValue();
    if (idEvento) {
      try {
        var evento = CalendarApp.getDefaultCalendar().getEventById(String(idEvento));
        if (evento) evento.deleteEvent();
        Logger.log('Evento de calendario eliminado: ' + idEvento);
      } catch (errCal) {
        Logger.log('Advertencia: no se pudo eliminar el evento de calendario: ' + errCal);
      }
    }
    hoja.deleteRow(fila);
    Logger.log('Cita eliminada: ' + id);
    _registrarActividad_(token, 'Citas', 'Eliminó cita', 'ID ' + id);
    return { exito: true, mensaje: 'Cita eliminada correctamente.' };
  } catch (err) {
    Logger.log('Error al eliminar cita: ' + err);
    return { exito: false, mensaje: 'Error al eliminar la cita: ' + err.message };
  }
}

/**
 * ============================================================
 *  RESERVAS EN LÍNEA (página pública para los clientes)
 *  Funciones SIN token: las usa cualquier persona, sin iniciar
 *  sesión. Solo devuelven disponibilidad y configuración; nunca
 *  exponen la lista de clientes ni datos sensibles.
 * ============================================================
 */

/**
 * Devuelve el estado público para la página de reservas:
 * nombre del negocio, etiqueta de cita, color, duración por
 * defecto y si las reservas en línea están habilitadas.
 * @return {Object}
 */
function obtenerEstadoPublico() {
  try {
    var cfg = obtenerConfiguracion();
    return {
      exito: true,
      configurada:                String(cfg.CONFIGURADA || 'NO').toUpperCase() === 'SI',
      nombreNegocio:              cfg.NOMBRE_NEGOCIO || CONFIGURACION_PREDETERMINADA.NOMBRE_NEGOCIO,
      etiquetaCita:               cfg.ETIQUETA_CITA || CONFIGURACION_PREDETERMINADA.ETIQUETA_CITA,
      colorPrimario:              cfg.COLOR_PRIMARIO || CONFIGURACION_PREDETERMINADA.COLOR_PRIMARIO,
      duracionCitaPredeterminada: parseInt(cfg.DURACION_CITA_PREDETERMINADA, 10) || 60,
      habilitarReservas:          String(cfg.HABILITAR_RESERVAS || 'NO').toUpperCase() === 'SI',
      pasoReserva:                parseInt(cfg.PASO_RESERVA_MIN, 10) || 30,
      horarioAtencion:            _parseHorarioAtencion_(cfg.HORARIO_ATENCION)
    };
  } catch (err) {
    Logger.log('Error en obtenerEstadoPublico: ' + err);
    return { exito: false, mensaje: 'Error al obtener la configuración: ' + err.message };
  }
}

/**
 * Interpreta el valor de HORARIO_ATENCION (JSON) como un objeto
 * { getDay: {abre, cierra, activo} }. getDay: 0=Domingo...6=Sábado.
 * @param {*} valor  JSON en texto o sin parsear.
 * @return {Object}
 */
function _parseHorarioAtencion_(valor) {
  var horario = {};
  for (var d = 0; d < 7; d++) {
    horario[d] = { abre: '', cierra: '', activo: false };
  }
  try {
    var obj = JSON.parse(String(valor || '{}'));
    for (var d = 0; d < 7; d++) {
      var item = obj[d] || {};
      var abre = String(item.abre || '').trim();
      var cierra = String(item.cierra || '').trim();
      if (abre && cierra) {
        horario[d] = { abre: abre, cierra: cierra, activo: true };
      }
    }
  } catch (err) {
    Logger.log('HORARIO_ATENCION no válido: ' + err);
  }
  return horario;
}

/** Convierte "HH:mm" a minutos desde las 00:00; null si no es válido. */
function _horaAMin_(h) {
  var m = String(h || '').match(/^(\d{1,2}):(\d{2})$/);
  if (!m) return null;
  var hh = parseInt(m[1], 10);
  var mm = parseInt(m[2], 10);
  if (hh < 0 || hh > 23 || mm < 0 || mm > 59) return null;
  return hh * 60 + mm;
}

/** Convierte minutos en "HH:mm" (24 h). */
function _minAHora_(min) {
  var hh = Math.floor(min / 60);
  var mm = min % 60;
  return (hh < 10 ? '0' : '') + hh + ':' + (mm < 10 ? '0' : '') + mm;
}

/** Convierte "HH:mm" a formato de 12 horas (h:mm AM/PM). */
function _hora12_(h) {
  var m = String(h || '').match(/^(\d{1,2}):(\d{2})/);
  if (!m) return String(h || '');
  var hh = parseInt(m[1], 10);
  var suf = hh >= 12 ? 'PM' : 'AM';
  var h12 = hh % 12;
  if (h12 === 0) h12 = 12;
  return h12 + ':' + m[2] + ' ' + suf;
}

/** Formatea un número como dinero ("$ 1,234.50"). */
function _formatoMoneda_(n) {
  n = parseFloat(n) || 0;
  return '$ ' + n.toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

/**
 * Genera los horarios candidatos del día: desde "abre" hasta que ya no
 * quepan "duracion" minutos dentro de "cierra", cada "paso" minutos.
 * @return {Array} [{ hora, inicio, fin }]
 */
function _generarSlots_(abre, cierra, duracion, paso) {
  var inicioTope = _horaAMin_(abre);
  var finTope = _horaAMin_(cierra);
  if (inicioTope === null || finTope === null) return [];
  var slots = [];
  for (var t = inicioTope; (t + duracion) <= finTope; t += paso) {
    slots.push({ hora: _minAHora_(t), inicio: t, fin: t + duracion });
  }
  return slots;
}

/**
 * Intervalos ocupados (en minutos de día) de la hoja Citas para una fecha:
 * las citas con estado distinto de Completada/Cancelada bloquean el hueco
 * [inicio, inicio+duración).
 * @param {string} fecha     Fecha "yyyy-MM-dd".
 * @param {string} exceptoId ID_Cita a ignorar (p. ej. la cita en edición).
 * @return {Array} [{ inicio, fin }]
 */
function _intervalosOcupados_(fecha, exceptoId) {
  var hoja = obtenerHoja_(HOJA_CITAS);
  var valores = hoja.getDataRange().getValues();
  if (valores.length < 2) return [];
  var cab = valores[0];
  var colId = cab.indexOf('ID_Cita');
  var colFecha = cab.indexOf('Fecha');
  var colHora = cab.indexOf('Hora');
  var colDur = cab.indexOf('Duracion_Mins');
  var colEstado = cab.indexOf('Estado');
  if (colFecha < 0 || colHora < 0 || colEstado < 0) return [];

  var ocupados = [];
  for (var i = 1; i < valores.length; i++) {
    var fila = valores[i];
    if (exceptoId && colId >= 0 && String(fila[colId] || '') === String(exceptoId)) {
      continue; // Ignorar la propia cita (edición).
    }
    var fechaCita = String(fila[colFecha] || '');
    // Fechas que se leyeron como Date se normalizan a texto.
    if (fechaCita instanceof Date) {
      fechaCita = Utilities.formatDate(fechaCita, obtenerZonaHoraria_(), 'yyyy-MM-dd');
    }
    var mFecha = String(fechaCita).match(/^(\d{4}-\d{2}-\d{2})/);
    if (!mFecha || mFecha[1] !== fecha) continue;
    var estado = String(fila[colEstado] || '').trim();
    if (estado === 'Completada' || estado === 'Cancelada') continue;
    var inicioMin = _horaAMin_(String(fila[colHora] || ''));
    if (inicioMin === null) continue;
    var duracion = parseInt(fila[colDur], 10) || 60;
    ocupados.push({ inicio: inicioMin, fin: inicioMin + duracion });
  }
  return ocupados;
}

/**
 * Determina si una etiqueta de cita es femenina (regla heurística para
 * concordar el artículo). Femenina si termina en "a", "ción", "sión",
 * "dad", "tad" o "tud"; en otro caso se asume masculina.
 * @param {string} etiqueta  Ej. "cita", "mantenimiento", "revisión".
 * @return {boolean}
 */
function _etiquetaFemenina_(etiqueta) {
  var e = String(etiqueta || '').trim().toLowerCase();
  if (!e) return true;
  return /(a|ción|sión|dad|tad|tud)$/.test(e);
}

/** Devuelve el artículo indeterminado ("una"/"un") para una etiqueta. */
function _articulo_(etiqueta) {
  return _etiquetaFemenina_(etiqueta) ? 'una' : 'un';
}

/**
 * Comprueba si una cita [hora, hora+duracion) de una fecha choca con otra
 * cita existente (estado distinto de Completada/Cancelada).
 * @param {string} fecha     Fecha "yyyy-MM-dd".
 * @param {string} hora      Hora "HH:mm".
 * @param {number|string} duracion Duración en minutos.
 * @param {string} [exceptoId] ID_Cita a ignorar (la propia al editar).
 * @return {Object|null} { choca, mensaje } o null si la hora es inválida.
 */
function _citaTieneChoque_(fecha, hora, duracion, exceptoId) {
  var inicioMin = _horaAMin_(String(hora || ''));
  if (inicioMin === null) return null;
  var dur = parseInt(duracion, 10) || 60;
  var finMin = inicioMin + dur;
  var ocupados = _intervalosOcupados_(String(fecha || ''), exceptoId);
  var etiqueta = obtenerConfiguracion().ETIQUETA_CITA || CONFIGURACION_PREDETERMINADA.ETIQUETA_CITA;
  for (var i = 0; i < ocupados.length; i++) {
    var o = ocupados[i];
    if (inicioMin < o.fin && finMin > o.inicio) {
      return {
        choca: true,
        mensaje: 'Ya existe ' + _articulo_(etiqueta) + ' ' + etiqueta +
          ' en ese horario (' + _minAHora_(o.inicio) + ' - ' + _minAHora_(o.fin) + '). Elija otra hora.'
      };
    }
  }
  return { choca: false, mensaje: '' };
}

/**
 * Calcula los horarios disponibles para una fecha y duración dadas.
 * No requiere sesión. No incluye datos sensibles.
 * @param {Object} datos  { fecha: "yyyy-MM-dd", duracionMins }
 * @return {Object} { exito, fecha, duracion, dia, mensaje?, slots: ["HH:mm", ...] }
 */
function obtenerHorariosDisponibles(datos) {
  try {
    var cfg = obtenerConfiguracion();
    if (String(cfg.HABILITAR_RESERVAS || 'NO').toUpperCase() !== 'SI') {
      return { exito: false, mensaje: 'La reserva en línea está deshabilitada en este momento.' };
    }
    datos = datos || {};
    var fecha = String(datos.fecha || '');
    if (!/^\d{4}-\d{2}-\d{2}$/.test(fecha)) {
      return { exito: false, mensaje: 'La fecha no es válida.' };
    }
    var duracion = parseInt(datos.duracionMins, 10) ||
      parseInt(cfg.DURACION_CITA_PREDETERMINADA, 10) || 60;
    if (duracion < 5) duracion = 60;

    var partes = fecha.split('-');
    var js = new Date(parseInt(partes[0], 10), parseInt(partes[1], 10) - 1, parseInt(partes[2], 10));
    var dia = js.getDay();
    var hDia = _parseHorarioAtencion_(cfg.HORARIO_ATENCION)[dia];
    if (!hDia || !hDia.activo) {
      return { exito: true, fecha: fecha, duracion: duracion, dia: dia, slots: [],
        mensaje: 'Este día no hay atención. Elija otro día.' };
    }

    var paso = parseInt(cfg.PASO_RESERVA_MIN, 10) || 30;
    var slots = _generarSlots_(hDia.abre, hDia.cierra, duracion, paso);
    var ocupados = _intervalosOcupados_(fecha);

    var hoy = Utilities.formatDate(new Date(), obtenerZonaHoraria_(), 'yyyy-MM-dd');
    var ahoraMin = null;
    if (fecha === hoy) {
      var fh = Utilities.formatDate(new Date(), obtenerZonaHoraria_(), 'HH:mm');
      ahoraMin = _horaAMin_(fh);
    }

    var libres = [];
    for (var i = 0; i < slots.length; i++) {
      var s = slots[i];
      // Ocultar los horarios que ya empezaron o pasaron si el día es hoy
      // (la cita debe iniciar en el futuro, no en un momento ya transcurrido).
      if (ahoraMin !== null && s.inicio <= ahoraMin) continue;
      var choca = false;
      for (var j = 0; j < ocupados.length; j++) {
        if (s.inicio < ocupados[j].fin && s.fin > ocupados[j].inicio) { choca = true; break; }
      }
      if (!choca) libres.push(s.hora);
    }

    return { exito: true, fecha: fecha, duracion: duracion, dia: dia, slots: libres };
  } catch (err) {
    Logger.log('Error en obtenerHorariosDisponibles: ' + err);
    return { exito: false, mensaje: 'Error al calcular los horarios: ' + err.message };
  }
}

/**
 * Busca un cliente existente por email (o teléfono); si no existe,
 * crea su ficha y devuelve su ID. Sin token (uso público).
 * @return {string} ID_Cliente
 */
function _buscarOCrearClientePublico_(nombre, apellido, email, telefono) {
  var hoja = obtenerHoja_(HOJA_CLIENTES);
  var datos = hoja.getDataRange().getValues();
  var cab = datos[0];
  var colId = cab.indexOf('ID_Cliente');
  var colEmail = cab.indexOf('Email');
  var colTel = cab.indexOf('Telefono');

  for (var i = 1; i < datos.length; i++) {
    if (colEmail >= 0 && email &&
        String(datos[i][colEmail] || '').trim().toLowerCase() === email) {
      return String(datos[i][colId] || '');
    }
    if (colTel >= 0 && telefono &&
        String(datos[i][colTel] || '').trim() === telefono) {
      return String(datos[i][colId] || '');
    }
  }

  var id = generarId('CLI');
  var fechaRegistro = Utilities.formatDate(new Date(), obtenerZonaHoraria_(), 'yyyy-MM-dd HH:mm');
  hoja.appendRow([id, nombre, apellido, telefono, email, '', '', fechaRegistro, '']);
  Logger.log('Cliente creado por reserva pública: ' + id);
  return id;
}

/**
 * Busca en Clientes por email o teléfono (uso público, sin token).
 * Solo devuelve datos básicos (nombre/apellido), no datos sensibles,
 * para que un cliente que ya está registrado no deba volver a dar sus datos.
 * @param {Object} datos { email?, telefono? }
 * @return {Object} { exito, encontrado, nombre, apellido }
 */
function buscarClientePublico(datos) {
  try {
    datos = datos || {};
    var email = String(datos.email || '').trim().toLowerCase();
    var telefono = String(datos.telefono || '').trim();
    if (!email && !telefono) {
      return { exito: false, mensaje: 'Indique su correo o su teléfono.' };
    }
    var hoja = obtenerHoja_(HOJA_CLIENTES);
    var valores = hoja.getDataRange().getValues();
    var cab = valores[0];
    var colNombre = cab.indexOf('Nombre');
    var colApellido = cab.indexOf('Apellido');
    var colEmail = cab.indexOf('Email');
    var colTel = cab.indexOf('Telefono');
    for (var i = 1; i < valores.length; i++) {
      var fila = valores[i];
      var coincide = false;
      if (colEmail >= 0 && email &&
          String(fila[colEmail] || '').trim().toLowerCase() === email) coincide = true;
      if (!coincide && colTel >= 0 && telefono &&
          String(fila[colTel] || '').trim() === telefono) coincide = true;
      if (coincide) {
        return {
          exito: true,
          encontrado: true,
          nombre: colNombre >= 0 ? String(fila[colNombre] || '').trim() : '',
          apellido: colApellido >= 0 ? String(fila[colApellido] || '').trim() : ''
        };
      }
    }
    return { exito: true, encontrado: false, nombre: '', apellido: '' };
  } catch (err) {
    Logger.log('Error en buscarClientePublico: ' + err);
    return { exito: false, mensaje: 'Error al buscar el cliente: ' + err.message };
  }
}

/**
 * Valida la selección de servicios del cliente contra el catálogo activo.
 * Usa SIEMPRE el precio y la duración guardados en la hoja (el cliente no
 * puede alterarlos). @return {Object|null} { items, total, duracion }
 */
function _serviciosDesdeCatalogo_(solicitados) {
  var lista = [];
  var total = 0;
  var duracion = 0;
  var solicitados = (solicitados && solicitados.length) ? solicitados : [];
  if (solicitados.length === 0) return { items: [], total: 0, duracion: 0 };

  var hoja = obtenerHoja_(HOJA_SERVICIOS);
  var valores = hoja.getDataRange().getValues();
  var cab = valores[0];
  var colId = cab.indexOf('ID_Servicio');
  var colNombre = cab.indexOf('Nombre');
  var colPrecio = cab.indexOf('Precio');
  var colDur = cab.indexOf('Duracion_Mins');
  var colActivo = cab.indexOf('Activo');
  if (colId < 0 || colNombre < 0) return { items: [], total: 0, duracion: 0 };

  var catalogo = {};
  for (var i = 1; i < valores.length; i++) {
    var fila = valores[i];
    var idS = String(fila[colId] || '');
    if (!idS) continue;
    if (colActivo >= 0 && String(fila[colActivo] || 'SI').toUpperCase() === 'NO') continue;
    catalogo[idS] = {
      nombre: String(fila[colNombre] || '').trim(),
      precio: parseFloat(fila[colPrecio]) || 0,
      duracionMins: parseInt(fila[colDur], 10) || 60
    };
  }

  for (var j = 0; j < solicitados.length; j++) {
    var sol = solicitados[j] || {};
    var s = catalogo[String(sol.id || '')];
    if (!s) continue; // Servicio inexistente o inactivo: se ignora.
    var cantidad = parseInt(sol.cantidad, 10);
    if (isNaN(cantidad) || cantidad < 1) cantidad = 1;
    var subtotal = s.precio * cantidad;
    total += subtotal;
    duracion += s.duracionMins * cantidad;
    lista.push({
      id: String(sol.id),
      nombre: s.nombre,
      precio: s.precio,
      cantidad: cantidad,
      subtotal: subtotal,
      duracionMins: s.duracionMins
    });
  }
  return { items: lista, total: total, duracion: duracion };
}

/**
 * Reserva pública de una cita: crea (o vincula) el cliente, agrega la cita
 * en la hoja, crea el evento de Calendar y envía la confirmación por correo.
 * Sin token. Usa LockService y vuelve a validar el slot en el servidor para
 * evitar que dos personas reserven la misma hora a la vez.
 * @param {Object} datos { fecha, hora, duracionMins, servicios,
 *                         cliente: { nombre, apellido, email, telefono } }
 * @return {Object}
 */
function reservarCitaPublica(datos) {
  var candado = LockService.getScriptLock();
  try {
    candado.waitLock(30000);
  } catch (errLock) {
    return { exito: false, mensaje: 'Hay otra reserva en curso. Intente de nuevo en unos segundos.' };
  }

  try {
    var cfg = obtenerConfiguracion();
    if (String(cfg.HABILITAR_RESERVAS || 'NO').toUpperCase() !== 'SI') {
      return { exito: false, mensaje: 'La reserva en línea está deshabilitada en este momento.' };
    }
    datos = datos || {};
    var fecha = String(datos.fecha || '');
    var hora = String(datos.hora || '');

    // Servicios elegidos por el cliente: precio y duración SIEMPRE se toman
    // del catálogo. La duración de la cita = suma de duraciones × cantidad.
    var selServicios = _serviciosDesdeCatalogo_(datos.servicios);
    var duracion;
    if (selServicios.items.length > 0) {
      duracion = selServicios.duracion;
    } else {
      duracion = parseInt(datos.duracionMins, 10) ||
        parseInt(cfg.DURACION_CITA_PREDETERMINADA, 10) || 60;
    }
    if (duracion < 5) duracion = 60;
    if (!/^\d{4}-\d{2}-\d{2}$/.test(fecha) || !_horaAMin_(hora)) {
      return { exito: false, mensaje: 'Los datos de la fecha u hora no son válidos.' };
    }

    // Datos del cliente.
    var cli = datos.cliente || {};
    var nombre = String(cli.nombre || '').trim();
    var apellido = String(cli.apellido || '').trim();
    var email = String(cli.email || '').trim().toLowerCase();
    var telefono = String(cli.telefono || '').trim();
    if (!nombre) {
      return { exito: false, mensaje: 'Ingrese su nombre.' };
    }
    if (email && !esEmailValido_(email)) {
      return { exito: false, mensaje: 'El correo electrónico no tiene un formato válido.' };
    }
    if (telefono && !esTelefonoValido_(telefono)) {
      return { exito: false, mensaje: 'El teléfono no tiene un formato válido.' };
    }
    if (!email && !telefono) {
      return { exito: false, mensaje: 'Ingrese su correo electrónico o su teléfono.' };
    }

    // Volver a validar el slot justo antes de escribir (evita dobles reservas).
    var dispon = obtenerHorariosDisponibles({ fecha: fecha, duracionMins: duracion });
    if (!dispon.exito) return dispon;
    if (dispon.slots.indexOf(hora) === -1) {
      return { exito: false, mensaje: 'Ese horario ya no está disponible. Elija otro.',
        fecha: fecha, duracion: duracion, sugerencias: dispon.slots };
    }

    var idCliente = _buscarOCrearClientePublico_(nombre, apellido, email, telefono);
    var titulo = cfg.ETIQUETA_CITA || CONFIGURACION_PREDETERMINADA.ETIQUETA_CITA || 'Cita';

    // Descripción con el desglose de servicios (si los hay).
    var descripcion = 'Reservada en línea por el cliente.';
    if (selServicios.items.length > 0) {
      descripcion = 'Reservada en línea por el cliente. Servicios:\n' +
        selServicios.items.map(function(it) {
          return '- ' + it.cantidad + ' × ' + it.nombre +
            ' (' + _formatoMoneda_(it.precio) + ' c/u = ' + _formatoMoneda_(it.subtotal) + ')';
        }).join('\n') + '\nTOTAL: ' + _formatoMoneda_(selServicios.total);
    }

    // Evento de Calendar.
    var partesFecha = fecha.split('-');
    var partesHora = hora.split(':');
    var inicio = new Date(
      parseInt(partesFecha[0], 10), parseInt(partesFecha[1], 10) - 1, parseInt(partesFecha[2], 10),
      parseInt(partesHora[0], 10), parseInt(partesHora[1], 10), 0);
    var fin = new Date(inicio.getTime() + duracion * 60000);
    var idEvento = '';
    var errorCalendario = '';
    try {
      var opcionesEvento = { description: descripcion };
      // El dueño es el organizador; el cliente recibe la invitación como invitado.
      if (email && esEmailValido_(email)) {
        opcionesEvento.guests = email;
      }
      idEvento = CalendarApp.getDefaultCalendar()
        .createEvent(titulo, inicio, fin, opcionesEvento).getId();
    } catch (errCal) {
      errorCalendario = errCal.message || String(errCal);
      Logger.log('Advertencia: no se pudo crear el evento de reserva: ' + errorCalendario);
    }

    // Fila en la hoja Citas (Fecha y Hora como texto).
    var hoja = obtenerHoja_(HOJA_CITAS);
    var idCita = generarId('CITA');
    hoja.appendRow([
      idCita, idCliente, titulo, fecha, hora, duracion, descripcion, idEvento, 'Programada',
      selServicios.items.length > 0 ? JSON.stringify(selServicios.items) : '',
      selServicios.items.length > 0 ? selServicios.total : '',
      'Cliente (en línea)'
    ]);
    var filaCita = hoja.getLastRow();
    var colFechaC = ENCABEZADOS.Citas.indexOf('Fecha') + 1;
    var colHoraC = ENCABEZADOS.Citas.indexOf('Hora') + 1;
    hoja.getRange(filaCita, colFechaC).setNumberFormat('@').setValue(fecha);
    hoja.getRange(filaCita, colHoraC).setNumberFormat('@').setValue(hora);

    // Confirmación por correo (la "constancia" del horario reservado).
    var avisoCorreo = '';
    try {
      if (email) {
        var cuerpo = 'Estimado/a ' + (nombre || 'cliente') + ':\n\n' +
          'Su ' + titulo.toLowerCase() + ' ha sido ' +
          (_etiquetaFemenina_(titulo) ? 'reservada' : 'reservado') + ' con los siguientes datos:\n' +
          (cfg.NOMBRE_NEGOCIO ? 'Empresa: ' + cfg.NOMBRE_NEGOCIO + '\n' : '') +
          'Fecha: ' + fecha + '\n' +
          'Hora: ' + _hora12_(hora) + '\n' +
          'Duración: ' + duracion + ' min\n';
        if (selServicios.items.length > 0) {
          cuerpo += '\nServicios:\n' + selServicios.items.map(function(it) {
            return '- ' + it.cantidad + ' × ' + it.nombre + ': ' + _formatoMoneda_(it.subtotal);
          }).join('\n') + '\nTOTAL: ' + _formatoMoneda_(selServicios.total) + '\n';
        }
        cuerpo += '\nLe esperamos. Gracias por preferirnos.';
        MailApp.sendEmail(email, titulo + ' ' + (_etiquetaFemenina_(titulo) ? 'reservada' : 'reservado'), cuerpo);
        avisoCorreo = ' Confirmación enviada a su correo.';
      }
    } catch (errCorreo) {
      avisoCorreo = ' No se pudo enviar la confirmación por correo.';
      Logger.log('Error al enviar correo de reserva: ' + errCorreo);
    }

    var avisoEquipo = '';
    try {
      var destinatariosEquipo = _correosAvisoCita_(null);
      if (destinatariosEquipo.length > 0) {
        var cuerpoEquipo =
          (cfg.NOMBRE_NEGOCIO ? 'Empresa: ' + cfg.NOMBRE_NEGOCIO + '\n' : '') +
          'Se recibió una nueva reserva en línea.\n\n' +
          'Cliente: ' + (nombre || '') + (apellido ? ' ' + apellido : '') + '\n' +
          (email ? 'Correo: ' + email + '\n' : '') +
          (telefono ? 'Teléfono: ' + telefono + '\n' : '') +
          'Fecha: ' + fecha + '\n' +
          'Hora: ' + _hora12_(hora) + '\n' +
          'Duración: ' + duracion + ' min\n' +
          'ID de cita: ' + idCita + '\n';
        if (selServicios.items.length > 0) {
          cuerpoEquipo += '\nServicios:\n' + selServicios.items.map(function(it) {
            return '- ' + it.cantidad + ' × ' + it.nombre + ': ' + _formatoMoneda_(it.subtotal);
          }).join('\n') + '\nTOTAL: ' + _formatoMoneda_(selServicios.total) + '\n';
        }
        MailApp.sendEmail(destinatariosEquipo.join(','), 'Nueva reserva en línea — ' + fecha + ' ' + _hora12_(hora), cuerpoEquipo);
        avisoEquipo = ' Aviso enviado al equipo.';
      }
    } catch (errEquipo) {
      Logger.log('Error al enviar aviso de reserva al equipo: ' + errEquipo);
    }

    Logger.log('Reserva pública creada: ' + idCita);
    return {
      exito: true,
      fecha: fecha,
      hora: hora,
      duracion: duracion,
      titulo: titulo,
      idCita: idCita,
      servicios: selServicios.items,
      totalPrecio: selServicios.total,
      mensaje: (idEvento
        ? 'Reserva confirmada y añadida al calendario.'
        : 'Reserva confirmada, pero no se pudo añadir al calendario: ' + errorCalendario) + avisoCorreo + avisoEquipo
    };
  } catch (err) {
    Logger.log('Error en reservarCitaPublica: ' + err);
    return { exito: false, mensaje: 'Error al reservar: ' + err.message };
  } finally {
    candado.releaseLock();
  }
}

/**
 * Devuelve la URL pública de la página de reservas ("…/exec?v=reservar").
 * Sin token: solo contiene la dirección pública.
 * @return {string}
 */
function obtenerUrlReservas() {
  var base = obtenerUrlAplicacion();
  if (!base) return '';
  return base + (base.indexOf('?') === -1 ? '?' : '&') + 'v=reservar';
}

/**
 * ============================================================
 *  HISTORIAL
 * ============================================================
 */

/**
 * Devuelve el historial de un cliente (o todo si no se indica).
 */
function obtenerHistorial(token, idCliente) {
  if (!_validarSesion_(token)) return _respuestaSesionExpirada_();
  _repararHistorialUnaVez_();
  var hoja = obtenerHoja_(HOJA_HISTORIAL);
  var todos = filasAObjetos_(hoja);
  if (!idCliente) {
    return todos;
  }
  return todos.filter(function(h) {
    return String(h.ID_Cliente) === String(idCliente);
  });
}

/**
 * Elimina un registro del historial por su ID (solo administrador).
 * @param {string} token
 * @param {string} idRegistro  ID_Registro.
 * @return {Object}
 */
function eliminarRegistroHistorial(token, idRegistro) {
  try {
    var usuario = _validarSesion_(token);
    if (!usuario) return _respuestaSesionExpirada_();
    if (!_esAdmin_(usuario)) return { exito: false, mensaje: 'Solo el administrador puede eliminar registros del historial.' };
    var hoja = obtenerHoja_(HOJA_HISTORIAL);
    var fila = buscarFilaPorId_(hoja, 'ID_Registro', idRegistro);
    if (!fila) return { exito: false, mensaje: 'No se encontró el registro indicado.' };
    hoja.deleteRow(fila);
    Logger.log('Registro de historial eliminado: ' + idRegistro);
    _registrarActividad_(token, 'Historial', 'Eliminó registro de historial', 'ID ' + idRegistro);
    return { exito: true, mensaje: 'Registro eliminado correctamente.' };
  } catch (err) {
    Logger.log('Error al eliminar registro de historial: ' + err);
    return { exito: false, mensaje: 'Error al eliminar el registro: ' + err.message };
  }
}

/**
 * Genera un PDF del historial a partir de las filas ya preparadas en el cliente.
 * Usa una hoja de cálculo temporal (SpreadsheetApp, ya autorizado) y la convierte
 * a PDF; luego manda la hoja a la papelera. No requiere nuevos permisos.
 *
 * @param {string} token Sesión del usuario.
 * @param {Array<string>} encabezados Fila de encabezados.
 * @param {Array<Array<string>>} filas Filas de datos (texto plano).
 * @return {Object} { exito, base64, nombre } o { exito:false, mensaje }.
 */
function generarPDFHistorial(token, encabezados, filas) {
  if (!_validarSesion_(token)) return _respuestaSesionExpirada_();
  var doc = null;
  try {
    var cfg = obtenerConfiguracion();
    var nombre = (cfg.NOMBRE_NEGOCIO || 'Historial').toString();
    var colorPrim = (cfg.COLOR_PRIMARIO || '#4285F4').toString();
    var logoUrl = (cfg.LOGO_URL || '').toString();

    doc = DocumentApp.create('Historial_' + new Date().getTime());
    var body = doc.getBody();

    // Logo arriba, a la izquierda (se omite si falla).
    if (logoUrl) {
      try {
        var blob = UrlFetchApp.fetch(logoUrl).getBlob();
        var img = body.appendImage(blob);
        var w = img.getWidth();
        var h = img.getHeight();
        var maxW = 130;
        if (w > 0) img.setWidth(maxW).setHeight(Math.round(h * (maxW / w)));
        body.appendParagraph('');
      } catch (eImg) { /* logo omitido */ }
    }

    // Nombre del negocio en el color de marca (título).
    var pNombre = body.appendParagraph(nombre);
    pNombre.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
    pNombre.editAsText().setBold(true).setFontSize(16).setForegroundColor(colorPrim);

    // Tabla: encabezados + datos.
    var matriz = [encabezados].concat(filas || []);
    var tabla = body.appendTable(matriz);
    tabla.setBorderColor('#CCCCCC');

    // Cabecera: fondo color de marca, texto blanco, tamaño moderado y padding ajustado.
    var filaCab = tabla.getRow(0);
    for (var c = 0; c < encabezados.length; c++) {
      var celda = filaCab.getCell(c);
      celda.setBackgroundColor(colorPrim);
      celda.setPaddingTop(3).setPaddingBottom(3).setPaddingLeft(4).setPaddingRight(4);
      var parr = celda.getChild(0).asParagraph();
      parr.editAsText().setForegroundColor('#FFFFFF').setBold(true).setFontSize(10);
    }

    // Filas de datos: texto oscuro y tamaño normal (evita heredar color/marca).
    for (var r = 1; r < tabla.getNumRows(); r++) {
      var filaD = tabla.getRow(r);
      for (var c2 = 0; c2 < encabezados.length; c2++) {
        var celdaD = filaD.getCell(c2);
        celdaD.setPaddingTop(2).setPaddingBottom(2).setPaddingLeft(4).setPaddingRight(4);
        var parrD = celdaD.getChild(0).asParagraph();
        parrD.editAsText().setForegroundColor('#222222').setBold(false).setFontSize(10);
      }
    }

    doc.saveAndClose();
    var pdf = DriveApp.getFileById(doc.getId()).getAs('application/pdf');
    var base64 = Utilities.base64Encode(pdf.getBytes());
    return { exito: true, base64: base64, nombre: 'historial.pdf' };
  } catch (e) {
    return { exito: false, mensaje: e.message };
  } finally {
    if (doc) {
      try { DriveApp.getFileById(doc.getId()).setTrashed(true); } catch (e2) {}
    }
  }
}

/**
 * Genera un archivo de Excel (.xlsx) del historial a partir de las filas ya
 * preparadas en el cliente. Se construye como un paquete OOXML válido
 * (zip de partes XML con celdas de texto "inline"), SIN depender de la
 * conversión de Drive (que no permite hoja nativa -> XLSX). No requiere
 * nuevos permisos ni crear archivos temporales.
 *
 * @param {string} token Sesión del usuario.
 * @param {Array<string>} encabezados Fila de encabezados.
 * @param {Array<Array<string>>} filas Filas de datos (texto plano).
 * @return {Object} { exito, base64, nombre } o { exito:false, mensaje }.
 */
function generarExcelHistorial(token, encabezados, filas) {
  if (!_validarSesion_(token)) return _respuestaSesionExpirada_();
  try {
    var cfg = obtenerConfiguracion();
    var nombre = (cfg.NOMBRE_NEGOCIO || 'Historial').toString();
    var colorPrim = (cfg.COLOR_PRIMARIO || '#4285F4').toString();
    var argb = _hexAArgb_(colorPrim);

    var nCols = encabezados.length;
    var ultCol = _columnaExcel_(nCols);
    var totalFilas = 2 + (filas ? filas.length : 0); // título + cabecera + datos

    // Hoja (sheet1.xml) con título de marca, cabecera coloreada y datos.
    var sheetXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<dimension ref="A1:' + ultCol + totalFilas + '"/>' +
      '<cols><col min="1" max="' + nCols + '" width="22" customWidth="1"/></cols>' +
      '<sheetData>';

    // Fila 1: título (nombre del negocio) con estilo 2.
    sheetXml += '<row r="1"><c r="A1" s="2" t="inlineStr"><is><t xml:space="preserve">' +
      _escaparXml_(nombre) + '</t></is></c></row>';

    // Fila 2: encabezados con estilo 1 (color de marca).
    sheetXml += '<row r="2">';
    for (var c = 0; c < encabezados.length; c++) {
      var refH = _columnaExcel_(c + 1) + '2';
      var valH = (encabezados[c] == null) ? '' : String(encabezados[c]);
      sheetXml += '<c r="' + refH + '" s="1" t="inlineStr"><is><t xml:space="preserve">' +
        _escaparXml_(valH) + '</t></is></c>';
    }
    sheetXml += '</row>';

    // Filas 3+: datos.
    if (filas) {
      for (var r = 0; r < filas.length; r++) {
        var fila = filas[r];
        sheetXml += '<row r="' + (r + 3) + '">';
        for (var c2 = 0; c2 < fila.length; c2++) {
          var refD = _columnaExcel_(c2 + 1) + (r + 3);
          var valD = (fila[c2] == null) ? '' : String(fila[c2]);
          sheetXml += '<c r="' + refD + '" t="inlineStr"><is><t xml:space="preserve">' +
            _escaparXml_(valD) + '</t></is></c>';
        }
        sheetXml += '</row>';
      }
    }
    sheetXml += '</sheetData>';
    sheetXml += '<mergeCells count="1"><mergeCell ref="A1:' + ultCol + '1"/></mergeCells>';
    sheetXml += '</worksheet>';

    // Estilos: fuentes blancas negrita + fill de color de marca.
    var stylesXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<fonts count="3">' +
      '<font><sz val="11"/><name val="Calibri"/></font>' +
      '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>' +
      '<font><b/><sz val="14"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>' +
      '</fonts>' +
      '<fills count="3">' +
      '<fill><patternFill patternType="none"/></fill>' +
      '<fill><patternFill patternType="gray125"/></fill>' +
      '<fill><patternFill patternType="solid"><fgColor rgb="' + argb + '"/><bgColor indexed="64"/></patternFill></fill>' +
      '</fills>' +
      '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      '<cellXfs count="3">' +
      '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
      '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>' +
      '<xf numFmtId="0" fontId="2" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>' +
      '</cellXfs>' +
      '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
      '</styleSheet>';

    var contentTypes = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
      '</Types>';

    var rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
      '</Relationships>';

    var workbook = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<sheets><sheet name="Historial" sheetId="1" r:id="rId1"/></sheets>' +
      '</workbook>';

    var workbookRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
      '</Relationships>';

    var partes = [
      Utilities.newBlob(contentTypes, 'application/xml', '[Content_Types].xml'),
      Utilities.newBlob(rels, 'application/xml', '_rels/.rels'),
      Utilities.newBlob(workbook, 'application/xml', 'xl/workbook.xml'),
      Utilities.newBlob(workbookRels, 'application/xml', 'xl/_rels/workbook.xml.rels'),
      Utilities.newBlob(stylesXml, 'application/xml', 'xl/styles.xml'),
      Utilities.newBlob(sheetXml, 'application/xml', 'xl/worksheets/sheet1.xml')
    ];
    var zip = Utilities.zip(partes);
    var base64 = Utilities.base64Encode(zip.getBytes());
    return { exito: true, base64: base64, nombre: 'historial.xlsx' };
  } catch (e) {
    return { exito: false, mensaje: e.message };
  }
}

/**
 * Registra una actividad de usuario en la hoja "Actividad" (auditoría) para
 * alimentar el monitor exclusivo del dueño. Es TOLERANTE A FALLOS: cualquier
 * error se envía a los logs pero NUNCA interrumpe la operación que lo invocó.
 *
 * @param {string} token    Token de sesión (puede ser null si no aplica).
 * @param {string} modulo   Nombre del módulo (Clientes, Citas, Usuarios, etc.).
 * @param {string} accion   Verbo de la acción (Agregó, Editó, Eliminó…).
 * @param {string} detalle  Descripción legible de lo ocurrido.
 * @param {string} [actor]  Correo del actor cuando no hay token (p. ej. auto-registro).
 */
function _registrarActividad_(token, modulo, accion, detalle, actor) {
  try {
    var cfg = obtenerConfiguracion();
    if (String(cfg.HABILITAR_MONITOR || 'NO').toUpperCase() !== 'SI') return;

    var email = actor || _validarSesion_(token);
    if (!email) return;

    var nombre = email;
    var rol = '';
    var u = _buscarUsuario_(email);
    if (u) {
      nombre = u.Nombre || email;
      rol = u.Rol || '';
    }

    var hoja = obtenerHoja_(HOJA_ACTIVIDAD);
    var zona = obtenerZonaHoraria_();
    hoja.appendRow([
      generarId('ACT'),
      Utilities.formatDate(new Date(), zona, 'yyyy-MM-dd HH:mm'),
      nombre,
      email,
      rol,
      modulo,
      accion,
      detalle || ''
    ]);

    // Limitar el tamaño de la hoja a las últimas MAX_ACTIVIDAD filas de datos.
    var total = hoja.getLastRow() - 1;
    if (total > MAX_ACTIVIDAD) {
      hoja.deleteRows(2, total - MAX_ACTIVIDAD);
    }
  } catch (e) {
    Logger.log('Error al registrar actividad (se ignora): ' + e);
  }
}

/**
 * Devuelve el registro de actividad para el monitor del dueño (primer admin).
 * Solo el dueño puede consultarlo y únicamente si el monitor está habilitado.
 *
 * @param {string} token    Token de sesión.
 * @param {Object} [filtros] { usuario, modulo, desde, hasta }
 * @return {Object} { exito, actividades } o { exito:false, mensaje }.
 */
function obtenerActividad(token, filtros) {
  var email = _validarSesion_(token);
  if (!email) return _respuestaSesionExpirada_();
  try {
    var cfg = obtenerConfiguracion();
    if (String(cfg.HABILITAR_MONITOR || 'NO').toUpperCase() !== 'SI') {
      return { exito: false, mensaje: 'El monitor de actividad está desactivado.', desactivado: true, habilitado: false };
    }
    if (!_esUsuarioPrincipal_(email)) {
      return { exito: false, mensaje: 'Esta función es exclusiva del dueño.' };
    }
    var hoja = obtenerHoja_(HOJA_ACTIVIDAD);
    var filas = filasAObjetos_(hoja);
    var totalAct = filas.length;
    filas.reverse(); // más recientes primero
    filtros = filtros || {};
    if (filtros.usuario) {
      var uF = String(filtros.usuario).toLowerCase();
      filas = filas.filter(function(a) {
        return String(a.Email || '').toLowerCase() === uF ||
               String(a.Usuario || '').toLowerCase() === uF;
      });
    }
    if (filtros.modulo) {
      filas = filas.filter(function(a) { return String(a.Modulo || '') === String(filtros.modulo); });
    }
    if (filtros.desde) {
      var dDesde = new Date(filtros.desde);
      filas = filas.filter(function(a) { return new Date(a.Fecha) >= dDesde; });
    }
    if (filtros.hasta) {
      var dHasta = new Date(filtros.hasta);
      filas = filas.filter(function(a) { return new Date(a.Fecha) <= dHasta; });
    }
    return { exito: true, actividades: filas.slice(0, 500), total: totalAct, habilitado: true };
  } catch (err) {
    Logger.log('Error en obtenerActividad: ' + err);
    return { exito: false, mensaje: 'Error al leer la actividad: ' + err.message };
  }
}

/** Escapa caracteres especiales de XML. */
function _escaparXml_(texto) {
  return String(texto)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

/** Convierte un índice de columna (1-based) a su letra de Excel (1->A, 27->AA). */
function _columnaExcel_(n) {
  var s = '';
  while (n > 0) {
    var m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/** Convierte #RRGGBB (o #RGB) a ARGB (FFRRGGBB) para atributos rgb de OOXML. */
function _hexAArgb_(hex) {
  var h = String(hex || '').replace('#', '').trim();
  if (h.length === 3) h = h[0] + h[0] + h[1] + h[1] + h[2] + h[2];
  if (h.length !== 6 || /[^0-9a-fA-F]/.test(h)) return 'FF4285F4';
  return 'FF' + h.toUpperCase();
}

/**
 * Reescribe la descripción de los registros existentes del historial usando el
 * formato limpio (sin TOTAL embebido ni desglose duplicado). Ejecutar una sola vez.
 */
function repararHistorial() {
  try {
    var hojaH = obtenerHoja_(HOJA_HISTORIAL);
    var cabH = ENCABEZADOS[HOJA_HISTORIAL];
    var historial = filasAObjetos_(hojaH);
    var hojaC = obtenerHoja_(HOJA_CITAS);
    var citas = filasAObjetos_(hojaC);
    var mapaCitas = {};
    citas.forEach(function(c) { if (c.ID_Cita) mapaCitas[String(c.ID_Cita)] = c; });
    var colDesc = cabH.indexOf('Descripcion') + 1;
    var actualizados = 0;
    for (var i = 0; i < historial.length; i++) {
      var h = historial[i];
      var cita = mapaCitas[String(h.ID_Cita || '')];
      if (!cita) continue;
      var valoresCita = [
        cita.ID_Cita, cita.ID_Cliente, cita.Titulo, cita.Fecha, cita.Hora,
        cita.Duracion_Mins, cita.Descripcion, cita.ID_Evento_Calendar, cita.Estado,
        cita.Servicios, cita.Total_Precio, cita.Agendado_Por
      ];
      var descLimpia = _descripcionHistorial_(valoresCita);
      if (descLimpia !== h.Descripcion) {
        hojaH.getRange(i + 2, colDesc, 1, 1).setValue(descLimpia);
        actualizados++;
      }
    }
    return { exito: true, mensaje: 'Historial reparado. Registros actualizados: ' + actualizados + '.' };
  } catch (e) {
    return { exito: false, mensaje: 'Error al reparar el historial: ' + e.message };
  }
}

// Repara el historial existente una sola vez (la primera vez que se consulta).
function _repararHistorialUnaVez_() {
  try {
    var props = PropertiesService.getScriptProperties();
    if (props.getProperty('historialReparadoV2')) return;
    var res = repararHistorial();
    if (res && res.exito) props.setProperty('historialReparadoV2', '1');
  } catch (e) {
    Logger.log('No se pudo reparar el historial automáticamente: ' + e);
  }
}

/**
 * ============================================================
 *  PROTECCIÓN DE LA BASE DE DATOS
 * ============================================================
 */

/**
 * Protege las hojas de datos (Clientes, Citas, Historial, Configuracion y
 * Usuarios) para que solo se puedan editar desde la aplicación web (el script
 * corre como el dueño y sigue escribiendo).
 * Los usuarios finales ya no podrán editar la hoja directamente.
 */
function protegerBaseDatos(token) {
  if (!_esContextoEditor_() && !_tieneHojaActiva_() && !_validarSesion_(token)) {
    return _respuestaSesionExpirada_();
  }
  try {
    var hojaCalculo = obtenerHojaCalculo_();
    var hojas = [HOJA_CLIENTES, HOJA_CITAS, HOJA_HISTORIAL, HOJA_CONFIGURACION, HOJA_USUARIOS, HOJA_SERVICIOS, HOJA_ACTIVIDAD];
    // Asegurar que la hoja de actividad exista para que quede protegida aunque
    // aún no se haya registrado ninguna acción.
    try { obtenerHoja_(HOJA_ACTIVIDAD); } catch (eAct) {}
    var protegidas = 0;
    var yaProtegidas = 0;
    hojas.forEach(function(nombre) {
      var hoja = hojaCalculo.getSheetByName(nombre);
      if (!hoja) return;
      // Si ya está protegida, contarla y no hacer nada.
      var existente = hoja.getProtections(SpreadsheetApp.ProtectionType.SHEET);
      if (existente.length > 0) {
        yaProtegidas++;
        return;
      }
      var proteccion = hoja.protect();
      proteccion.setWarningOnly(false);
      var editores = proteccion.getEditors();
      proteccion.removeEditors(editores);
      // Mantener como editores a la identidad que ejecuta el script (admin)
      // y al dueño de la hoja de cálculo, para que la aplicación web siga
      // pudiendo escribir en hojas protegidas sin importar quién ejecuta el
      // ítem del menú (el desplegador suele ser el dueño).
      var identidades = [];
      var identidad = Session.getEffectiveUser().getEmail();
      if (identidad) identidades.push(identidad);
      var dueno = hojaCalculo.getOwner();
      var emailDueno = dueno ? dueno.getEmail() : '';
      if (emailDueno && identidades.indexOf(emailDueno) === -1) identidades.push(emailDueno);
      identidades.forEach(function(email) {
        proteccion.addEditor(email);
      });
      protegidas++;
      Logger.log('Hoja protegida: ' + nombre);
    });
    var msg;
    if (protegidas === 0 && yaProtegidas === 0) {
      msg = 'No se encontró ninguna hoja de la base de datos para proteger.';
    } else {
      msg = 'Base de datos protegida: ' + protegidas + ' hoja(s) nuevas';
      if (yaProtegidas > 0) {
        msg += ' · ' + yaProtegidas + ' ya estaban protegidas';
      }
      msg += '. Los usuarios solo editan desde la aplicación web.';
    }
    _mostrarMensaje_('Protección de datos', msg);
    return { exito: true, mensaje: msg };
  } catch (err) {
    Logger.log('Error al proteger la base de datos: ' + err);
    var msgErr = 'Error al proteger la base de datos: ' + err.message;
    _mostrarMensaje_('Protección de datos', msgErr);
    return { exito: false, mensaje: msgErr };
  }
}

/**
 * Quita la protección de todas las hojas de la base de datos.
 */
function desprotegerBaseDatos(token) {
  if (!_esContextoEditor_() && !_tieneHojaActiva_() && !_validarSesion_(token)) {
    return _respuestaSesionExpirada_();
  }
  try {
    var hojaCalculo = obtenerHojaCalculo_();
    // Quitar SOLO las protecciones de las hojas de la aplicación; no toca
    // protecciones manuales de otras hojas o rangos creadas por el usuario.
    var hojas = [HOJA_CLIENTES, HOJA_CITAS, HOJA_HISTORIAL, HOJA_CONFIGURACION, HOJA_USUARIOS, HOJA_SERVICIOS, HOJA_ACTIVIDAD];
    var quitadas = 0;
    hojas.forEach(function(nombre) {
      var hoja = hojaCalculo.getSheetByName(nombre);
      if (!hoja) return;
      var protecciones = hoja.getProtections(SpreadsheetApp.ProtectionType.SHEET);
      protecciones.forEach(function(proteccion) {
        proteccion.remove();
        quitadas++;
      });
    });
    var msg = 'Protección eliminada de ' + quitadas + ' hoja(s) de la base de datos. Ahora se pueden editar directamente.';
    _mostrarMensaje_('Protección de datos', msg);
    return { exito: true, mensaje: msg };
  } catch (err) {
    Logger.log('Error al desproteger la base de datos: ' + err);
    var msgErr = 'Error al desproteger la base de datos: ' + err.message;
    _mostrarMensaje_('Protección de datos', msgErr);
    return { exito: false, mensaje: msgErr };
  }
}
