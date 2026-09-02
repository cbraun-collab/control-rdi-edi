/**
 * PWA-COM-005 · Control RDI / EDI Senercom
 * Backend Apps Script — Paso 2: Sheet maestro + correlativo por proyecto/tipo + búsqueda de carpeta de obra
 *
 * IMPORTANTE: este script se despliega desde carlos.braun@gmail.com (cuenta personal),
 * NUNCA desde cbraun@senercom.cl, para evitar el problema de CORS ya conocido en las
 * demás PWA de Senercom. El Sheet puede vivir en cualquier cuenta con acceso compartido.
 */

// ==================== CONFIGURACIÓN ====================
const SHEET_ID = '17ocfHCDBRq9Lps1sdlNOQG2Y7jjjJp5W1ORKM7ERUMw';
const CARPETA_RAIZ_OBRAS_ID = '1Lm2X3Uz_H_sUkD7Zll0JjKPSdaa1EW9n'; // carpeta raíz que contiene las carpetas por año
const CLAVE_APP = 'senercom2026'; // TODO: reemplazar por la clave definitiva antes de liberar la app
const NOMBRE_SUBCARPETA_RDI_EDI = '03 - RDI EDI';

// ==================== INICIALIZACIÓN (ejecutar una sola vez, a mano, desde el editor) ====================
function inicializarEstructura() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  crearHojaSiNoExiste_(ss, 'REGISTROS', [
    'ID', 'Tipo', 'Codigo_Proyecto', 'Numero', 'Estado',
    'Tema', 'Area', 'Plano_Referencia', 'Materia', 'Prioridad', 'Fecha_Requerida_Respuesta', 'Descripcion', 'Incidencia',
    'Emisor_Nombre', 'Emisor_Cargo', 'Emisor_Fecha', 'Emisor_Firma_URL', 'Adjuntos_Emisor',
    'Token_Firma',
    'Receptor_Nombre', 'Receptor_RUT', 'Receptor_Cargo', 'Receptor_Fecha', 'Receptor_Firma_URL', 'Respuesta', 'Adjuntos_Receptor',
    'Cumple', 'Genera_Nueva_RDI', 'Genera_Modif_Obra', 'Responsable_Analisis', 'Revision',
    'PDF_Final_URL', 'Fecha_Creacion', 'Fecha_Cierre'
  ]);

  crearHojaSiNoExiste_(ss, 'CORRELATIVOS', [
    'Codigo_Proyecto', 'Tipo', 'Ultimo_Numero'
  ]);

  crearHojaSiNoExiste_(ss, 'CONFIG', [
    'Clave', 'Valor'
  ]);
  const configSheet = ss.getSheetByName('CONFIG');
  if (configSheet.getLastRow() < 2) {
    configSheet.getRange(2, 1, 2, 2).setValues([
      ['CARPETA_RAIZ_OBRAS_ID', CARPETA_RAIZ_OBRAS_ID],
      ['CLAVE_APP', CLAVE_APP]
    ]);
  }

  Logger.log('Estructura inicializada correctamente.');
}

function crearHojaSiNoExiste_(ss, nombre, headers) {
  let sheet = ss.getSheetByName(nombre);
  if (!sheet) {
    sheet = ss.insertSheet(nombre);
  }
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
}

// ==================== CORRELATIVO POR PROYECTO Y TIPO ====================
/**
 * Devuelve el siguiente número correlativo para un código de proyecto y tipo (RDI o EDI),
 * y deja registrado el nuevo último número. Ej: proyecto 001-001-O26-CB, tipo RDI -> "RDI-001", "RDI-002"...
 */
function siguienteCorrelativo_(codigoProyecto, tipo) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('CORRELATIVOS');
  const datos = sheet.getDataRange().getValues();

  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] === codigoProyecto && datos[i][1] === tipo) {
      const nuevoNumero = datos[i][2] + 1;
      sheet.getRange(i + 1, 3).setValue(nuevoNumero);
      return formatearNumero_(tipo, nuevoNumero);
    }
  }

  // No existe combinación previa: se crea partiendo de 1
  sheet.appendRow([codigoProyecto, tipo, 1]);
  return formatearNumero_(tipo, 1);
}

function formatearNumero_(tipo, numero) {
  const correlativo = String(numero).padStart(3, '0');
  return tipo + '-' + correlativo; // ej: RDI-001, EDI-014
}

// ==================== BÚSQUEDA DE LA CARPETA DE LA OBRA ====================
/**
 * Busca en TODAS las carpetas de año el código de proyecto exacto (no como substring de
 * otro código: usa límites de palabra para que "002-014-026-MA" no confunda con
 * "002-014-026-MA-2"). Devuelve TODAS las coincidencias encontradas, para poder avisar
 * si el código está duplicado en más de una carpeta.
 */
function buscarTodasCarpetasProyecto_(codigoProyecto) {
  const raiz = DriveApp.getFolderById(CARPETA_RAIZ_OBRAS_ID);
  const carpetasAnio = raiz.getFolders();
  const coincidencias = [];
  const codigoEscapado = codigoProyecto.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const patron = new RegExp('(^|[^0-9A-Za-z])' + codigoEscapado + '([^0-9A-Za-z]|$)', 'i');

  while (carpetasAnio.hasNext()) {
    const carpetaAnio = carpetasAnio.next();
    const subcarpetas = carpetaAnio.getFolders();
    while (subcarpetas.hasNext()) {
      const carpeta = subcarpetas.next();
      if (patron.test(carpeta.getName())) {
        coincidencias.push({
          carpeta: carpeta,
          ruta: carpetaAnio.getName() + ' / ' + carpeta.getName()
        });
      }
    }
  }
  return coincidencias;
}

/**
 * Recorre la carpeta raíz de obras -> cada carpeta de año -> busca la carpeta cuyo nombre
 * contiene exactamente el código de proyecto (como token completo, no substring parcial),
 * y dentro de ella busca la subcarpeta "03 - RDI EDI" en CUALQUIER nivel de profundidad.
 * Si el código aparece en más de una carpeta, devuelve un error de duplicado con el detalle
 * de dónde está cada una, para que se pueda revisar y corregir antes de cargar información.
 */
function buscarCarpetaRdiEdi_(codigoProyecto) {
  const coincidencias = buscarTodasCarpetasProyecto_(codigoProyecto);

  if (coincidencias.length === 0) {
    Logger.log('No se encontró ninguna carpeta de proyecto con el código ' + codigoProyecto);
    return { ok: true, carpetaId: null, duplicado: false };
  }

  if (coincidencias.length > 1) {
    return {
      ok: true,
      carpetaId: null,
      duplicado: true,
      ubicaciones: coincidencias.map((c) => c.ruta)
    };
  }

  const carpetaProyecto = coincidencias[0].carpeta;
  const encontrada = buscarSubcarpetaRecursiva_(carpetaProyecto, NOMBRE_SUBCARPETA_RDI_EDI, 0);
  if (!encontrada) {
    Logger.log('Se encontró la carpeta del proyecto ' + codigoProyecto + ' pero no tiene subcarpeta "' + NOMBRE_SUBCARPETA_RDI_EDI + '" en ningún nivel');
    return { ok: true, carpetaId: null, duplicado: false };
  }

  return { ok: true, carpetaId: encontrada.getId(), duplicado: false };
}

// Búsqueda en profundidad (máximo 4 niveles, suficiente para la estructura de obras) de una
// subcarpeta por nombre exacto dentro de una carpeta dada.
function buscarSubcarpetaRecursiva_(carpetaPadre, nombreBuscado, profundidad) {
  if (profundidad > 4) return null;
  const subcarpetas = carpetaPadre.getFolders();
  const pendientes = [];
  while (subcarpetas.hasNext()) {
    const carpeta = subcarpetas.next();
    if (carpeta.getName() === nombreBuscado) return carpeta;
    pendientes.push(carpeta);
  }
  for (let i = 0; i < pendientes.length; i++) {
    const resultado = buscarSubcarpetaRecursiva_(pendientes[i], nombreBuscado, profundidad + 1);
    if (resultado) return resultado;
  }
  return null;
}

/**
 * Valida que el código de proyecto tenga una única carpeta con subcarpeta "03 - RDI EDI".
 * Devuelve un error claro y accionable si no existe o si está duplicado.
 */
function validarCarpetaProyecto_(codigoProyecto) {
  const busqueda = buscarCarpetaRdiEdi_(codigoProyecto);

  if (busqueda.duplicado) {
    return {
      ok: false,
      error: 'El código ' + codigoProyecto + ' está duplicado en más de una carpeta. Revísalo antes de continuar:\n- ' + busqueda.ubicaciones.join('\n- ')
    };
  }
  if (!busqueda.carpetaId) {
    return { ok: false, error: 'No se encontró la carpeta "03 - RDI EDI" para el código ' + codigoProyecto + '. Verifica que el proyecto exista y tenga sus subcarpetas creadas.' };
  }
  return { ok: true, carpetaId: busqueda.carpetaId };
}

// ==================== ENDPOINTS (JSONP) ====================
function doGet(e) {
  const callback = e.parameter.callback;
  const accion = e.parameter.accion;
  let resultado;

  try {
    if (e.parameter.clave !== CLAVE_APP) {
      resultado = { ok: false, error: 'Clave de acceso inválida' };
    } else if (accion === 'verificarProyecto') {
      const busqueda = buscarCarpetaRdiEdi_(e.parameter.codigoProyecto);
      if (busqueda.duplicado) {
        resultado = { ok: true, existe: false, duplicado: true, ubicaciones: busqueda.ubicaciones };
      } else {
        resultado = { ok: true, existe: !!busqueda.carpetaId, carpetaId: busqueda.carpetaId, duplicado: false };
      }
    } else if (accion === 'ping') {
      resultado = { ok: true, mensaje: 'Backend Control RDI/EDI operativo' };
    } else {
      resultado = { ok: false, error: 'Acción no reconocida' };
    }
  } catch (err) {
    resultado = { ok: false, error: err.message };
  }

  return respuestaJsonp_(callback, resultado);
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents);
  let resultado;

  try {
    if (body.clave !== CLAVE_APP) {
      resultado = { ok: false, error: 'Clave de acceso inválida' };
    } else if (body.accion === 'crearRegistroBase') {
      resultado = crearRegistroBase_(body);
    } else if (body.accion === 'guardarSeccionA') {
      resultado = guardarSeccionA_(body);
    } else {
      resultado = { ok: false, error: 'Acción no reconocida' };
    }
  } catch (err) {
    resultado = { ok: false, error: err.message };
  }

  return ContentService.createTextOutput(JSON.stringify(resultado))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Crea la fila base del registro (tipo + código de proyecto + correlativo asignado).
 * El resto de los campos del formulario (Sección A completa, adjuntos, firma emisor, envío
 * de correo con link de firma) se conecta en los próximos pasos.
 */
function crearRegistroBase_(body) {
  const codigoProyecto = body.codigoProyecto;
  const tipo = body.tipo; // 'RDI' o 'EDI'

  const validacion = validarCarpetaProyecto_(codigoProyecto);
  if (!validacion.ok) return validacion;
  const carpetaId = validacion.carpetaId;

  const numero = siguienteCorrelativo_(codigoProyecto, tipo);
  const id = Utilities.getUuid();

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('REGISTROS');
  sheet.appendRow([
    id, tipo, codigoProyecto, numero, 'Borrador',
    '', '', '', '', '', '', '', '',
    '', '', '', '', '',
    '',
    '', '', '', '', '', '', '',
    '', '', '', '', '',
    '', new Date(), ''
  ]);

  return { ok: true, id: id, numero: numero, carpetaId: carpetaId };
}

/**
 * Guarda el RDI/EDI completo con toda la Sección A (emisor) ya llena y firmada.
 * El envío al receptor con el link de firma remota se conecta en el próximo paso.
 */
function guardarSeccionA_(body) {
  const codigoProyecto = body.codigoProyecto;
  const tipo = body.tipo;

  const validacion = validarCarpetaProyecto_(codigoProyecto);
  if (!validacion.ok) return validacion;
  const carpetaId = validacion.carpetaId;

  const numero = siguienteCorrelativo_(codigoProyecto, tipo);
  const id = Utilities.getUuid();

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('REGISTROS');
  sheet.appendRow([
    id, tipo, codigoProyecto, numero, 'Pendiente de envío',
    body.tema || '', body.area || '', body.planoReferencia || '', body.materia || '', body.prioridad || '', body.fechaRequerida || '', body.descripcion || '', body.incidencia || '',
    body.emisorNombre || '', body.emisorCargo || '', body.emisorFecha || '', body.emisorFirmaUrl || '', (body.adjuntosEmisor || []).join(', '),
    '', // Token_Firma (se genera en el paso de envío)
    '', '', '', '', '', '', '', // Receptor_*
    '', '', '', '', '', // Cumple / Genera_Nueva_RDI / Genera_Modif_Obra / Responsable_Analisis / Revision
    '', new Date(), '' // PDF_Final_URL, Fecha_Creacion, Fecha_Cierre
  ]);

  return { ok: true, id: id, numero: numero, carpetaId: carpetaId };
}

// ==================== UTILIDADES ====================
function respuestaJsonp_(callback, resultado) {
  const json = JSON.stringify(resultado);
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}
