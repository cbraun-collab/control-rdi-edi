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
 * Busca en profundidad, dentro de cada carpeta de año, el código de proyecto exacto (no
 * como substring de otro código: usa límites de palabra). Recorre también carpetas
 * intermedias (ej. carpetas de cliente como "002 SALFA") para encontrar el proyecto sin
 * importar cuántos niveles de organización haya antes de llegar a él. Devuelve TODAS las
 * coincidencias encontradas, para poder guardar en cada una si el código está en más de
 * un lugar.
 */
function buscarTodasCarpetasProyecto_(codigoProyecto) {
  const raiz = DriveApp.getFolderById(CARPETA_RAIZ_OBRAS_ID);
  const carpetasAnio = raiz.getFolders();
  const coincidencias = [];
  const codigoEscapado = codigoProyecto.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const patron = new RegExp('(^|[^0-9A-Za-z])' + codigoEscapado + '([^0-9A-Za-z]|$)', 'i');

  while (carpetasAnio.hasNext()) {
    const carpetaAnio = carpetasAnio.next();
    buscarCodigoRecursivo_(carpetaAnio, carpetaAnio.getName(), patron, coincidencias, 0);
  }
  return coincidencias;
}

// Recorre en profundidad (máximo 4 niveles) buscando carpetas cuyo nombre coincida con el
// patrón del código de proyecto. No sigue bajando dentro de una carpeta que ya coincidió
// (el proyecto no puede estar anidado dentro de sí mismo).
function buscarCodigoRecursivo_(carpetaPadre, rutaActual, patron, coincidencias, profundidad) {
  if (profundidad > 4) return;
  const subcarpetas = carpetaPadre.getFolders();
  while (subcarpetas.hasNext()) {
    const carpeta = subcarpetas.next();
    const ruta = rutaActual + ' / ' + carpeta.getName();
    if (patron.test(carpeta.getName())) {
      coincidencias.push({ carpeta: carpeta, ruta: ruta });
    } else {
      buscarCodigoRecursivo_(carpeta, ruta, patron, coincidencias, profundidad + 1);
    }
  }
}

/**
 * Recorre la carpeta raíz de obras -> cada carpeta de año -> busca TODAS las carpetas cuyo
 * nombre contiene el código de proyecto (como token completo), y dentro de cada una busca
 * la subcarpeta "03 - RDI EDI" en cualquier nivel de profundidad.
 * Si el mismo código existe en más de un lugar (caso conocido: "03 Proyectos 2026" y
 * "PROYECTOS 2026" duplicadas a propósito mientras se define cuál es la oficial), se
 * devuelven TODAS las subcarpetas encontradas para que el registro se guarde en cada una
 * y no se pierda información sea cual sea la carpeta que finalmente se use.
 */
function buscarCarpetaRdiEdi_(codigoProyecto) {
  const coincidencias = buscarTodasCarpetasProyecto_(codigoProyecto);
  const carpetasRdiEdi = [];

  coincidencias.forEach((c) => {
    const encontrada = buscarSubcarpetaRecursiva_(c.carpeta, NOMBRE_SUBCARPETA_RDI_EDI, 0);
    if (encontrada) {
      carpetasRdiEdi.push({ id: encontrada.getId(), ruta: c.ruta });
    } else {
      Logger.log('Se encontró la carpeta del proyecto en "' + c.ruta + '" pero no tiene subcarpeta "' + NOMBRE_SUBCARPETA_RDI_EDI + '"');
    }
  });

  return { ok: true, carpetas: carpetasRdiEdi };
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
 * Valida que el código de proyecto tenga al menos una carpeta con subcarpeta "03 - RDI EDI".
 * Si hay más de una (código duplicado en distintas carpetas de año), las devuelve todas
 * para que el registro se guarde en cada una.
 */
function validarCarpetaProyecto_(codigoProyecto) {
  const busqueda = buscarCarpetaRdiEdi_(codigoProyecto);
  if (busqueda.carpetas.length === 0) {
    return { ok: false, error: 'No se encontró la carpeta "03 - RDI EDI" para el código ' + codigoProyecto + '. Verifica que el proyecto exista y tenga sus subcarpetas creadas.' };
  }
  return { ok: true, carpetas: busqueda.carpetas };
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
      resultado = {
        ok: true,
        existe: busqueda.carpetas.length > 0,
        multiple: busqueda.carpetas.length > 1,
        ubicaciones: busqueda.carpetas.map((c) => c.ruta)
      };
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
  const carpetaIds = validacion.carpetas.map((c) => c.id);

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

  return { ok: true, id: id, numero: numero, carpetaIds: carpetaIds };
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
  const carpetaIds = validacion.carpetas.map((c) => c.id);

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

  return { ok: true, id: id, numero: numero, carpetaIds: carpetaIds };
}

// ==================== UTILIDADES ====================
/**
 * Borra todas las filas de prueba de REGISTROS (deja solo el encabezado) y reinicia el
 * correlativo. Ejecutar UNA SOLA VEZ a mano desde el editor cuando se quiera limpiar
 * datos de prueba antes de empezar a usar la app en serio.
 */
function limpiarDatosDePrueba() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const registros = ss.getSheetByName('REGISTROS');
  if (registros.getLastRow() > 1) {
    registros.deleteRows(2, registros.getLastRow() - 1);
  }
  const correlativos = ss.getSheetByName('CORRELATIVOS');
  if (correlativos.getLastRow() > 1) {
    correlativos.deleteRows(2, correlativos.getLastRow() - 1);
  }
  Logger.log('Datos de prueba eliminados. Los próximos RDI/EDI partirán desde el número 001.');
}

function respuestaJsonp_(callback, resultado) {
  const json = JSON.stringify(resultado);
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}
