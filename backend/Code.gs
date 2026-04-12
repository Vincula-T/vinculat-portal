// ================================================================
// VINCULAT — Apps Script Completo
// Versión 2.4 — Secretos movidos a Script Properties
// ================================================================
// CAMBIOS vs v2.3:
//   [SEC] SPREADSHEET_ID, EMAIL_ADMIN y FOLDER_* se leen desde
//         Script Properties (no mas secretos en el codigo fuente)
//   [SEC] Helper getProp() con error claro si falta configuracion
// ================================================================
//
// CONFIGURACION INICIAL (solo una vez):
//   Apps Script → ⚙️ Configuracion del proyecto → "Propiedades de
//   la secuencia de comandos" → agregar estas 4 propiedades:
//
//     SPREADSHEET_ID       = (ID del Google Sheet)
//     EMAIL_ADMIN          = (correo que recibe notificaciones)
//     FOLDER_COMPROBANTES  = (ID de la carpeta Drive de comprobantes)
//     FOLDER_FOTOS         = (ID de la carpeta Drive de fotos)
//
//   Si falta alguna, getProp() lanzara un error descriptivo.
//
// ================================================================

var SHEET_NAME = 'Pacientes JotForm';

function getProp(name) {
  var v = PropertiesService.getScriptProperties().getProperty(name);
  if (!v) {
    throw new Error('Falta Script Property: ' + name +
      '. Configurala en Apps Script → ⚙️ Configuracion del proyecto → Propiedades de la secuencia de comandos.');
  }
  return v;
}

function SPREADSHEET_ID()      { return getProp('SPREADSHEET_ID'); }
function EMAIL_ADMIN()         { return getProp('EMAIL_ADMIN'); }
function FOLDER_COMPROBANTES() { return getProp('FOLDER_COMPROBANTES'); }
function FOLDER_FOTOS()        { return getProp('FOLDER_FOTOS'); }

var COL_ESTADO        = 1;
var COL_FECHA_REG     = 2;
var COL_NOMBRE        = 3;
var COL_TEL           = 4;
var COL_EXPEDIENTE    = 5;
var COL_TRATAMIENTO   = 6;
var COL_EXTRACCIONES  = 7;
var COL_DIENTE        = 8;
var COL_RESINAS       = 9;
var COL_EDAD          = 10;
var COL_APARATO       = 11;
var COL_HORARIO_MAN   = 12;
var COL_HORARIO_TAR   = 13;
var COL_CONSENTIMIENTO = 14;
var COL_EXT_PROTESIS  = 15;
var COL_DOLOR         = 16;
var COL_FRACTURA      = 17;
var COL_ID            = 19;
var COL_ESTADO_PORTAL = 23;
var COL_ALUMNO        = 24;
var COL_FECHA_ASIG    = 25;
var COL_PAGO          = 26;
var COL_FOTO          = 27;
var COL_REF_PAGO      = 28;
var COL_COMPROBANTE   = 29;
var COL_TEL_ALUMNO    = 30;


function enviarAlertaVIP(e) {
  if (!e || !e.values) return;
  var respuestas     = e.values;
  var nombre         = respuestas[1] || 'Paciente';
  var telefono       = respuestas[2] || 'Sin numero';
  var motivoOriginal = respuestas[4] || '';
  var motivo         = motivoOriginal.toString().toLowerCase();
  var esCasoVIP = motivo.match(/fractura|frente|anterior|golpe|quebrado|trauma|oscuro/) ||
                  (motivo.match(/endo/) && !motivo.match(/molar|muela/));
  if (esCasoVIP) {
    MailApp.sendEmail(EMAIL_ADMIN(), 'ALERTA VIP: Posible Fractura/Endo',
      'Llego un paciente urgente:\n\nNombre: ' + nombre + '\nMotivo: ' + motivoOriginal + '\nTel: ' + telefono + '\n\nRevisa el panel admin.');
  }
}


function onEdit(e) {
  var sheet = e.source.getActiveSheet();
  var range = e.range;
  if (range.getColumn() !== 1) return;
  if (range.getRow() <= 1) return;
  var estado = range.getValue();
  var row    = range.getRow();
  sheet.getRange(row, COL_ESTADO_PORTAL).setValue(estado);
  if (estado === 'Asignado') {
    sheet.getRange(row, COL_FECHA_ASIG).setValue(new Date());
    sheet.getRange(row, COL_PAGO).setValue('Validado');
  }
  if (estado === 'Garantia') {
    var nombre = sheet.getRange(row, COL_NOMBRE).getValue();
    var idPac  = sheet.getRange(row, COL_ID).getValue();
    MailApp.sendEmail(EMAIL_ADMIN(), 'Garantia solicitada: ' + idPac,
      'Paciente: ' + nombre + ' | ID: ' + idPac + '\nRevisa el panel admin para validar.');
  }
  if (['Disponible', 'Sabado', 'Regalado', 'Contactado'].indexOf(estado) !== -1) {
    sheet.getRange(row, COL_FECHA_ASIG).clearContent();
    sheet.getRange(row, COL_ALUMNO).clearContent();
    sheet.getRange(row, COL_PAGO).clearContent();
    sheet.getRange(row, COL_REF_PAGO).clearContent();
    sheet.getRange(row, COL_COMPROBANTE).clearContent();
    sheet.getRange(row, COL_TEL_ALUMNO).clearContent();
  }
}


function doGet(e) {
  var sheet  = SpreadsheetApp.openById(SPREADSHEET_ID()).getSheetByName(SHEET_NAME);
  var accion = e.parameter.accion || '';
  if (accion === 'cambiarEstado') return accionCambiarEstado(e.parameter, sheet);
  if (accion === 'crearVIP')      return accionCrearVIP(e.parameter, sheet);
  if (accion === 'getPrecios')    return accionGetPrecios();
  if (accion === 'setPrecios')    return accionSetPrecios(e.parameter);
  if (accion === 'admin')         return leerTodosLosProspectos(sheet);
  if (accion === 'miscasos')      return leerMisCasos(e.parameter.tel, sheet);
  return leerProspectosPortal(sheet);
}


function doPost(e) {
  var sheet    = SpreadsheetApp.openById(SPREADSHEET_ID()).getSheetByName(SHEET_NAME);
  var postData = e.postData ? e.postData.contents : '';
  var params   = {};
  Logger.log('=== doPost recibido ===');
  Logger.log('Tipo: ' + (e.postData ? e.postData.type : 'sin postData'));
  Logger.log('Tamano body: ' + postData.length + ' chars (~' + Math.round(postData.length / 1024) + ' KB)');
  if (e.parameter && (e.parameter.formID || e.parameter.submissionID)) {
    return recibirWebhookJotForm(e.parameter, e.parameters || {});
  }
  try {
    params = JSON.parse(postData);
    Logger.log('JSON parseado OK. Claves: ' + Object.keys(params).join(', '));
    if (params.formID || params.submissionID)  return recibirWebhookJotForm(params);
    if (params.accion === 'registrarSolicitud') return accionRegistrarSolicitud(params, sheet);
    if (params.accion === 'solicitarGarantia')  return accionSolicitarGarantia(params, sheet);
    return manejarPostPortal(params);
  } catch(err) {
    Logger.log('Error parseando JSON: ' + err.message);
    Logger.log('Primeros 500 chars del body: ' + postData.substring(0, 500));
  }
  return jsonOk({ ok: false, error: 'Solicitud POST no reconocida.' });
}


function accionCambiarEstado(params, sheet) {
  var id     = params.id;
  var estado = params.estado;
  var data   = sheet.getRange('A1:AZ2000').getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][COL_ID - 1]).trim() === String(id).trim()) {
      var row = i + 1;
      sheet.getRange(row, COL_ESTADO).setValue(estado);
      sheet.getRange(row, COL_ESTADO_PORTAL).setValue(estado);
      if (estado === 'Asignado') {
        sheet.getRange(row, COL_FECHA_ASIG).setValue(new Date());
        sheet.getRange(row, COL_PAGO).setValue('Validado');
      }
      if (['Disponible', 'Sabado', 'Regalado', 'Contactado'].indexOf(estado) !== -1) {
        sheet.getRange(row, COL_FECHA_ASIG).clearContent();
        sheet.getRange(row, COL_ALUMNO).clearContent();
        sheet.getRange(row, COL_PAGO).clearContent();
        sheet.getRange(row, COL_REF_PAGO).clearContent();
        sheet.getRange(row, COL_COMPROBANTE).clearContent();
        sheet.getRange(row, COL_TEL_ALUMNO).clearContent();
      }
      return jsonOk({ ok: true });
    }
  }
  return jsonOk({ ok: false, error: 'ID no encontrado' });
}


function accionCrearVIP(params, sheet) {
  var lastRow = sheet.getLastRow() + 1;
  var newId   = 'P-' + (lastRow - 1);
  sheet.getRange(lastRow, COL_ESTADO).setValue('VIP');
  sheet.getRange(lastRow, COL_FECHA_REG).setValue(new Date());
  sheet.getRange(lastRow, COL_NOMBRE).setValue(params.nombre || '');
  sheet.getRange(lastRow, COL_TEL).setValue(params.telefono || '');
  sheet.getRange(lastRow, COL_TRATAMIENTO).setValue(params.motivo || '');
  sheet.getRange(lastRow, COL_ID).setValue(newId);
  sheet.getRange(lastRow, COL_ESTADO_PORTAL).setValue('VIP');
  return jsonOk({ ok: true, id: newId });
}


function accionRegistrarSolicitud(params, sheet) {
  var id        = params.id;
  var alumno    = params.alumno      || '';
  var telAlumno = params.telAlumno   || '';
  var ref       = params.ref         || '';
  var monto     = params.monto       || '';
  var compB64   = params.comprobante || '';
  Logger.log('=== accionRegistrarSolicitud ===');
  Logger.log('ID: ' + id + ' | Ref: ' + ref + ' | Alumno: ' + alumno);
  Logger.log('Comprobante recibido: ' + compB64.length + ' chars (~' + Math.round(compB64.length / 1024) + ' KB)');
  if (compB64.length < 100) Logger.log('ALERTA: comprobante vacio o muy corto');
  var data = sheet.getRange('A1:AZ2000').getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][COL_ID - 1]).trim() === String(id).trim()) {
      var row          = i + 1;
      var estadoActual = data[i][COL_ESTADO - 1];
      if (estadoActual !== 'Disponible' && estadoActual !== 'VIP' && estadoActual !== 'Sabado') {
        return jsonOk({ ok: false, error: 'Este prospecto ya no esta disponible' });
      }
      var urlComprobante = '';
      if (compB64 && compB64.length > 100) {
        urlComprobante = guardarComprobanteEnDrive(compB64, ref);
        Logger.log('URL comprobante: ' + (urlComprobante || '(vacia)'));
      }
      sheet.getRange(row, COL_ESTADO).setValue('Pago enviado');
      sheet.getRange(row, COL_ESTADO_PORTAL).setValue('Pago enviado');
      sheet.getRange(row, COL_ALUMNO).setValue(alumno);
      sheet.getRange(row, COL_TEL_ALUMNO).setValue(telAlumno);
      sheet.getRange(row, COL_FECHA_ASIG).setValue(new Date());
      sheet.getRange(row, COL_PAGO).setValue('Pendiente validacion');
      sheet.getRange(row, COL_REF_PAGO).setValue(ref);
      if (urlComprobante) {
        sheet.getRange(row, COL_COMPROBANTE).setValue(urlComprobante);
      } else if (compB64 && compB64.length > 100) {
        sheet.getRange(row, COL_COMPROBANTE).setValue('[BASE64-PENDIENTE] ' + compB64.substring(0, 100));
      }
      var tratamiento = data[i][COL_TRATAMIENTO - 1] || '';
      MailApp.sendEmail(EMAIL_ADMIN(), 'Pago pendiente: ' + ref,
        'Nuevo comprobante recibido!\n\nAlumno: ' + alumno + '\nTel: ' + telAlumno +
        '\nRef: ' + ref + '\nMonto: ' + monto + '\nTratamiento: ' + tratamiento +
        '\nComprobante: ' + (urlComprobante || '(no guardado)') +
        '\n\nEntra al panel admin para validar.');
      return jsonOk({ ok: true, ref: ref });
    }
  }
  return jsonOk({ ok: false, error: 'ID no encontrado' });
}


// ================================================================
// [NEW] LECTURA: Mis casos
// ================================================================
function leerMisCasos(tel, sheet) {
  if (!tel) return jsonOk({ ok: false, error: 'Sin telefono' });
  var data      = sheet.getRange('A1:AZ2000').getValues();
  var resultado = [];
  var telLimpio = String(tel).replace(/\D/g, '');
  Logger.log('leerMisCasos tel: ' + telLimpio);
  for (var i = 1; i < data.length; i++) {
    var telAlumno = String(data[i][COL_TEL_ALUMNO - 1] || '').replace(/\D/g, '');
    if (telAlumno && telAlumno.indexOf(telLimpio) !== -1) {
      resultado.push({
        id:              data[i][COL_ID - 1],
        tratamiento:     data[i][COL_TRATAMIENTO - 1],
        extracciones:    data[i][COL_EXTRACCIONES - 1],
        horario:         data[i][COL_HORARIO_MAN - 1],
        estado:          String(data[i][COL_ESTADO - 1] || ''),
        alumno:          data[i][COL_ALUMNO - 1],
        fechaAsignacion: data[i][COL_FECHA_ASIG - 1] ? Utilities.formatDate(
                           new Date(data[i][COL_FECHA_ASIG - 1]),
                           Session.getScriptTimeZone(), 'dd/MM/yyyy') : ''
      });
    }
  }
  Logger.log('leerMisCasos encontrados: ' + resultado.length);
  return jsonOk({ ok: true, data: resultado });
}


// ================================================================
// [NEW] ACCION: Solicitar garantia
// ================================================================
function accionSolicitarGarantia(params, sheet) {
  var id        = params.id;
  var alumno    = params.alumno    || '';
  var telAlumno = params.telAlumno || '';
  var evidencia = params.evidencia || '';
  Logger.log('=== accionSolicitarGarantia === ID: ' + id + ' | Alumno: ' + alumno);
  var data = sheet.getRange('A1:AZ2000').getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][COL_ID - 1]).trim() === String(id).trim()) {
      var row = i + 1;
      sheet.getRange(row, COL_ESTADO).setValue('Garantia');
      sheet.getRange(row, COL_ESTADO_PORTAL).setValue('Garantia');
      var urlEvidencia = '';
      if (evidencia && evidencia.length > 100) {
        urlEvidencia = guardarComprobanteEnDrive(evidencia, 'garantia_' + id + '_' + Date.now());
        Logger.log('Evidencia guardada: ' + (urlEvidencia || '(error)'));
      }
      var tratamiento = data[i][COL_TRATAMIENTO - 1] || '';
      MailApp.sendEmail(EMAIL_ADMIN(), 'Garantia solicitada: ' + id,
        'El alumno reporta que el prospecto no respondio.\n\n' +
        'Alumno: ' + alumno + '\nTel: ' + telAlumno +
        '\nProspecto: ' + id + ' - ' + tratamiento +
        '\nEvidencia: ' + (urlEvidencia || '(sin archivo)') +
        '\n\nValida y reasigna desde el panel admin.');
      return jsonOk({ ok: true, id: id });
    }
  }
  return jsonOk({ ok: false, error: 'ID no encontrado' });
}


function leerProspectosPortal(sheet) {
  var data       = sheet.getRange('A1:AZ2000').getValues();
  var prospectos = [];
  for (var i = 1; i < data.length; i++) {
    var estado = data[i][COL_ESTADO - 1];
    if (estado !== 'Disponible' && estado !== 'VIP' && estado !== 'Sabado') continue;
    if (!data[i][COL_ID - 1]) continue;
    prospectos.push({
      id:           data[i][COL_ID - 1],
      tratamiento:  data[i][COL_TRATAMIENTO - 1],
      extracciones: data[i][COL_EXTRACCIONES - 1],
      expediente:   data[i][COL_EXPEDIENTE - 1],
      edad:         data[i][COL_EDAD - 1],
      diente:       data[i][COL_DIENTE - 1],
      resinas:      data[i][COL_RESINAS - 1],
      aparato:      String(data[i][COL_APARATO - 1] || '').toLowerCase().indexOf('si') === 0,
      horario:      data[i][COL_HORARIO_MAN - 1],
      horarioTarde: data[i][COL_HORARIO_TAR - 1],
      dolor:        data[i][COL_DOLOR - 1],
      fractura:     data[i][COL_FRACTURA - 1],
      estado:       estado,
      foto:         data[i][COL_FOTO - 1],
    });
  }
  return jsonOk({ ok: true, data: prospectos, precios: getDefaultPrecios() });
}


function leerTodosLosProspectos(sheet) {
  var data          = sheet.getRange('A1:AZ2000').getValues();
  var prospectos    = [];
  var estadosValidos = ['Disponible','Asignado','VIP','Garantia','Contactado',
                        'No respondio','Regalado','Sabado','Pago enviado'];
  for (var i = 1; i < data.length; i++) {
    var estado = data[i][COL_ESTADO - 1];
    if (!estado || estadosValidos.indexOf(String(estado)) === -1) continue;
    if (!data[i][COL_ID - 1]) continue;
    prospectos.push({
      id:              data[i][COL_ID - 1],
      nombre:          data[i][COL_NOMBRE - 1],
      telefono:        data[i][COL_TEL - 1],
      tratamiento:     data[i][COL_TRATAMIENTO - 1],
      extracciones:    data[i][COL_EXTRACCIONES - 1],
      edad:            data[i][COL_EDAD - 1],
      diente:          data[i][COL_DIENTE - 1],
      resinas:         data[i][COL_RESINAS - 1],
      aparato:         String(data[i][COL_APARATO - 1] || '').toLowerCase().indexOf('si') === 0,
      expediente:      data[i][COL_EXPEDIENTE - 1],
      horario:         data[i][COL_HORARIO_MAN - 1],
      horarioTarde:    data[i][COL_HORARIO_TAR - 1],
      dolor:           data[i][COL_DOLOR - 1],
      fractura:        data[i][COL_FRACTURA - 1],
      estado:          estado,
      foto:            data[i][COL_FOTO - 1],
      alumnoAsignado:  data[i][COL_ALUMNO - 1],
      telAlumno:       data[i][COL_TEL_ALUMNO - 1],
      fechaAsignacion: data[i][COL_FECHA_ASIG - 1] ? Utilities.formatDate(
                         new Date(data[i][COL_FECHA_ASIG - 1]),
                         Session.getScriptTimeZone(), 'dd/MM/yyyy') : '',
      pago:            data[i][COL_PAGO - 1],
      refPago:         data[i][COL_REF_PAGO - 1],
      comprobante:     data[i][COL_COMPROBANTE - 1],
    });
  }
  return jsonOk({ ok: true, data: prospectos, precios: getDefaultPrecios() });
}


function guardarComprobanteEnDrive(base64String, ref) {
  try {
    Logger.log('guardarComprobanteEnDrive: ' + base64String.length + ' chars');
    var partes = base64String.split(',');
    if (partes.length !== 2) { Logger.log('Error: base64 sin encabezado'); return ''; }
    var mimeMatch = partes[0].match(/:(.*?);/);
    var mimeType  = mimeMatch ? mimeMatch[1] : 'image/jpeg';
    var datos  = Utilities.base64Decode(partes[1]);
    var blob   = Utilities.newBlob(datos, mimeType, ref + '.jpg');
    var folder = DriveApp.getFolderById(FOLDER_COMPROBANTES());
    var file   = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var url = 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w1000';
    Logger.log('Archivo guardado: ' + url);
    return url;
  } catch (err) {
    Logger.log('Error en guardarComprobanteEnDrive: ' + err.message);
    return '';
  }
}


function subirFotoADrive(jotformUrl, nombrePaciente) {
  Logger.log('subirFotoADrive URL: ' + jotformUrl);
  try {
    var folderFotosId = FOLDER_FOTOS();
    if (!folderFotosId || folderFotosId === 'PEGA_AQUI_EL_ID_DE_TU_CARPETA_FOTOS') return jotformUrl;
    var response = UrlFetchApp.fetch(jotformUrl, {
      muteHttpExceptions: true, followRedirects: true,
      headers: { 'User-Agent': 'Mozilla/5.0 (compatible; GoogleAppsScript)', 'Accept': 'image/*, */*' }
    });
    var statusCode = response.getResponseCode();
    Logger.log('HTTP status Jotform: ' + statusCode);
    if (statusCode !== 200) { Logger.log('Jotform devolvio ' + statusCode); return jotformUrl; }
    var blob        = response.getBlob();
    var contentType = blob.getContentType();
    Logger.log('Content-Type: ' + contentType + ' | Tamano: ' + blob.getBytes().length + ' bytes');
    if (contentType && contentType.indexOf('image/') === -1 && contentType.indexOf('application/octet') === -1) {
      Logger.log('Lo descargado no es imagen: ' + contentType); return jotformUrl;
    }
    var ext = jotformUrl.split('?')[0].split('.').pop().toLowerCase();
    ext = ['jpg','jpeg','png','gif','webp'].indexOf(ext) > -1 ? ext : 'jpg';
    blob.setName((nombrePaciente || 'paciente').replace(/\s+/g,'_') + '_' + Date.now() + '.' + ext);
    var folder = DriveApp.getFolderById(folderFotosId);
    var file   = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var url = 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w1000';
    Logger.log('Foto guardada: ' + url);
    return url;
  } catch(err) {
    Logger.log('Excepcion en subirFotoADrive: ' + err.message); return jotformUrl;
  }
}


function manejarPostPortal(params) {
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID()).getSheetByName(SHEET_NAME);
  var data  = sheet.getDataRange().getValues();
  if (params.accion === 'crearVIP') return accionCrearVIP(params, sheet);
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][COL_ID - 1]).trim() === String(params.id).trim()) {
      if (params.estado) sheet.getRange(i + 1, COL_ESTADO).setValue(params.estado);
      if (params.alumno) sheet.getRange(i + 1, COL_ALUMNO).setValue(params.alumno);
      if (params.pago)   sheet.getRange(i + 1, COL_PAGO).setValue(params.pago);
      return jsonOk({ ok: true });
    }
  }
  return jsonOk({ ok: false, error: 'ID no encontrado' });
}


function recibirWebhookJotForm(p, pArr) {
  pArr = pArr || {};
  try {
    var sheet  = SpreadsheetApp.openById(SPREADSHEET_ID()).getSheetByName(SHEET_NAME);
    var pretty = (p.pretty || '').replace(/<[^>]+>/g, '');
    function fromPretty(qRegex) {
      var searchFrom = 0;
      while (searchFrom < pretty.length) {
        var sub   = pretty.slice(searchFrom);
        var match = sub.search(qRegex);
        if (match === -1) return '';
        var absIdx     = searchFrom + match;
        var prevComma  = pretty.lastIndexOf(',', absIdx - 1);
        var textBefore = pretty.substring(prevComma === -1 ? 0 : prevComma + 1, absIdx);
        if (textBefore.indexOf(':') !== -1) { searchFrom = absIdx + 1; continue; }
        var colonIdx = pretty.indexOf(':', absIdx);
        if (colonIdx === -1) return '';
        var rest  = pretty.substring(colonIdx + 1).trim();
        var nextQ = rest.search(/,\s*[A-ZÁÉÍÓÚÑa-záéíóúñ¿]/);
        if (nextQ === -1) return rest.trim().replace(/\.$/, '');
        return rest.substring(0, nextQ).trim().replace(/\.$/, '');
      }
      return '';
    }
    var nombre      = fromPretty(/Nombre completo/i);
    var telefono    = fromPretty(/WhatsApp|[Tt]el[eé]fono|[Cc]elular/i);
    var expediente  = fromPretty(/expediente/i);
    var tratRaw = fromPretty(/tratamiento|[Ss]elecciona el trat/i);
    if (!tratRaw) tratRaw = fromPretty(/motivo.*consulta|consulta/i);
    var tratamientoIconUrl = '';
    if (tratRaw.indexOf('|') > -1) {
      var afterPipe = tratRaw.split('|').slice(1).join('|').trim();
      var urlClean  = afterPipe.match(/^(https?:\/\/[^\s,<"]+)/);
      tratamientoIconUrl = urlClean ? urlClean[1] : '';
    }
    var tratamiento    = tratRaw.split('|')[0].trim();
    var extracciones   = fromPretty(/cu[aá]ntas extracciones|para informar.*extracci/i);
    if (!extracciones) extracciones = fromPretty(/extracci[oó]n/i);
    var horario        = fromPretty(/horario|acudir|citas/i);
    var edad           = fromPretty(/edad.*paciente|selecciona.*edad|edad\?/i);
    var diente         = fromPretty(/cu[aá]l diente|en qu[eé] diente|diente afectado/i);
    var resinas        = fromPretty(/cu[aá]ntas resinas|cu[aá]ntas caries/i);
    var aparato        = fromPretty(/aparato.*ortop|ortop|arco lingual|mantenedor/i);
    var consentimiento = fromPretty(/compartir.*informaci[oó]n|acuerdo.*compartir|acuerdo.*informaci[oó]n personal/i);
    var extProtesis    = fromPretty(/placa.*dentadura|dentadura.*sacar|ponerte.*placa|sacar.*dientes.*placa/i);
    var dolor          = fromPretty(/dolor.*fr[ií]a|fr[ií]a.*caliente|sientes.*dolor.*agua/i);
    var fractura       = fromPretty(/qu[eé] tan grande.*fractura|grande.*fractura/i);
    if (telefono && String(telefono).replace(/\D/g, '').length < 8) telefono = '';
    var META  = /^(formID|submissionID|formTitle|username|webhookURL|ip|type|rawRequest|pretty|action|appID|teamID|unread|isSilent|subject|parent|product|fromTable|event|documentID|customTitle|customParams|customBody|slug|uploadServerUrl)$/i;
    var fotos = [];
    function esUrlFoto(url) {
      if (!url || url.indexOf('http') !== 0) return false;
      if (!(url.indexOf('jotform.com') > -1 || url.indexOf('jotformz.com') > -1 || url.indexOf('jotform.io') > -1)) return false;
      if (tratamientoIconUrl) {
        var uShort = url.length < tratamientoIconUrl.length ? url : tratamientoIconUrl;
        var uLong  = url.length < tratamientoIconUrl.length ? tratamientoIconUrl : url;
        if (uLong.indexOf(uShort) === 0) return false;
      }
      return true;
    }
    for (var k in pArr) {
      if (META.test(k)) continue;
      var arr = pArr[k];
      if (!Array.isArray(arr)) continue;
      arr.forEach(function(v) { var url = String(v||'').trim(); if (esUrlFoto(url) && fotos.indexOf(url)===-1) fotos.push(url); });
    }
    for (var k in p) {
      if (META.test(k)) continue;
      var urlVal = String(p[k]||'').trim();
      if (esUrlFoto(urlVal) && fotos.indexOf(urlVal)===-1) fotos.push(urlVal);
    }
    var JOTFORM_UPLOAD_RE = /https?:\/\/(?:www\.)?jotform(?:z)?\.(?:com|io)\/uploads\/[^\s,|"<\]]+/gi;
    try {
      if (p.rawRequest) {
        var raw = JSON.parse(p.rawRequest);
        JSON.stringify(raw).replace(JOTFORM_UPLOAD_RE, function(url) { if (esUrlFoto(url) && fotos.indexOf(url)===-1) fotos.push(url); });
      }
    } catch(fe) {}
    JOTFORM_UPLOAD_RE.lastIndex = 0;
    pretty.replace(JOTFORM_UPLOAD_RE, function(url) { if (esUrlFoto(url) && fotos.indexOf(url)===-1) fotos.push(url); });
    pretty.replace(/\buploads\/fowuanl\/[^\s,|"<\]]+/gi, function(relPath) {
      var fullUrl = 'https://www.jotform.com/' + relPath;
      if (fotos.indexOf(fullUrl) === -1) {
        var uShort2 = fullUrl.length < tratamientoIconUrl.length ? fullUrl : tratamientoIconUrl;
        var uLong2  = fullUrl.length < tratamientoIconUrl.length ? tratamientoIconUrl : fullUrl;
        if (!tratamientoIconUrl || uLong2.indexOf(uShort2) !== 0) fotos.push(fullUrl);
      }
    });
    Logger.log('URLs foto en webhook: ' + fotos.length);
    fotos.forEach(function(u,idx){ Logger.log('  Foto '+(idx+1)+': '+u); });
    var fotosPublicas = fotos.map(function(url){ return subirFotoADrive(url, nombre); });
    var fotoUrl = fotosPublicas.filter(function(u){ return u && u.length > 0; }).join(' | ');
    var lastRow = sheet.getLastRow();
    var newRow  = lastRow + 1;
    var newId   = 'P-' + lastRow;
    sheet.getRange(newRow, COL_ESTADO).setValue('Revision');
    sheet.getRange(newRow, COL_FECHA_REG).setValue(new Date());
    sheet.getRange(newRow, COL_NOMBRE).setValue(nombre);
    sheet.getRange(newRow, COL_TEL).setValue(telefono);
    sheet.getRange(newRow, COL_EXPEDIENTE).setValue(expediente);
    sheet.getRange(newRow, COL_TRATAMIENTO).setValue(tratamiento);
    if (extracciones)   sheet.getRange(newRow, COL_EXTRACCIONES).setValue(extracciones);
    if (diente)         sheet.getRange(newRow, COL_DIENTE).setValue(diente);
    if (resinas)        sheet.getRange(newRow, COL_RESINAS).setValue(resinas);
    if (edad)           sheet.getRange(newRow, COL_EDAD).setValue(edad);
    if (aparato)        sheet.getRange(newRow, COL_APARATO).setValue(aparato);
    sheet.getRange(newRow, COL_HORARIO_MAN).setValue(horario || '');
    if (consentimiento) sheet.getRange(newRow, COL_CONSENTIMIENTO).setValue(consentimiento);
    if (extProtesis)    sheet.getRange(newRow, COL_EXT_PROTESIS).setValue(extProtesis);
    if (dolor)          sheet.getRange(newRow, COL_DOLOR).setValue(dolor);
    if (fractura)       sheet.getRange(newRow, COL_FRACTURA).setValue(fractura);
    sheet.getRange(newRow, COL_ID).setValue(newId);
    sheet.getRange(newRow, COL_ESTADO_PORTAL).setValue('Revision');
    if (fotoUrl) sheet.getRange(newRow, COL_FOTO).setValue(fotoUrl);
    Logger.log('Guardado: ' + newId + ' | ' + nombre + ' | foto: ' + (fotoUrl || 'sin foto'));
    return jsonOk({ ok: true, id: newId });
  } catch(err) {
    Logger.log('ERROR webhook JotForm: ' + err.message);
    return jsonOk({ ok: false, error: err.message });
  }
}

function jsonOk(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function accionGetPrecios() { return jsonOk({ ok: true, precios: getDefaultPrecios() }); }

function getDefaultPrecios() {
  var props = PropertiesService.getScriptProperties();
  return {
    vip:             parseInt(props.getProperty('precio_vip')             || '1500'),
    infantil_manana: parseInt(props.getProperty('precio_infantil_manana') || '200'),
    infantil_tarde:  parseInt(props.getProperty('precio_infantil_tarde')  || '130'),
    extraccion_2:    parseInt(props.getProperty('precio_extraccion_2')    || '260'),
    extraccion_3:    parseInt(props.getProperty('precio_extraccion_3')    || '390'),
    aparato:         parseInt(props.getProperty('precio_aparato')         || '300'),
    aparato_manana:  parseInt(props.getProperty('precio_aparato_manana')  || '600'),
    general:         parseInt(props.getProperty('precio_general')         || '130'),
    blanqueamiento:  parseInt(props.getProperty('precio_blanqueamiento')  || '1500'),
    mantenedor:      parseInt(props.getProperty('precio_mantenedor')      || '320'),
    corona_acero:    parseInt(props.getProperty('precio_corona_acero')    || '285')
  };
}

function accionSetPrecios(params) {
  var props  = PropertiesService.getScriptProperties();
  var campos = ['precio_vip','precio_infantil_manana','precio_infantil_tarde',
                'precio_extraccion_2','precio_extraccion_3','precio_aparato',
                'precio_aparato_manana','precio_general','precio_blanqueamiento',
                'precio_mantenedor','precio_corona_acero'];
  campos.forEach(function(c) {
    if (params[c] && !isNaN(parseInt(params[c]))) props.setProperty(c, String(parseInt(params[c])));
  });
  return jsonOk({ ok: true, precios: getDefaultPrecios() });
}
// ================================================================
// MENÚ DE BOTONES Y FILTROS — Columna A
// ================================================================

function onOpen() {
    var ui = SpreadsheetApp.getUi();
      ui.createMenu('📋 Prospectos')
          .addSubMenu(ui.createMenu('🔘 Cambiar Estado')
                .addItem('✅ Disponible', 'marcarDisponible')
                      .addItem('⭐ VIP', 'marcarVIP')
                            .addItem('👤 Asignado', 'marcarAsignado')
                                  .addItem('🔄 Garantía', 'marcarGarantia')
                                        .addItem('❌ No respondió', 'marcarNoRespondio'))
                                            .addSeparator()
                                                .addSubMenu(ui.createMenu('🔍 Filtros Rápidos')
                                                      .addItem('Ver solo Disponibles', 'filtrarDisponibles')
                                                            .addItem('Ver solo VIP', 'filtrarVIP')
                                                                        .addItem('Ver todos', 'quitarFiltros'))
                                                                            .addToUi();
}

function cambiarEstadoSeleccionados(nuevoEstado) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  var range = sheet.getActiveRange();
  if (!range) { SpreadsheetApp.getUi().alert('Selecciona una celda primero.'); return; }
  var filas = range.getNumRows();
  var filaInicio = range.getRow();
  var cambios = 0;
  for (var i = 0; i < filas; i++) {
    var fila = filaInicio + i;
    if (fila < 2) continue;
    sheet.getRange(fila, COL_ESTADO).setValue(nuevoEstado);
    cambios++;
    }
    aplicarColorEstado(sheet);
    SpreadsheetApp.getUi().alert(cambios + ' fila(s) marcadas como: ' + nuevoEstado);
}

function marcarDisponible() { cambiarEstadoSeleccionados('Disponible'); }
function marcarVIP() { cambiarEstadoSeleccionados('VIP'); }
function marcarAsignado() { cambiarEstadoSeleccionados('Asignado'); }
function marcarGarantia() { cambiarEstadoSeleccionados('Garantia'); }
function marcarNoRespondio() { cambiarEstadoSeleccionados('No respondio'); }

function aplicarColorEstado(sheet) {
  if (!sheet) { var ss = SpreadsheetApp.getActiveSpreadsheet(); sheet = ss.getSheetByName(SHEET_NAME); }
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  var estados = sheet.getRange(2, COL_ESTADO, lastRow - 1, 1).getValues();
  var colores = { 'Disponible': '#C8E6C9', 'VIP': '#FFF9C4', 'Asignado': '#B3E5FC', 'Garantia': '#F8BBD0', 'No respondio': '#FFCCBC' };
  var lastCol = sheet.getLastColumn();
  for (var i = 0; i < estados.length; i++) {
    var fila = i + 2;
    var color = colores[estados[i][0]] || '#FFFFFF';
    sheet.getRange(fila, 1, 1, lastCol).setBackground(color);
    }
}

function filtrarDisponibles() { filtrarPorEstado('Disponible'); }
function filtrarVIP() { filtrarPorEstado('VIP'); }

function filtrarPorEstado(estado) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  var range = sheet.getDataRange();
  var filter = range.getFilter();
  if (!filter) filter = range.createFilter();
  var criterio = SpreadsheetApp.newFilterCriteria().whenTextEqualTo(estado).build();
  filter.setColumnFilterCriteria(COL_ESTADO, criterio);
  SpreadsheetApp.getUi().alert('Mostrando solo: ' + estado);
}

function quitarFiltros() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  var filter = sheet.getDataRange().getFilter();
  if (filter) { filter.remove(); SpreadsheetApp.getUi().alert('Filtros eliminados. Mostrando todos los prospectos.'); }
  else { SpreadsheetApp.getUi().alert('No hay filtros activos.'); }
}
