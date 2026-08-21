const HOJA_PERSONAS = 'PERSONAS';
const HOJA_REGISTROS = 'REGISTROS';

function doGet(e) {
  e = e || { parameter: {} };
  const p = e.parameter || {};

  // API JSONP para GitHub Pages.
  if (p.callback) {
    const callback = String(p.callback).replace(/[^a-zA-Z0-9_.$]/g, '');
    let resultado;
    try {
      if (p.accion === 'buscar') {
        resultado = buscarPersona(p.rut);
      } else if (p.accion === 'guardar') {
        resultado = guardarRegistro(p.rut, p.nombre, p.equipo);
      } else {
        resultado = { ok: false, mensaje: 'Acción no válida' };
      }
    } catch (err) {
      resultado = { ok: false, mensaje: err.message || String(err) };
    }
    return ContentService
      .createTextOutput(callback + '(' + JSON.stringify(resultado) + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('ASISTENCIA COMPUTACIÓN')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function normalizarRut(rut) {
  if (!rut) return '';
  return rut.toString().trim().toUpperCase()
    .replace(/\./g, '').replace(/-/g, '').replace(/\s/g, '');
}

function buscarPersona(rut) {
  const rutBuscado = normalizarRut(rut);
  if (!rutBuscado) return { encontrado: false, mensaje: 'Debe ingresar un RUT' };

  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_PERSONAS);
  if (!hoja) throw new Error('No existe la hoja PERSONAS');

  const datos = hoja.getDataRange().getValues();
  for (let i = 1; i < datos.length; i++) {
    if (normalizarRut(datos[i][0]) === rutBuscado) {
      return {
        encontrado: true,
        rut: datos[i][0].toString(),
        nombre: datos[i][1] ? datos[i][1].toString() : ''
      };
    }
  }
  return { encontrado: false, mensaje: 'RUT no encontrado' };
}

function guardarRegistro(rut, nombre, equipo) {
  const rutNormalizado = normalizarRut(rut);
  if (!rutNormalizado) throw new Error('El RUT es obligatorio');
  if (!nombre || nombre.toString().trim() === '') throw new Error('El nombre es obligatorio');

  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_REGISTROS);
  if (!hoja) throw new Error('No existe la hoja REGISTROS');

  const ahora = new Date();
  const zona = Session.getScriptTimeZone() || 'America/Santiago';
  const fecha = Utilities.formatDate(ahora, zona, 'dd/MM/yyyy');
  const hora = Utilities.formatDate(ahora, zona, 'HH:mm:ss');

  hoja.appendRow([fecha, hora, rutNormalizado, nombre.toString().trim(), equipo ? equipo.toString().trim() : '']);
  return { ok: true, mensaje: 'Asistencia registrada correctamente', fecha: fecha, hora: hora };
}

function probarConexion() {
  return { ok: true, archivo: SpreadsheetApp.getActiveSpreadsheet().getName(), mensaje: 'Conexión con Google Sheets funcionando correctamente' };
}
