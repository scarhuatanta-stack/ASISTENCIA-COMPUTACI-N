const SHEET_NAME = 'Asistencia';

function doGet() {
  return ContentService.createTextOutput(JSON.stringify({ok:true,service:'Asistencia Computación'}))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents || '{}');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Fecha','Hora','Identificación','Dato QR','Origen']);
    }
    const now = new Date();
    const tz = Session.getScriptTimeZone() || 'America/Santiago';
    sheet.appendRow([
      Utilities.formatDate(now, tz, 'dd/MM/yyyy'),
      Utilities.formatDate(now, tz, 'HH:mm:ss'),
      data.identificacion || '',
      data.dato || '',
      data.origen || 'Web'
    ]);
    return ContentService.createTextOutput(JSON.stringify({ok:true,message:'Registro guardado'}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ok:false,error:String(err)}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
