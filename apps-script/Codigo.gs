// Pega este código en Apps Script (Extensiones → Apps Script en tu Google Sheet).
// Luego: Implementar → Nueva implementación → Aplicación web
//   Ejecutar como: Yo · Quién tiene acceso: Cualquier usuario
// Copia la URL terminada en /exec y pégala en CONFIG.googleSheet.scriptURL
// dentro de index.html.

const NOMBRE_HOJA = 'Confirmaciones'; // nombre de la pestaña donde se guardan las filas

function doPost(e) {
  const hoja = obtenerHoja();
  const p = e.parameter;

  hoja.appendRow([
    new Date(),
    p.titulo || '',
    p.invitado || '',
    p.asistencia || '',
    p.personas || '',
    p.mensaje || ''
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  return ContentService
    .createTextOutput('El endpoint funciona. Usa POST para enviar confirmaciones.')
    .setMimeType(ContentService.MimeType.TEXT);
}

function obtenerHoja() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = libro.getSheetByName(NOMBRE_HOJA);
  if (!hoja) {
    hoja = libro.insertSheet(NOMBRE_HOJA);
    hoja.appendRow(['Fecha', 'Título', 'Invitado', 'Asistencia', 'Personas', 'Mensaje']);
  }
  return hoja;
}
