// Pega este código en Apps Script (Extensiones → Apps Script en tu Google Sheet
// de CONFIRMACIONES:
// https://docs.google.com/spreadsheets/d/1e12zmwT71qSo7ft3jrFO5n8vGUZJ21rgeDQuys5pUKM/edit ).
// Luego: Implementar → Nueva implementación → Aplicación web
//   Ejecutar como: Yo · Quién tiene acceso: Cualquier usuario
// Copia la URL terminada en /exec y pégala en CONFIG.googleSheet.scriptURL
// dentro de index.html.
//
// ⚠️ Si ya tenías una versión anterior de este archivo desplegada, después de
// pegar este código debes crear una "Nueva implementación" (o "Gestionar
// implementaciones" → editar → versión nueva) para que el chequeo de
// "¿ya respondió?" (doGet) empiece a funcionar. Si conservas la MISMA URL
// /exec, no hay que tocar index.html.

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
    p.personasNombres || '',
    p.mensaje || ''
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

// GET con ?check=1&titulo=...&invitado=...  →  { respondido: true|false }
// Lo usa index.html al cargar la invitación: si esa invitación ya tiene una
// fila en la hoja, en vez del formulario se muestra el aviso de que la
// respuesta ya fue enviada.
//
// Para probarlo a mano, abre en el navegador:
//   TU_URL/exec?check=1&titulo=Familia&invitado=Perez
function doGet(e) {
  const p = (e && e.parameter) || {};

  if (!p.check) {
    return ContentService
      .createTextOutput('El endpoint funciona. Usa POST para enviar confirmaciones.')
      .setMimeType(ContentService.MimeType.TEXT);
  }

  try {
    const encontrada = buscarRespuesta(p.titulo || '', p.invitado || '');
    return ContentService
      .createTextOutput(JSON.stringify(encontrada
        ? { respondido: true, asistencia: encontrada.asistencia, fecha: encontrada.fecha }
        : { respondido: false }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    // Se devuelve el error como JSON (en vez de una excepción de Apps Script)
    // para poder diagnosticarlo abriendo la URL /exec?check=... en el navegador.
    return ContentService
      .createTextOutput(JSON.stringify({ respondido: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Devuelve la última fila que corresponde a esa invitación, o null.
function buscarRespuesta(titulo, invitado) {
  const claveInvitado = normalizar(invitado);
  if (!claveInvitado) return null; // sin invitado no hay nada que comparar

  const libro = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = libro.getSheetByName(NOMBRE_HOJA);
  if (!hoja) return null;

  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return null; // solo encabezado (o vacía)

  const datos = hoja.getRange(1, 1, ultimaFila, hoja.getLastColumn()).getValues();
  const cols = ubicarColumnas(datos[0]);
  const claveTitulo = normalizar(titulo);

  // Se recorre de abajo hacia arriba para quedarse con la respuesta más
  // reciente si por algún motivo hubiera más de una fila del mismo invitado.
  for (let i = datos.length - 1; i >= 1; i--) {
    const fila = datos[i];
    if (normalizar(fila[cols.invitado]) !== claveInvitado) continue;
    if (normalizar(fila[cols.titulo]) !== claveTitulo) continue;
    return {
      asistencia: String(fila[cols.asistencia] || ''),
      fecha: fila[cols.fecha] instanceof Date
        ? Utilities.formatDate(fila[cols.fecha], Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')
        : String(fila[cols.fecha] || '')
    };
  }
  return null;
}

// Ubica las columnas por el texto del encabezado; si no lo reconoce, usa las
// posiciones con las que doPost escribe las filas.
function ubicarColumnas(encabezado) {
  const indices = {};
  encabezado.forEach((celda, i) => {
    const clave = normalizar(celda);
    if (indices[clave] === undefined) indices[clave] = i;
  });
  const tomar = (nombre, porDefecto) =>
    indices[nombre] !== undefined ? indices[nombre] : porDefecto;

  return {
    fecha: tomar('fecha', 0),
    titulo: tomar('titulo', 1),
    invitado: tomar('invitado', 2),
    asistencia: tomar('asistencia', 3)
  };
}

// Minúsculas, sin tildes, sin espacios de más: así "Familia  Pérez" y
// "familia perez" cuentan como la misma invitación.
function normalizar(valor) {
  return String(valor === null || valor === undefined ? '' : valor)
    .normalize('NFD').replace(/\p{M}/gu, '')
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .trim();
}

function obtenerHoja() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = libro.getSheetByName(NOMBRE_HOJA);
  if (!hoja) {
    hoja = libro.insertSheet(NOMBRE_HOJA);
    hoja.appendRow(['Fecha', 'Título', 'Invitado', 'Asistencia', 'Personas', 'Nombres asistentes', 'Mensaje']);
  }
  return hoja;
}
