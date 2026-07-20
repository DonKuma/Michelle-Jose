// Pega este código en el Apps Script DEL SHEET donde llevas la lista de
// invitaciones generadas (el que tiene la columna con la URL de cada invitado).
// Es un proyecto de Apps Script SEPARADO del de Codigo.gs (ese es para las
// confirmaciones de asistencia; este es solo para validar que la URL que
// alguien abre esté en tu lista).
//
// CÓMO CONFIGURARLO (una sola vez):
//  1. Abre el Google Sheet con la lista de invitaciones generadas.
//  2. Menú "Extensiones" → "Apps Script".
//  3. Borra el contenido de Code.gs y pega TODO este archivo. Guarda.
//  4. Ajusta abajo NOMBRE_HOJA (o GID_HOJA) y COLUMNA_URL para que coincidan
//     exactamente con tu sheet.
//  5. Botón "Implementar" → "Nueva implementación" → tipo "Aplicación web".
//     - Ejecutar como: Yo (tu cuenta)
//     - Quién tiene acceso: Cualquier usuario
//     Pulsa "Implementar" y autoriza los permisos que pida Google.
//  6. Copia la URL que termina en "/exec" y pégala en CONFIG.validarInvitacion.scriptURL
//     dentro de index.html.
//  7. Si vuelves a editar este código, crea una "Nueva implementación" (o
//     "Gestionar implementaciones" → editar → nueva versión) para que los
//     cambios se apliquen.

// Deja NOMBRE_HOJA vacío ("") para ubicar la pestaña por GID_HOJA en su lugar
// (el número que ves en la URL del sheet después de "gid="). Si NOMBRE_HOJA
// tiene un valor, se usa ese nombre y se ignora GID_HOJA.
const NOMBRE_HOJA = '';
const GID_HOJA = 1035945475;

// Encabezado (fila 1) de la columna que contiene la URL completa generada
// para cada invitado. Debe coincidir tal cual con el texto de esa celda.
const COLUMNA_URL = 'URL';

function doGet(e) {
  const entrante = (e.parameter && e.parameter.check) || '';
  const valido = entrante ? validarUrl(entrante) : false;

  return ContentService
    .createTextOutput(JSON.stringify({ valid: valido }))
    .setMimeType(ContentService.MimeType.JSON);
}

function validarUrl(urlEntrante) {
  const hoja = obtenerHoja();
  const datos = hoja.getDataRange().getValues();
  if (datos.length < 2) return false;

  const encabezados = datos[0].map(h => String(h).trim().toLowerCase());
  const col = encabezados.indexOf(COLUMNA_URL.trim().toLowerCase());
  if (col === -1) {
    throw new Error('No se encontró la columna "' + COLUMNA_URL + '". Revisa COLUMNA_URL en el script.');
  }

  const claveEntrante = normalizarQuery(urlEntrante);
  for (let i = 1; i < datos.length; i++) {
    const urlFila = String(datos[i][col] || '').trim();
    if (!urlFila) continue;
    if (normalizarQuery(urlFila) === claveEntrante) return true;
  }
  return false;
}

function obtenerHoja() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  if (NOMBRE_HOJA) {
    const hoja = libro.getSheetByName(NOMBRE_HOJA);
    if (!hoja) throw new Error('No existe una pestaña llamada "' + NOMBRE_HOJA + '".');
    return hoja;
  }
  const hoja = libro.getSheets().find(h => h.getSheetId() === GID_HOJA);
  if (!hoja) throw new Error('No existe ninguna pestaña con gid=' + GID_HOJA + '.');
  return hoja;
}

// Compara solo los parámetros de la URL (?titulo=...&invitado=...), sin
// importar el dominio/ruta ni el orden en que vengan, para que pequeñas
// diferencias de formato no rompan la validación.
function normalizarQuery(url) {
  const sinHash = String(url).split('#')[0];
  const qIndex = sinHash.indexOf('?');
  const qs = qIndex === -1 ? '' : sinHash.slice(qIndex + 1);
  if (!qs) return '';

  const params = {};
  qs.split('&').forEach(par => {
    if (!par) return;
    const partes = par.split('=');
    const clave = decodeURIComponent((partes[0] || '').replace(/\+/g, ' ')).trim();
    const valor = decodeURIComponent((partes[1] || '').replace(/\+/g, ' ')).trim();
    if (clave) params[clave] = valor;
  });

  return Object.keys(params).sort()
    .map(k => k + '=' + params[k])
    .join('&');
}
