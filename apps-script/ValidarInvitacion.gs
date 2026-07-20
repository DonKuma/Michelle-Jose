// ⚠️ IMPORTANTE: este código va en un proyecto de Apps Script DISTINTO al
// de Codigo.gs (el de las confirmaciones). NO lo pegues como un archivo
// nuevo dentro de ESE MISMO proyecto — si los mezclas, ambos archivos
// declaran cosas con el mismo nombre (NOMBRE_HOJA, obtenerHoja, doGet...) y
// Apps Script se rompe por completo (incluidas las confirmaciones), con un
// error como: "Identifier 'NOMBRE_HOJA' has already been declared".
//
// CÓMO CONFIGURARLO (una sola vez):
//  1. Abre el Google Sheet de la LISTA DE INVITACIONES generadas (el de
//     https://docs.google.com/spreadsheets/d/1Tzul34pFXPegBFhKQ5mZlM2383bgpAQp/edit)
//     — NO el de confirmaciones.
//  2. En ESE sheet: menú "Extensiones" → "Apps Script". Esto abre (o crea)
//     un proyecto propio de Apps Script, separado del de confirmaciones.
//  3. Borra TODO lo que haya en Code.gs (incluido el "function myFunction(){}"
//     de ejemplo) y pega este archivo completo. Guarda.
//  4. Si la URL generada no está en la columna F de tu sheet, ajusta
//     VI_COLUMNA_URL. (No hace falta indicar pestaña/gid: el script revisa
//     TODAS las pestañas del spreadsheet, así que funciona sin importar
//     cómo se llamen o en cuál estén.)
//  5. Botón "Implementar" → "Nueva implementación" → tipo "Aplicación web".
//     - Ejecutar como: Yo (tu cuenta)
//     - Quién tiene acceso: Cualquier usuario
//     Pulsa "Implementar" y autoriza los permisos que pida Google.
//  6. Copia la URL que termina en "/exec" (será una URL nueva y distinta a
//     la del script de confirmaciones) y pégala en
//     CONFIG.validarInvitacion.scriptURL dentro de index.html.
//  7. Si vuelves a editar este código, crea una "Nueva implementación" (o
//     "Gestionar implementaciones" → editar → nueva versión) para que los
//     cambios se apliquen.
//
// Para comprobar que funciona sin depender del sitio web: pega en el
// navegador tu URL "/exec?check=" + una URL de prueba codificada, por
// ejemplo:
//   TU_URL/exec?check=https%3A%2F%2Fdonkuma.github.io%2FMichelle-Jose%2F%3Ftitulo%3DProba%26invitado%3DProbando
// y deberías ver {"valid":false} (porque esa no está en tu lista) sin
// ningún mensaje de "Error" de Google Apps Script.

// Columna (1 = A, 2 = B, ...) donde está la URL completa generada para cada
// invitado. En "Generador de invitaciones.xlsx" esa es la columna F
// ("URL generada"), por eso el valor por defecto es 6. No se busca por el
// texto del encabezado (puede variar) sino por posición.
const VI_COLUMNA_URL = 6;

function doGet(e) {
  const entrante = (e.parameter && e.parameter.check) || '';
  let valido = false;
  try {
    valido = entrante ? viValidarUrl(entrante) : false;
  } catch (err) {
    // Si algo falla (columna mal configurada, permisos, etc.) se devuelve
    // el error en vez de una excepción sin formato, para poder detectarlo
    // abriendo esta misma URL "/exec?check=..." directo en el navegador.
    return ContentService
      .createTextOutput(JSON.stringify({ valid: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService
    .createTextOutput(JSON.stringify({ valid: valido }))
    .setMimeType(ContentService.MimeType.JSON);
}

function viValidarUrl(urlEntrante) {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  const hojas = libro.getSheets();
  const claveEntrante = viNormalizarQuery(urlEntrante);

  for (let h = 0; h < hojas.length; h++) {
    const hoja = hojas[h];
    const ultimaFila = hoja.getLastRow();
    if (ultimaFila < 1) continue;

    const columna = hoja.getRange(1, VI_COLUMNA_URL, ultimaFila, 1).getValues();
    for (let i = 0; i < columna.length; i++) {
      const urlFila = String(columna[i][0] || '').trim();
      // Salta la fila de encabezado ("URL generada") y cualquier celda que
      // no sea realmente una URL (por ejemplo la nota "URL base del sitio:").
      if (urlFila.toLowerCase().indexOf('http') !== 0) continue;
      if (viNormalizarQuery(urlFila) === claveEntrante) return true;
    }
  }
  return false;
}

// Compara solo los parámetros de la URL (?titulo=...&invitado=...), sin
// importar el dominio/ruta ni el orden en que vengan, para que pequeñas
// diferencias de formato no rompan la validación.
function viNormalizarQuery(url) {
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
