/**
 * Google Apps Script — Dashboard Principalidad
 *
 * INSTRUCCIONES DE DEPLOY:
 * 1. Abrí script.google.com con tu cuenta MeLi
 * 2. Nuevo proyecto > pegá este código
 * 3. Guardar
 * 4. Implementar > Nueva implementación > Tipo: Aplicación web
 *    - Ejecutar como: Yo (tu cuenta)
 *    - Quién puede acceder: Todos (o "Personas de tu organización" si preferís)
 * 5. Autorizá los permisos cuando te los pida
 * 6. Copiá la URL que te da y pegala en index.html donde dice:
 *    const GAS_URL = ''; // Pegá acá la URL de tu Apps Script web app
 *
 * IMPORTANTE: Para que funcione fuera de VPN, elegí "Todos" en el paso 4.
 * Los datos de las planillas solo son accesibles porque el script corre
 * bajo TU cuenta de Google (que tiene acceso).
 */

// IDs de las planillas
const BASE_MADRE_ID = '1ymNdH1UnJ2VQZEiHJuHCrQbu2JT180_VVIlzbRXXXHQ';
const BASE_MADRE_TAB = 'base de mes actual';

const REPAGOS_ID = '11xFxl_XYFIhLGmokYJM9HpBUu53uRoUhWNqdqfHVPs8';
const REPAGOS_TAB = 'lookerv2';

// Máximo de filas a leer (para no exceder límites de Apps Script)
const MAX_ROWS = 10000;

function doGet(e) {
  try {
    const baseMadre = readSheet(BASE_MADRE_ID, BASE_MADRE_TAB);
    const lookerv2 = readSheet(REPAGOS_ID, REPAGOS_TAB);

    const output = {
      timestamp: new Date().toISOString(),
      baseMadre: baseMadre,
      lookerv2: lookerv2
    };

    const json = JSON.stringify(output);

    // Soporte JSONP: si se pasa ?callback=nombreFuncion, retorna JS en vez de JSON
    // Esto permite cargar los datos desde un <script> tag, evitando restricciones CORS
    const callback = e && e.parameter && e.parameter.callback;
    if (callback) {
      return ContentService
        .createTextOutput(callback + '(' + json + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }

    return ContentService
      .createTextOutput(json)
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    const errorOutput = {
      error: err.message,
      timestamp: new Date().toISOString()
    };
    const json = JSON.stringify(errorOutput);
    const callback = e && e.parameter && e.parameter.callback;
    if (callback) {
      return ContentService
        .createTextOutput(callback + '(' + json + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService
      .createTextOutput(json)
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Lee una hoja y devuelve un array de arrays (primera fila = headers)
 * Limita a MAX_ROWS para no exceder cuotas
 */
function readSheet(spreadsheetId, sheetName) {
  try {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      Logger.log('No se encontró la hoja: ' + sheetName + ' en ' + spreadsheetId);
      return [];
    }
    const lastRow = Math.min(sheet.getLastRow(), MAX_ROWS + 1);
    const lastCol = sheet.getLastColumn();
    if (lastRow < 1 || lastCol < 1) return [];
    const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    return data;
  } catch (err) {
    Logger.log('Error leyendo hoja ' + sheetName + ': ' + err.message);
    return [];
  }
}

/**
 * Función de test — ejecutala manualmente desde el editor para verificar
 * que las planillas se leen correctamente
 */
function testRead() {
  const bm = readSheet(BASE_MADRE_ID, BASE_MADRE_TAB);
  Logger.log('Base Madre — filas: ' + bm.length);
  if (bm.length > 0) Logger.log('Headers: ' + bm[0].join(' | '));
  if (bm.length > 1) Logger.log('Primera fila: ' + bm[1].join(' | '));

  const rep = readSheet(REPAGOS_ID, REPAGOS_TAB);
  Logger.log('Looker v2 — filas: ' + rep.length);
  if (rep.length > 0) Logger.log('Headers: ' + rep[0].join(' | '));
  if (rep.length > 1) Logger.log('Primera fila: ' + rep[1].join(' | '));
}
