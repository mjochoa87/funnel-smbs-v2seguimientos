/**
 * Google Apps Script — Dashboard Principalidad
 * Con soporte JSONP y CacheService (5 minutos de caché para requests subsiguientes)
 *
 * INSTRUCCIONES DE DEPLOY:
 * 1. Abrí script.google.com con tu cuenta MeLi
 * 2. Pegá este código en Código.gs
 * 3. Guardar (Ctrl+S)
 * 4. Implementar → Gestionar implementaciones → editar implementación existente
 *    → Versión nueva → Implementar
 */

const BASE_MADRE_ID = '1ymNdH1UnJ2VQZEiHJuHCrQbu2JT180_VVIlzbRXXXHQ';
const BASE_MADRE_TAB = 'base de mes actual';
const REPAGOS_ID = '11xFxl_XYFIhLGmokYJM9HpBUu53uRoUhWNqdqfHVPs8';
const REPAGOS_TAB = 'lookerv2';
const MAX_ROWS = 5000;
const CACHE_KEY = 'dash_data_v1';
const CACHE_TTL = 300; // 5 minutos

function doGet(e) {
  try {
    const callback = e && e.parameter && e.parameter.callback;
    const forceRefresh = e && e.parameter && e.parameter.force === '1';

    // Intentar leer del caché primero (segunda llamada en adelante será instantánea)
    let json = null;
    if (!forceRefresh) {
      try {
        const cache = CacheService.getScriptCache();
        json = cache.get(CACHE_KEY);
        if (json) Logger.log('Sirviendo desde caché');
      } catch(ce) { /* ignorar errores de caché */ }
    }

    if (!json) {
      Logger.log('Leyendo planillas...');
      const baseMadre = readSheet(BASE_MADRE_ID, BASE_MADRE_TAB);
      const lookerv2 = readSheet(REPAGOS_ID, REPAGOS_TAB);
      const output = {
        timestamp: new Date().toISOString(),
        cached: false,
        baseMadre: baseMadre,
        lookerv2: lookerv2
      };
      json = JSON.stringify(output);
      // Guardar en caché (límite 100KB del CacheService)
      try {
        const cache = CacheService.getScriptCache();
        if (json.length < 90000) {
          cache.put(CACHE_KEY, json, CACHE_TTL);
          Logger.log('Datos guardados en caché');
        }
      } catch(ce) { /* ignorar */ }
    }

    // Soporte JSONP: evita restricciones CORS completamente
    if (callback) {
      return ContentService
        .createTextOutput(callback + '(' + json + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService
      .createTextOutput(json)
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    const errorOutput = JSON.stringify({ error: err.message, timestamp: new Date().toISOString() });
    const callback = e && e.parameter && e.parameter.callback;
    if (callback) {
      return ContentService
        .createTextOutput(callback + '(' + errorOutput + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(errorOutput).setMimeType(ContentService.MimeType.JSON);
  }
}

function readSheet(spreadsheetId, sheetName) {
  try {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      Logger.log('Hoja no encontrada: ' + sheetName);
      return [];
    }
    const lastRow = Math.min(sheet.getLastRow(), MAX_ROWS + 1);
    const lastCol = sheet.getLastColumn();
    if (lastRow < 1 || lastCol < 1) return [];
    return sheet.getRange(1, 1, lastRow, lastCol).getValues();
  } catch (err) {
    Logger.log('Error leyendo hoja ' + sheetName + ': ' + err.message);
    return [];
  }
}

function testRead() {
  const bm = readSheet(BASE_MADRE_ID, BASE_MADRE_TAB);
  Logger.log('Base Madre — filas: ' + bm.length);
  if (bm.length > 0) Logger.log('Headers: ' + bm[0].join(' | '));

  const rep = readSheet(REPAGOS_ID, REPAGOS_TAB);
  Logger.log('Looker v2 — filas: ' + rep.length);
  if (rep.length > 0) Logger.log('Headers: ' + rep[0].join(' | '));
}
