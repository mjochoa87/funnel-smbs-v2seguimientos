/**
 * Google Apps Script — Dashboard Unificado Principalidad
 * Con soporte JSONP, CacheService (5 min) y filtro de mes dinamico en LOOKERV2
 *
 * INSTRUCCIONES DE DEPLOY:
 * 1. Abri script.google.com con tu cuenta MeLi
 * 2. Pega este codigo en Codigo.gs
 * 3. Guardar (Ctrl+S)
 * 4. Implementar -> Nueva implementacion -> Aplicacion web
 *    - Ejecutar como: Yo
 *    - Quienes tienen acceso: Cualquiera
 * 5. Copia la URL del tipo: https://script.google.com/macros/s/XXXX/exec
 */

const REPAGOS_ID = '11xFxl_XYFIhLGmokYJM9HpBUu53uRoUhWNqdqfHVPs8';
const REPAGOS_TAB = 'lookerv2';
const REPAGOS_TARGET_TAB = 'target por asesor';
const RESUMEN_GESTION_TAB = 'RESUMEN GESTION';
const MAX_ROWS = 5000;
const MESES_HISTORIAL = 4; // Cuantos meses hacia atras incluir en LOOKERV2 (columna B)
const CACHE_KEY = 'dash_unificado_v2';
const CACHE_TTL = 300; // 5 minutos

// Genera array con los ultimos N meses en formato YYYYMM
function getUltimosMeses(n) {
  var hoy = new Date();
  var meses = [];
  for (var i = 0; i < n; i++) {
    var d = new Date(hoy.getFullYear(), hoy.getMonth() - i, 1);
    var yyyymm = d.getFullYear() * 100 + (d.getMonth() + 1);
    meses.push(yyyymm);
  }
  return meses;
}

function doGet(e) {
  try {
    var callback = e && e.parameter && e.parameter.callback;
    var forceRefresh = e && e.parameter && e.parameter.force === '1';

    var json = null;
    if (!forceRefresh) {
      try {
        var cache = CacheService.getScriptCache();
        json = cache.get(CACHE_KEY);
        if (json) Logger.log('Sirviendo desde cache');
      } catch(ce) {}
    }

    if (!json) {
      Logger.log('Leyendo planillas...');
      var lookerv2 = readLookerV2Filtrado();
      var output = {
        timestamp: new Date().toISOString(),
        cached: false,
        lookerv2: lookerv2,
        targetPorAsesor: getTargetPorAsesor(),
        resumenGestion: getResumenGestion()
      };
      json = JSON.stringify(output);
      try {
        var cache2 = CacheService.getScriptCache();
        if (json.length < 90000) {
          cache2.put(CACHE_KEY, json, CACHE_TTL);
          Logger.log('Guardado en cache. Tamano: ' + json.length + ' bytes');
        } else {
          Logger.log('JSON demasiado grande para cache: ' + json.length + ' bytes — se sirve sin cache');
        }
      } catch(ce) {}
    }

    if (callback) {
      return ContentService.createTextOutput(callback + '(' + json + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    var errorOutput = JSON.stringify({ error: err.message, timestamp: new Date().toISOString() });
    var cb = e && e.parameter && e.parameter.callback;
    if (cb) {
      return ContentService.createTextOutput(cb + '(' + errorOutput + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(errorOutput).setMimeType(ContentService.MimeType.JSON);
  }
}

// Lee LOOKERV2 filtrando por columna B (mes en formato YYYYMM), solo ultimos MESES_HISTORIAL meses
function readLookerV2Filtrado() {
  try {
    var ss = SpreadsheetApp.openById(REPAGOS_ID);
    var sheet = ss.getSheetByName(REPAGOS_TAB);
    if (!sheet) { Logger.log('Hoja no encontrada: ' + REPAGOS_TAB); return []; }

    var lastRow = Math.min(sheet.getLastRow(), MAX_ROWS + 1);
    var lastCol = sheet.getLastColumn();
    if (lastRow < 2) return [];

    var allRows = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    var mesesValidos = {};
    getUltimosMeses(MESES_HISTORIAL).forEach(function(m) { mesesValidos[String(m)] = true; });

    var result = [allRows[0]]; // header siempre incluido
    var filtradas = 0;
    for (var i = 1; i < allRows.length; i++) {
      var mesVal = String(allRows[i][1] || '').trim(); // columna B = indice 1
      if (mesesValidos[mesVal]) {
        result.push(allRows[i]);
        filtradas++;
      }
    }

    Logger.log('LOOKERV2 filtrado: ' + filtradas + ' filas (de ' + (allRows.length - 1) + ' totales)');
    return result;
  } catch (err) {
    Logger.log('Error readLookerV2Filtrado: ' + err.message);
    return [];
  }
}

function readSheet(spreadsheetId, sheetName) {
  try {
    var ss = SpreadsheetApp.openById(spreadsheetId);
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) { Logger.log('Hoja no encontrada: ' + sheetName); return []; }
    var lastRow = Math.min(sheet.getLastRow(), MAX_ROWS + 1);
    var lastCol = sheet.getLastColumn();
    if (lastRow < 1 || lastCol < 1) return [];
    return sheet.getRange(1, 1, lastRow, lastCol).getValues();
  } catch (err) {
    Logger.log('Error leyendo ' + sheetName + ': ' + err.message);
    return [];
  }
}

/**
 * Lee "target por asesor": columna A = ASESOR, C = MES-COHORT, I = OBJETIVO DC AJUSTADO
 */
function getTargetPorAsesor() {
  try {
    var rows = readSheet(REPAGOS_ID, REPAGOS_TARGET_TAB);
    if (!rows || rows.length < 2) return [];
    var result = [];
    for (var i = 1; i < rows.length; i++) {
      var row = rows[i];
      var asesor = String(row[0] || '').trim();
      var cohort = String(row[2] || '').trim();
      var objetivo = row[8];
      if (asesor && cohort && objetivo !== '' && objetivo !== null && objetivo !== undefined) {
        result.push([asesor, cohort, objetivo]);
      }
    }
    Logger.log('Target por asesor: ' + result.length + ' filas');
    return result;
  } catch(e) {
    Logger.log('Error getTargetPorAsesor: ' + e.message);
    return [];
  }
}

/**
 * Lee RESUMEN GESTION dinamicamente — sin hardcodear meses ni equipos.
 * Devuelve array de [asesor, team, mes1, mes2, mes3] donde mes1/2/3 son los ultimos 3 meses.
 */
function getResumenGestion() {
  try {
    var rows = readSheet(REPAGOS_ID, RESUMEN_GESTION_TAB);
    if (!rows || rows.length < 2) return [];

    // Encontrar fila de headers (la que tiene meses en formato YYYYMM)
    var headerIdx = -1;
    var mesRegex = /^20\d{4}$/;
    for (var i = 0; i < rows.length; i++) {
      for (var j = 0; j < rows[i].length; j++) {
        if (mesRegex.test(String(rows[i][j]).trim())) { headerIdx = i; break; }
      }
      if (headerIdx >= 0) break;
    }
    if (headerIdx < 0) { Logger.log('No se encontro fila de headers en RESUMEN GESTION'); return []; }

    var header = rows[headerIdx].map(function(c) { return String(c).trim(); });
    var ultMeses = getUltimosMeses(3); // ultimos 3 meses
    var mesIdxMap = {};
    ultMeses.forEach(function(m) {
      var idx = header.indexOf(String(m));
      if (idx >= 0) mesIdxMap[m] = idx;
    });

    var toNum = function(v) {
      if (v === '' || v === null || v === undefined) return null;
      var n = parseFloat(String(v).replace(',', '.'));
      return isNaN(n) ? null : Math.round(n * 10000) / 100;
    };

    var result = [];
    for (var i = headerIdx + 1; i < rows.length; i++) {
      var row = rows[i];
      var asesor = String(row[1] || '').trim();
      var team = String(row[2] || '').trim();
      if (!asesor) continue;
      var entry = [asesor, team];
      ultMeses.forEach(function(m) {
        entry.push(mesIdxMap[m] !== undefined ? toNum(row[mesIdxMap[m]]) : null);
      });
      result.push(entry);
    }

    Logger.log('getResumenGestion: ' + result.length + ' filas');
    return result;
  } catch(e) {
    Logger.log('getResumenGestion error: ' + e.message);
    return [];
  }
}

function testRead() {
  var rows = readLookerV2Filtrado();
  Logger.log('LookerV2 filtrado — filas: ' + rows.length);
  if (rows.length > 0) Logger.log('Headers: ' + rows[0].join(' | '));
  var target = getTargetPorAsesor();
  Logger.log('Target: ' + target.length + ' filas');
  var resumen = getResumenGestion();
  Logger.log('Resumen gestion: ' + resumen.length + ' filas');
}
