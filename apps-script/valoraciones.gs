/**
 * BiciFriendly · Valoraciones y comentarios desde la misma página (sin Google Forms).
 *
 * Va en el MISMO proyecto de Apps Script que formularios.gs y registro.gs, como un archivo
 * nuevo llamado "valoraciones". Usa la misma aplicación web (la URL /exec del registro).
 *
 * Cómo funciona:
 *  - La página envía { tipo: 'valorar', id, n (1 a 5), c (comentario opcional), sid }.
 *    registro.gs > doPost se lo pasa a valorarWeb_(), que lo anota en la hoja "Valoraciones web".
 *  - Las estrellas cuentan al tiro (se recalcula "Resumen").
 *  - Los comentarios NO se publican solos: aparecen en la página solo cuando escribes "Sí"
 *    en la columna "Comentario aprobado". Si escribes "Sí" en "Oculto", el voto deja de contar.
 *  - doGet (?q=valoraciones) entrega a la página el resumen de estrellas y los comentarios aprobados.
 *
 * Límites contra malas prácticas:
 *  - Un voto vigente por navegador y lugar: si vuelve a votar, el anterior queda "Reemplazado".
 *  - Máximo VW_CAMBIOS_DIA envíos por navegador y lugar al día, y VW_MAX_DIA lugares por navegador al día.
 *  - Por lugar, entre todos: máximo RAFAGA_MAX votos en RAFAGA_MIN minutos y DIA_MAX en 24 horas
 *    (mismas reglas que el formulario; lo que exceda se ignora y queda en la hoja "Alertas").
 *  - Tope global de VW_MAX_MIN valoraciones por minuto para todo el sitio.
 *  - Solo lugares que existen en bicifriendly.cl/places.json; comentario de hasta 500 caracteres,
 *    sin fórmulas ni caracteres invisibles.
 *
 * Después de pegar este archivo: Implementar > Gestionar implementaciones > (lápiz) >
 * Versión: Nueva versión > Implementar. La URL /exec sigue siendo la misma.
 */

const VW_HOJA = 'Valoraciones web';
const VW_COLS = ['Fecha', 'Lugar', 'Estrellas', 'Comentario', 'Sesión', 'Estado', 'Comentario aprobado', 'Oculto'];
const VW_MAX_DIA = 8;        // lugares distintos que un navegador puede valorar en 24 horas
const VW_CAMBIOS_DIA = 3;    // envíos por navegador y lugar en 24 horas (incluye cambios de opinión)
const VW_MAX_MIN = 30;       // valoraciones por minuto para todo el sitio
const VW_MAX_COMENTARIO = 500;
const VW_COMENTARIOS_POR_LUGAR = 5;

function valorarWeb_(body) {
  const sid = String(body.sid || '');
  const id = String(body.id || '');
  const n = Number(body.n);
  if (!/^[a-z0-9]{8,24}$/.test(sid)) return 'sesion';
  if (!/^[a-z0-9-]{2,60}$/.test(id) || !vwLugarExiste_(id)) return 'lugar';
  if (!Number.isInteger(n) || n < 1 || n > 5) return 'estrellas';
  const comentario = regTxt_(String(body.c || '').replace(/\s+/g, ' ').trim(), VW_MAX_COMENTARIO);

  // Cuotas con la caché del script: no se guarda nada personal.
  const cache = CacheService.getScriptCache();
  const dia = Math.floor(Date.now() / 86400000), min = Math.floor(Date.now() / 60000);
  const kMin = 'vw:m:' + min, kLug = 'vw:l:' + sid + ':' + id + ':' + dia, kSes = 'vw:s:' + sid + ':' + dia;
  const usadoMin = Number(cache.get(kMin) || 0);
  if (usadoMin >= VW_MAX_MIN) return 'cupo';
  const usadoLug = Number(cache.get(kLug) || 0);
  if (usadoLug >= VW_CAMBIOS_DIA) return 'cambios';
  const lugaresHoy = JSON.parse(cache.get(kSes) || '[]');
  const nuevoLugar = lugaresHoy.indexOf(id) === -1;
  if (nuevoLugar && lugaresHoy.length >= VW_MAX_DIA) return 'dia';
  cache.put(kMin, String(usadoMin + 1), 120);
  cache.put(kLug, String(usadoLug + 1), 90000);
  if (nuevoLugar) { lugaresHoy.push(id); cache.put(kSes, JSON.stringify(lugaresHoy), 90000); }

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const hoja = regHoja_(VW_HOJA, VW_COLS);
    const ultima = hoja.getLastRow();
    // Un voto vigente por navegador y lugar: los anteriores pasan a "Reemplazado".
    if (ultima > 1) {
      const filas = hoja.getRange(2, 2, ultima - 1, 5).getValues(); // Lugar..Estado
      for (let i = 0; i < filas.length; i++) {
        if (String(filas[i][0]) === id && String(filas[i][3]) === sid && String(filas[i][4]) === 'Cuenta') {
          hoja.getRange(i + 2, 6).setValue('Reemplazado');
        }
      }
    }
    hoja.appendRow([new Date(), id, n, comentario, sid, 'Cuenta', '', '']);
  } finally { lock.releaseLock(); }

  actualizarResumen();
  CacheService.getScriptCache().remove('vw:get');
  return 'ok';
}

// Lista de ids publicados en el sitio (se guarda 6 horas en caché).
function vwLugarExiste_(id) {
  const cache = CacheService.getScriptCache();
  let ids = cache.get('vw:ids');
  if (!ids) {
    try {
      const r = UrlFetchApp.fetch(URL_LUGARES + '?t=' + Date.now(), { muteHttpExceptions: true });
      const lista = r.getResponseCode() === 200 ? JSON.parse(r.getContentText()) : [];
      if (Array.isArray(lista) && lista.length >= 5) {
        ids = '|' + lista.map(function (p) { return p.id; }).join('|') + '|';
        cache.put('vw:ids', ids, 21600);
      }
    } catch (e) { console.warn('vwLugarExiste_: ' + e); }
  }
  // Si places.json no se puede leer, se acepta el id con formato válido (ya filtrado antes).
  return ids ? ids.indexOf('|' + id + '|') !== -1 : true;
}

/** Filas de "Valoraciones web" que cuentan, en el formato que usa actualizarResumen. */
function vwFilasParaResumen_(ss) {
  const h = ss.getSheetByName(VW_HOJA);
  if (!h || h.getLastRow() < 2) return [];
  return h.getRange(2, 1, h.getLastRow() - 1, VW_COLS.length).getValues()
    .filter(function (f) { return String(f[5]) === 'Cuenta'; })
    .map(function (f) { return [f[0], '(' + f[1] + ')', f[2], '', '', f[7]]; });
}

/** La página pide ?q=valoraciones: resumen de estrellas y comentarios aprobados. */
function doGet(e) {
  const q = e && e.parameter && e.parameter.q;
  if (q !== 'valoraciones') return ContentService.createTextOutput('ok');
  const cache = CacheService.getScriptCache();
  let json = cache.get('vw:get');
  if (!json) {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const resumen = {};
    const res = ss.getSheetByName('Resumen');
    if (res && res.getLastRow() > 1) {
      res.getRange(2, 1, res.getLastRow() - 1, 3).getValues().forEach(function (f) {
        if (f[0] && Number(f[2]) > 0) resumen[String(f[0])] = [Number(f[1]), Number(f[2])];
      });
    }
    const comentarios = {};
    const h = ss.getSheetByName(VW_HOJA);
    if (h && h.getLastRow() > 1) {
      const filas = h.getRange(2, 1, h.getLastRow() - 1, VW_COLS.length).getValues();
      for (let i = filas.length - 1; i >= 0; i--) {
        const f = filas[i];
        const ok = String(f[6]).trim().toLowerCase(), oculto = String(f[7]).trim().toLowerCase();
        if (!f[3] || String(f[5]) !== 'Cuenta' || (ok !== 'sí' && ok !== 'si') || oculto === 'sí' || oculto === 'si') continue;
        const lista = comentarios[f[1]] || (comentarios[f[1]] = []);
        if (lista.length >= VW_COMENTARIOS_POR_LUGAR) continue;
        const fecha = f[0] instanceof Date ? f[0] : new Date(f[0]);
        lista.push([Number(f[2]), String(f[3]).replace(/^'/, ''), Utilities.formatDate(fecha, 'America/Santiago', 'yyyy-MM')]);
      }
    }
    json = JSON.stringify({ resumen: resumen, comentarios: comentarios });
    cache.put('vw:get', json, 120);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

/** Al aprobar u ocultar algo a mano en la planilla, la página se entera en menos de 2 minutos. */
function vwAlEditar(e) {
  try {
    if (e && e.range && e.range.getSheet().getName() === VW_HOJA) CacheService.getScriptCache().remove('vw:get');
  } catch (x) {}
}

/** Ejecutar una vez: crea la hoja y el activador que refresca la página al aprobar comentarios. */
function instalarValoraciones() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const h = regHoja_(VW_HOJA, VW_COLS);
  const regla = SpreadsheetApp.newDataValidation().requireValueInList(['Sí'], true).setAllowInvalid(true).build();
  h.getRange(2, 7, h.getMaxRows() - 1, 2).setDataValidation(regla);
  h.setColumnWidth(4, 420);
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'vwAlEditar') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('vwAlEditar').forSpreadsheet(ss).onEdit().create();
  actualizarResumen();
  console.log('Listo: hoja "' + VW_HOJA + '" creada. Recuerda publicar una nueva versión de la implementación.');
}
