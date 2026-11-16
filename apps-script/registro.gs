/**
 * BiciFriendly · Registro de visitas y uso + contadores.
 *
 * Va en el MISMO proyecto de Apps Script que formularios.gs (Extensiones > Apps Script
 * desde la planilla "Bicifriendly · Base de lugares"), como un archivo nuevo llamado "registro".
 *
 * Qué hace:
 *  - doPost: recibe los eventos que manda la página (visita, ficha, ir_bici, mapas, filtro,
 *    cerca, guardar, buscar, sugerir, valorar, tema) y los anota en la hoja "Visitas".
 *    No guarda IP ni datos personales: solo un identificador al azar por navegador.
 *  - actualizarContadores: resume la hoja "Visitas" en la hoja "Contadores"
 *    (totales del sitio y por lugar). Se ejecuta sola cada hora.
 *
 * Las valoraciones hechas en la página también llegan por doPost (ver valoraciones.gs).
 *
 * Pasos (una sola vez):
 *  1. Pega este archivo y guarda.
 *  2. Ejecuta instalarRegistro (acepta los permisos). Crea las hojas y el activador por hora.
 *  3. Implementar > Nueva implementación > Tipo: Aplicación web.
 *     Ejecutar como: Yo. Quién tiene acceso: Cualquier usuario. Implementar.
 *  4. Copia la URL que termina en /exec y pásala a la página (CONFIG.registroUrl).
 *  5. Opcional, para mostrar contadores públicos en la página: en la planilla,
 *     Archivo > Compartir > Publicar en la web > hoja "Contadores" > CSV, y pasa ese enlace
 *     (CONFIG.contadoresCsvUrl). Publica SOLO "Contadores", nunca "Visitas".
 */

const REG_HOJA_VISITAS = 'Visitas';
const REG_HOJA_CONTADORES = 'Contadores';
const REG_EVENTOS = ['visita', 'ficha', 'ir_bici', 'mapas', 'filtro', 'cerca', 'guardar', 'buscar', 'sugerir', 'valorar', 'tema', 'error'];
// Límites contra abuso: por navegador y en total, por minuto.
const REG_MAX_POR_SESION = 60;     // eventos por navegador y minuto
const REG_MAX_TOTAL = 300;         // eventos totales por minuto (todas las sesiones)
const REG_MAX_BLOQUE = 6000;       // eventos totales por bloque de 6 horas
const REG_MAX_FILAS = 150000;      // filas máximas en "Visitas"; las más antiguas se borran
const REG_MAX_BYTES = 20000;
const REG_COLUMNAS = ['Fecha', 'Sesión', 'Evento', 'Lugar', 'Detalle', 'Dispositivo', 'Tema', 'Origen'];

function doPost(e) {
  try {
    const raw = (e && e.postData && e.postData.contents) || '';
    if (!raw || raw.length > REG_MAX_BYTES) return ContentService.createTextOutput('ok');
    const body = JSON.parse(raw);
    // Valoración enviada desde la ficha de un lugar (ver valoraciones.gs).
    if (body && body.tipo === 'valorar') return ContentService.createTextOutput(valorarWeb_(body));
    const sid = regTxt_(body.sid, 40) || 'x';
    let lista = Array.isArray(body.ev) ? body.ev.slice(0, 50) : [];
    // Cuotas por minuto con la caché del script (no guarda nada en la planilla).
    const cache = CacheService.getScriptCache();
    const min = Math.floor(Date.now() / 60000), bloque = Math.floor(Date.now() / 21600000);
    const kS = 's:' + sid + ':' + min, kT = 't:' + min, kB = 'b:' + bloque;
    const usadoS = Number(cache.get(kS) || 0), usadoT = Number(cache.get(kT) || 0), usadoB = Number(cache.get(kB) || 0);
    const cupo = Math.max(0, Math.min(REG_MAX_POR_SESION - usadoS, REG_MAX_TOTAL - usadoT, REG_MAX_BLOQUE - usadoB));
    if (cupo <= 0) return ContentService.createTextOutput('ok');
    lista = lista.slice(0, cupo);
    cache.put(kS, String(usadoS + lista.length), 120);
    cache.put(kT, String(usadoT + lista.length), 120);
    cache.put(kB, String(usadoB + lista.length), 21600);
    const ahora = new Date();
    const filas = lista
      .filter(function (v) { return v && REG_EVENTOS.indexOf(String(v.e)) !== -1; })
      .map(function (v) {
        return [ahora, sid, regTxt_(v.e, 20), regTxt_(v.id, 60), regTxt_(v.d, 160),
                regTxt_(body.dev, 12), regTxt_(body.th, 8), regTxt_(body.ref, 60)];
      });
    if (filas.length) {
      const lock = LockService.getScriptLock();
      lock.waitLock(10000);
      try {
        const hoja = regHoja_(REG_HOJA_VISITAS, REG_COLUMNAS);
        hoja.getRange(hoja.getLastRow() + 1, 1, filas.length, REG_COLUMNAS.length).setValues(filas);
      } finally { lock.releaseLock(); }
    }
  } catch (err) { console.warn('doPost: ' + err); /* nunca romper la página por un error de registro */ }
  return ContentService.createTextOutput('ok');
}

/** Resume "Visitas" en "Contadores". Primera fila: totales del sitio (id = _sitio). */
function actualizarContadores() {
  const hoja = regHoja_(REG_HOJA_VISITAS, REG_COLUMNAS);
  // Mantención: si "Visitas" crece demasiado (por uso normal o un ataque), se borran las filas más antiguas.
  const exceso = hoja.getLastRow() - 1 - REG_MAX_FILAS;
  if (exceso > 0) hoja.deleteRows(2, exceso);
  const datos = hoja.getLastRow() > 1 ? hoja.getRange(2, 1, hoja.getLastRow() - 1, REG_COLUMNAS.length).getValues() : [];
  const ahora = Date.now(), dia = 86400000;
  const hoyIni = new Date(); hoyIni.setHours(0, 0, 0, 0);
  const sitio = { visitas: 0, hoy: 0, d7: 0, d30: 0, sesiones: {}, ficha: 0, ir_bici: 0, mapas: 0, guardar: 0, cerca: 0, filtro: 0, buscar: 0, sugerir: 0, valorar: 0, error: 0, error7: 0, movil: 0, escritorio: 0 };
  const lugares = {};
  datos.forEach(function (r) {
    const f = r[0] instanceof Date ? r[0].getTime() : new Date(r[0]).getTime();
    const ev = String(r[2]), id = String(r[3]);
    if (ev === 'visita') {
      sitio.visitas++;
      if (f >= hoyIni.getTime()) sitio.hoy++;
      if (ahora - f <= 7 * dia) sitio.d7++;
      if (ahora - f <= 30 * dia) sitio.d30++;
      sitio.sesiones[r[1]] = 1;
      if (r[5] === 'movil') sitio.movil++; else if (r[5] === 'escritorio') sitio.escritorio++;
    } else if (sitio[ev] !== undefined) { sitio[ev]++; if (ev === 'error' && ahora - f <= 7 * dia) sitio.error7++; }
    if (id && ['ficha', 'ir_bici', 'mapas', 'guardar'].indexOf(ev) !== -1) {
      const l = lugares[id] || (lugares[id] = { ficha: 0, ir_bici: 0, mapas: 0, guardar: 0 });
      l[ev]++;
    }
  });
  const cab = ['id', 'visitas', 'visitas_hoy', 'visitas_7d', 'visitas_30d', 'personas', 'fichas', 'ir_bici', 'mapas', 'guardados', 'cerca_de_mi', 'filtros', 'busquedas', 'sugerir', 'valorar', 'errores', 'errores_7d', 'movil', 'escritorio', 'actualizado'];
  const filas = [['_sitio', sitio.visitas, sitio.hoy, sitio.d7, sitio.d30, Object.keys(sitio.sesiones).length, sitio.ficha, sitio.ir_bici, sitio.mapas, sitio.guardar, sitio.cerca, sitio.filtro, sitio.buscar, sitio.sugerir, sitio.valorar, sitio.error, sitio.error7, sitio.movil, sitio.escritorio, new Date()]];
  Object.keys(lugares).sort(function (a, b) { return (lugares[b].ir_bici + lugares[b].ficha) - (lugares[a].ir_bici + lugares[a].ficha); })
    .forEach(function (id) { const l = lugares[id]; filas.push([id, '', '', '', '', '', l.ficha, l.ir_bici, l.mapas, l.guardar, '', '', '', '', '', '', '', '', '', '']); });
  const c = regHoja_(REG_HOJA_CONTADORES, cab);
  if (c.getLastRow() > 1) c.getRange(2, 1, c.getLastRow() - 1, cab.length).clearContent();
  c.getRange(2, 1, filas.length, cab.length).setValues(filas);
}

/** Ejecutar una vez: crea las hojas y el activador que actualiza los contadores cada hora. */
function instalarRegistro() {
  regHoja_(REG_HOJA_VISITAS, REG_COLUMNAS);
  regHoja_(REG_HOJA_CONTADORES, ['id']);
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'actualizarContadores') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('actualizarContadores').timeBased().everyHours(1).create();
  actualizarContadores();
}

function regHoja_(nombre, cabecera) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let h = ss.getSheetByName(nombre);
  if (!h) { h = ss.insertSheet(nombre); }
  if (h.getLastRow() === 0 || h.getRange(1, 1).getValue() !== cabecera[0] || h.getLastColumn() < cabecera.length) {
    h.getRange(1, 1, 1, cabecera.length).setValues([cabecera]).setFontWeight('bold');
    h.setFrozenRows(1);
  }
  return h;
}

// Texto corto y sin fórmulas (evita que algo como "=..." se ejecute en la planilla).
function regTxt_(x, n) {
  // Quita caracteres de control e invisibles, espacios al inicio, y neutraliza fórmulas.
  const t = String(x == null ? '' : x).replace(/[\u0000-\u001F\u007F-\u009F\u200B-\u200F\u2028-\u202F\u2060-\u206F\uFEFF]/g, '').replace(/^\s+/, '').slice(0, n);
  return /^[=+\-@]/.test(t) ? "'" + t : t;
}
