/**
 * Bicifriendly · Formularios de la comunidad
 *
 * Crea dos Google Forms públicos conectados a la planilla "Bicifriendly · Base de lugares":
 *   1. "Sugiere un lugar"  -> respuestas en la hoja "Aportes"
 *   2. "Valora un lugar"   -> respuestas en la hoja "Valoraciones" (estrellas 1 a 5 + comentario)
 * y una hoja "Resumen" con el promedio de estrellas y el número de votos por lugar,
 * que la página lee para mostrar las estrellas.
 *
 * CÓMO USARLO
 *   1. Abre la planilla > Extensiones > Apps Script.
 *   2. Borra lo que aparezca, pega este archivo completo y guarda.
 *   3. Elige la función "crearFormularios" arriba y presiona Ejecutar. Acepta los permisos.
 *   4. Revisa la hoja "Config": ahí quedan los links que hay que pasar a la página.
 *
 * Ejecútalo una sola vez. Para sumar lugares nuevos a la lista del formulario de valoración,
 * usa "actualizarLugares" (además corre sola una vez al día).
 */

const SHEET_ID = '1NdzCOAtkJ5RBJFyT-tFAydwUjtzXS9FrIjmnauJf36w';

const TIPOS = ['Café', 'Bar', 'Comida rápida', 'Restaurant', 'Punto ciclista'];

function crearFormularios() {
  const props = PropertiesService.getScriptProperties();
  let yaExiste = !!props.getProperty('FORM_APORTES_ID');
  if (!yaExiste) { try { formId_('FORM_APORTES_ID'); yaExiste = true; } catch (e) {} }
  if (yaExiste) {
    throw new Error('Los formularios ya fueron creados. Revisa la hoja "Config". Para refrescar la lista de lugares usa actualizarLugares.');
  }
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const lugares = ss.getSheets()[0];
  if (lugares.getName() !== 'Lugares') lugares.setName('Lugares');

  // ---------- 1. Formulario de aportes ----------
  const fa = FormApp.create('Bicifriendly · Sugiere un lugar');
  fa.setDescription(
    'Ayúdanos a armar el mapa de lugares bikefriendly de Santiago. ' +
    'Cuéntanos de un café, bar, restaurant o punto donde se reciba bien a los ciclistas. ' +
    'Revisamos cada aporte antes de publicarlo en bicifriendly.cl.'
  );
  fa.setCollectEmail(false);
  fa.setAllowResponseEdits(false);
  fa.setLimitOneResponsePerUser(false);
  fa.setProgressBar(false);
  fa.setConfirmationMessage('¡Gracias! Revisaremos tu aporte y, si todo calza, aparecerá en bicifriendly.cl.');

  fa.addTextItem().setTitle('Nombre del lugar').setRequired(true);
  fa.addMultipleChoiceItem().setTitle('Tipo de lugar')
    .setChoiceValues(TIPOS).showOtherOption(true).setRequired(true);
  fa.addTextItem().setTitle('Dirección').setHelpText('Calle y número, o una referencia clara.').setRequired(true);
  fa.addTextItem().setTitle('Comuna').setRequired(true);
  fa.addTextItem().setTitle('Link de Google Maps (opcional)')
    .setHelpText('En Google Maps: busca el lugar > Compartir > Copiar vínculo.');
  fa.addCheckboxItem().setTitle('¿Qué tiene para ciclistas?')
    .setChoiceValues([
      'Bicicletero o lugar seguro para dejar la bici',
      'Te dejan entrar con la bici',
      'Taller o herramientas',
      'Descuento a ciclistas',
      'Agua o rellenan botellas',
      'Está junto a una ciclovía',
      'Parada típica de rutas (ruta o MTB)'
    ]).showOtherOption(true).setRequired(true);
  fa.addParagraphTextItem().setTitle('Cuéntanos más (opcional)')
    .setHelpText('Ej: el bicicletero es techado, el descuento es de 10%, abren temprano los domingos.');
  fa.addMultipleChoiceItem().setTitle('¿Cómo lo sabes?')
    .setChoiceValues(['Voy seguido', 'Fui una vez', 'Me contaron', 'Trabajo ahí o es mi local'])
    .setRequired(true);
  fa.addTextItem().setTitle('Tu Instagram (opcional)').setHelpText('Para darte crédito o preguntarte algo.');

  fa.setDestination(FormApp.DestinationType.SPREADSHEET, SHEET_ID);

  // ---------- 2. Formulario de valoración ----------
  const fv = FormApp.create('Bicifriendly · Valora un lugar');
  fv.setDescription('¿Fuiste en bici? Cuéntanos cómo te recibieron. Tu valoración aparece en bicifriendly.cl.');
  fv.setCollectEmail(false);
  fv.setAllowResponseEdits(false);
  fv.setLimitOneResponsePerUser(false);
  fv.setConfirmationMessage('¡Gracias por tu valoración!');

  const lista = fv.addListItem().setTitle('Lugar').setRequired(true);
  const opciones = opcionesLugares_(lugares);
  lista.setChoiceValues(opciones);

  // Estrellas 1 a 5. Si la cuenta no tiene el tipo "Calificación", usa una escala 1 a 5.
  try {
    fv.addRatingItem().setTitle('¿Qué tan bikefriendly es?')
      .setRatingScaleLevel(5).setRatingIcon(FormApp.RatingIconType.STAR).setRequired(true);
  } catch (e) {
    fv.addScaleItem().setTitle('¿Qué tan bikefriendly es?')
      .setBounds(1, 5).setLabels('Nada', 'Muy bikefriendly').setRequired(true);
  }
  fv.addParagraphTextItem().setTitle('Comentario (opcional)')
    .setHelpText('Ej: dejé la bici a la vista, me rellenaron la botella, buen café post-ruta.');
  fv.addTextItem().setTitle('Tu Instagram (opcional)');

  fv.setDestination(FormApp.DestinationType.SPREADSHEET, SHEET_ID);

  // ---------- 3. Renombrar hojas de respuestas ----------
  SpreadsheetApp.flush();
  Utilities.sleep(3000);
  renombrarHojaDeForm_(ss, fa.getId(), 'Aportes');
  const hv = renombrarHojaDeForm_(ss, fv.getId(), 'Valoraciones');
  if (hv) {
    // Columna para moderar: escribir "Sí" oculta esa valoración del promedio.
    hv.getRange(1, 6).setValue('Oculto (Sí = no cuenta)');
  }

  // ---------- 4. Hoja Resumen (la lee la página) ----------
  let res = ss.getSheetByName('Resumen') || ss.insertSheet('Resumen');
  res.clear();
  // Valoraciones: A marca temporal, B lugar "Nombre (id)", C estrellas, D comentario, E instagram, F oculto
  res.getRange('A1').setFormula(
    '=IFERROR(QUERY(FILTER({REGEXEXTRACT(Valoraciones!B2:B,"\\(([^()]+)\\)$"),Valoraciones!C2:C},' +
    'Valoraciones!B2:B<>"",Valoraciones!C2:C<>"",Valoraciones!F2:F<>"Sí"),' +
    '"select Col1, avg(Col2), count(Col2) group by Col1 label Col1 \'id\', avg(Col2) \'promedio\', count(Col2) \'votos\'",0),' +
    '{"id","promedio","votos"})'
  );
  res.getRange('B:B').setNumberFormat('0.0');

  // ---------- 5. Config ----------
  const entry = idEntradaLugar_(fv, lista, opciones[0]);
  const urlValorar = fv.getPublishedUrl().replace(/\/viewform.*$/, '/viewform') + '?usp=pp_url&entry.' + entry + '=';
  let cfg = ss.getSheetByName('Config') || ss.insertSheet('Config');
  cfg.clear();
  cfg.getRange(1, 1, 6, 2).setValues([
    ['Qué es', 'Valor'],
    ['Link formulario "Sugiere un lugar"', fa.getPublishedUrl()],
    ['Link base formulario "Valora un lugar" (la página agrega el lugar al final)', urlValorar],
    ['Editar formulario de aportes', fa.getEditUrl()],
    ['Editar formulario de valoración', fv.getEditUrl()],
    ['Siguiente paso', 'Archivo > Compartir > Publicar en la web > hoja "Resumen" > CSV > Publicar. Copia ese link y pásalo junto a los dos primeros.']
  ]);
  cfg.setColumnWidth(1, 380); cfg.setColumnWidth(2, 700);

  props.setProperty('FORM_APORTES_ID', fa.getId());
  props.setProperty('FORM_VALORAR_ID', fv.getId());

  // Refresca la lista de lugares del formulario de valoración una vez al día.
  ScriptApp.newTrigger('actualizarLugares').timeBased().everyDays(1).atHour(4).create();

  Logger.log('Sugiere un lugar: ' + fa.getPublishedUrl());
  Logger.log('Valora un lugar (base): ' + urlValorar);
}

// Busca los formularios ya creados aunque este proyecto no tenga guardados sus identificadores
// (por ejemplo, si el código se pegó en un proyecto nuevo). Revisa las hojas conectadas a un
// formulario ("Valoraciones" y "Aportes") y, si no, los links de edición de la hoja "Config".
function formId_(clave) {
  const props = PropertiesService.getScriptProperties();
  const guardado = props.getProperty(clave);
  if (guardado) return guardado;
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const nombre = clave === 'FORM_VALORAR_ID' ? 'Valoraciones' : 'Aportes';
  const palabra = clave === 'FORM_VALORAR_ID' ? 'valora' : 'sugiere';
  let id = null;
  const hojas = ss.getSheets();
  for (let i = 0; i < hojas.length && !id; i++) {
    const url = hojas[i].getFormUrl();
    if (!url) continue;
    const f = FormApp.openByUrl(url);
    if (hojas[i].getName().indexOf(nombre) === 0 || f.getTitle().toLowerCase().indexOf(palabra) !== -1) id = f.getId();
  }
  const cfg = ss.getSheetByName('Config');
  if (!id && cfg) {
    const fila = clave === 'FORM_VALORAR_ID' ? 5 : 4;
    const m = String(cfg.getRange(fila, 2).getValue()).match(/\/forms\/d\/([a-zA-Z0-9_-]+)/);
    if (m) id = m[1];
  }
  if (!id) throw new Error('No encontré el formulario "' + nombre + '". Revisa que la hoja "' + nombre + '" siga conectada a su formulario.');
  props.setProperty(clave, id);
  console.log('Formulario ' + nombre + ' encontrado y guardado: ' + id);
  return id;
}

/** Refresca la lista "Lugar" del formulario de valoración con los lugares publicados. */
function actualizarLugares() {
  const id = formId_('FORM_VALORAR_ID');
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const fv = FormApp.openById(id);
  const lista = fv.getItems(FormApp.ItemType.LIST)[0].asListItem();
  lista.setChoiceValues(opcionesLugares_(ss.getSheetByName('Lugares')));
}

// "Nombre (id)" para cada lugar con Estado = Publicado. La página arma el mismo texto.
function opcionesLugares_(hoja) {
  const filas = hoja.getRange(2, 1, Math.max(hoja.getLastRow() - 1, 1), 16).getValues();
  const out = filas
    .filter(function (f) { return f[0] && f[1] && String(f[15]).trim() === 'Publicado'; })
    .map(function (f) { return String(f[1]).trim() + ' (' + String(f[0]).trim() + ')'; });
  return out.length ? out : ['Sin lugares publicados'];
}

function renombrarHojaDeForm_(ss, formId, nombre) {
  const hojas = ss.getSheets();
  for (let i = 0; i < hojas.length; i++) {
    const url = hojas[i].getFormUrl();
    if (url && FormApp.openByUrl(url).getId() === formId) {
      const previa = ss.getSheetByName(nombre);
      if (previa && previa.getSheetId() !== hojas[i].getSheetId()) previa.setName(nombre + ' (antigua)');
      hojas[i].setName(nombre);
      return hojas[i];
    }
  }
  return null;
}

function idEntradaLugar_(form, item, valor) {
  const url = form.createResponse().withItemResponse(item.createResponse(valor)).toPrefilledUrl();
  const m = url.match(/entry\.(\d+)=/);
  return m ? m[1] : '';
}


/* =====================================================================
 * ARREGLOS (ejecutar una vez): función "arreglarPlanilla"
 *  - Corrige latitud/longitud de los lugares iniciales (la planilla en español
 *    leyó el punto decimal como separador de miles).
 *  - Reemplaza la fórmula de la hoja Resumen por valores que calcula este script,
 *    así no depende de la configuración regional.
 *  - Recalcula el Resumen cada vez que llega una valoración y cuando editas la planilla.
 * ===================================================================== */

const COORDS_INICIALES = {
    holley: [-33.4219, -70.6038],
    draga: [-33.437, -70.6372],
    felix: [-33.4248, -70.609],
    lasubida: [-33.4205, -70.6155],
    goodfriends: [-33.4042, -70.5745],
    singular: [-33.3868, -70.5318],
    tebicite: [-33.4512, -70.5992],
    soulbikes: [-33.7385, -70.7135],
    sucucho: [-33.4446, -70.6292],
    litbust: [-33.4412, -70.6322],
    litbalm: [-33.432, -70.6262],
    labici: [-33.4302, -70.5996],
    liguria: [-33.4286, -70.6178],
    barconst: [-33.4348, -70.6342],
    baco: [-33.4228, -70.607],
    roberta: [-33.4472, -70.6214],
    caramagnola: [-33.4335, -70.5772],
    alemana: [-33.4272, -70.6112],
    fletch: [-33.4352, -70.6326],
    peumayen: [-33.4355, -70.6316],
    rivoli: [-33.4258, -70.6128],
    danoi: [-33.4521, -70.6204],
    elcamino: [-33.4431, -70.6249],
    niguiri: [-33.3592, -70.5182],
    mut: [-33.4172, -70.6012],
    parador: [-33.3528532934991, -70.36875222789324],
    bikecostanera: [-33.417, -70.607],
    ajiamarillo: [-33.428, -70.6168]
};

function arreglarPlanilla() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const hoja = ss.getSheetByName('Lugares');
  const n = hoja.getLastRow() - 1;
  if (n > 0) {
    const ids = hoja.getRange(2, 1, n, 1).getValues();
    const rango = hoja.getRange(2, 6, n, 2);
    const actuales = rango.getValues();
    const nuevos = ids.map(function (f, i) {
      const c = COORDS_INICIALES[String(f[0]).trim()];
      return c ? c : actuales[i];
    });
    rango.setValues(nuevos);
    rango.setNumberFormat('0.000000');
  }

  actualizarResumen();

  // Triggers: al recibir una valoración y al editar la planilla (p. ej. marcar "Oculto").
  const props = PropertiesService.getScriptProperties();
  const fvId = formId_('FORM_VALORAR_ID');
  const existentes = ScriptApp.getProjectTriggers().map(function (t) { return t.getHandlerFunction(); });
  if (fvId && existentes.indexOf('actualizarResumen') === -1) {
    ScriptApp.newTrigger('actualizarResumen').forForm(FormApp.openById(fvId)).onFormSubmit().create();
    ScriptApp.newTrigger('actualizarResumen').forSpreadsheet(ss).onEdit().create();
  }
  Logger.log('Listo: coordenadas corregidas y Resumen actualizado.');
}

/** Escribe en "Resumen": id | promedio | votos (sin contar las filas con Oculto = Sí). */
function actualizarResumen() {
  // Un solo cálculo a la vez: el formulario y las ediciones manuales pueden disparar esto juntos.
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) { console.warn('actualizarResumen: otra ejecución en curso, se omite'); return; }
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const val = ss.getSheetByName('Valoraciones');
    const res = ss.getSheetByName('Resumen') || ss.insertSheet('Resumen');
    const datos = {};
    let descartadas = 0;
    const rafagas = {};       // id -> tiempos de votos aceptados (ventana deslizante)
    const ignorados = {};     // id -> votos ignorados por ráfaga
    if (val && val.getLastRow() > 1) {
      const filas = val.getRange(2, 1, val.getLastRow() - 1, 6).getValues();
      filas.sort(function (a, b) { return new Date(a[0]) - new Date(b[0]); });
      filas.forEach(function (f) {
        const m = String(f[1]).match(/\(([a-z0-9-]{2,60})\)\s*$/);
        const estrellas = Number(f[2]);
        const oculto = String(f[5]).trim().toLowerCase();
        if (oculto === 'sí' || oculto === 'si') return;
        // Solo cuentan respuestas con un lugar reconocible y de 1 a 5 estrellas enteras.
        if (!m || !Number.isInteger(estrellas) || estrellas < 1 || estrellas > 5) { descartadas++; return; }
        const id = m[1];
        // Anti-relleno: por lugar cuentan como máximo RAFAGA_MAX votos en RAFAGA_MIN minutos y DIA_MAX en 24 horas.
        const ts = new Date(f[0]).getTime() || 0;
        const r = rafagas[id] || (rafagas[id] = []);
        const recientes = r.filter(function (x) { return ts - x <= RAFAGA_MIN * 60000; }).length;
        const delDia = r.filter(function (x) { return ts - x <= 86400000; }).length;
        if (recientes >= RAFAGA_MAX || delDia >= DIA_MAX) { ignorados[id] = (ignorados[id] || 0) + 1; return; }
        r.push(ts);
        while (r.length && ts - r[0] > 86400000) r.shift();
        if (!datos[id]) datos[id] = { suma: 0, n: 0 };
        datos[id].suma += estrellas;
        datos[id].n += 1;
      });
    }
    const out = [['id', 'promedio', 'votos']];
    Object.keys(datos).sort().forEach(function (id) {
      out.push([id, Math.round(datos[id].suma / datos[id].n * 10) / 10, datos[id].n]);
    });
    // Escribe primero y después limpia lo que sobra: el CSV publicado nunca queda vacío a medias.
    res.getRange(1, 1, out.length, 3).setValues(out);
    const sobra = res.getLastRow() - out.length;
    if (sobra > 0) res.getRange(out.length + 1, 1, sobra, 3).clearContent();
    if (descartadas) console.warn('actualizarResumen: ' + descartadas + ' respuestas descartadas por datos inválidos');
    escribirAlertas_(ss, ignorados);
  } catch (e) {
    console.error('actualizarResumen falló: ' + (e && e.stack || e));
    throw e;
  } finally {
    lock.releaseLock();
  }
}

// Votos que superan la ráfaga permitida: se anotan en la hoja "Alertas" para revisar.
// Si son legítimos, se pueden aceptar quitándolos de la ventana (o subiendo RAFAGA_MAX).
const RAFAGA_MAX = 3;
const RAFAGA_MIN = 10;
const DIA_MAX = 8;          // y como máximo 8 votos por lugar en 24 horas
function escribirAlertas_(ss, ignorados) {
  const ids = Object.keys(ignorados);
  const h = ss.getSheetByName('Alertas') || ss.insertSheet('Alertas');
  const out = [['Lugar', 'Votos ignorados por ráfaga', 'Revisado el']];
  ids.sort().forEach(function (id) { out.push([id, ignorados[id], new Date()]); });
  h.getRange(1, 1, out.length, 3).setValues(out);
  const sobra = h.getLastRow() - out.length;
  if (sobra > 0) h.getRange(out.length + 1, 1, sobra, 3).clearContent();
  if (ids.length) console.warn('Posible relleno de votos en: ' + ids.join(', '));
}

// ---------- Lista de lugares del formulario de valoraciones ----------
// La página publica https://bicifriendly.cl/places.json con todos sus lugares. Esta función
// copia esa lista al formulario, así el botón «Valorar» de cada lugar siempre lo encuentra.
// Si el archivo no se puede leer o viene raro, no toca el formulario.
const URL_LUGARES = 'https://bicifriendly.cl/places.json';

function sincronizarLugares() {
  const id = formId_('FORM_VALORAR_ID');
  let lista;
  try {
    const r = UrlFetchApp.fetch(URL_LUGARES + '?t=' + Date.now(), { muteHttpExceptions: true, followRedirects: true });
    if (r.getResponseCode() !== 200) throw new Error('HTTP ' + r.getResponseCode());
    lista = JSON.parse(r.getContentText());
  } catch (e) {
    console.warn('sincronizarLugares: no se pudo leer places.json (' + e + '). El formulario queda como estaba.');
    return;
  }
  if (!Array.isArray(lista) || lista.length < 5) {
    console.warn('sincronizarLugares: places.json trae ' + (lista && lista.length) + ' lugares. No se cambia el formulario.');
    return;
  }
  const opciones = lista
    .filter(function (p) { return p && /^[a-z0-9-]{2,60}$/.test(String(p.id)) && String(p.name || '').trim(); })
    .map(function (p) { return String(p.name).trim() + ' (' + String(p.id).trim() + ')'; });
  const unicas = Array.from(new Set(opciones)).sort(function (a, b) { return a.localeCompare(b, 'es'); });
  const item = FormApp.openById(id).getItems(FormApp.ItemType.LIST)[0].asListItem();
  item.setChoiceValues(unicas);
  console.log('sincronizarLugares: ' + unicas.length + ' lugares en el formulario.');
}

/** Ejecutar una vez: sincroniza ahora y deja la sincronización cada 6 horas. */
function instalarSincronizacion() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'sincronizarLugares') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sincronizarLugares').timeBased().everyHours(6).create();
  sincronizarLugares();
}
