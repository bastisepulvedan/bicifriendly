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
  if (props.getProperty('FORM_APORTES_ID')) {
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

/** Refresca la lista "Lugar" del formulario de valoración con los lugares publicados. */
function actualizarLugares() {
  const id = PropertiesService.getScriptProperties().getProperty('FORM_VALORAR_ID');
  if (!id) throw new Error('Primero ejecuta crearFormularios.');
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
  const fvId = props.getProperty('FORM_VALORAR_ID');
  const existentes = ScriptApp.getProjectTriggers().map(function (t) { return t.getHandlerFunction(); });
  if (fvId && existentes.indexOf('actualizarResumen') === -1) {
    ScriptApp.newTrigger('actualizarResumen').forForm(FormApp.openById(fvId)).onFormSubmit().create();
    ScriptApp.newTrigger('actualizarResumen').forSpreadsheet(ss).onEdit().create();
  }
  Logger.log('Listo: coordenadas corregidas y Resumen actualizado.');
}

/** Escribe en "Resumen": id | promedio | votos (sin contar las filas con Oculto = Sí). */
function actualizarResumen() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const val = ss.getSheetByName('Valoraciones');
  const res = ss.getSheetByName('Resumen') || ss.insertSheet('Resumen');
  const datos = {};
  if (val && val.getLastRow() > 1) {
    const filas = val.getRange(2, 1, val.getLastRow() - 1, 6).getValues();
    filas.forEach(function (f) {
      const m = String(f[1]).match(/\(([^()]+)\)\s*$/);
      const estrellas = Number(f[2]);
      const oculto = String(f[5]).trim().toLowerCase();
      if (!m || !estrellas || oculto === 'sí' || oculto === 'si') return;
      const id = m[1].trim();
      if (!datos[id]) datos[id] = { suma: 0, n: 0 };
      datos[id].suma += estrellas;
      datos[id].n += 1;
    });
  }
  const out = [['id', 'promedio', 'votos']];
  Object.keys(datos).sort().forEach(function (id) {
    out.push([id, Math.round(datos[id].suma / datos[id].n * 10) / 10, datos[id].n]);
  });
  res.clear();
  res.getRange(1, 1, out.length, 3).setValues(out);
}
