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
