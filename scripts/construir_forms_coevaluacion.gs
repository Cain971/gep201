/**
 * construir_forms_coevaluacion.gs — GEP201
 *
 * Genera los 4 Google Forms de coevaluación (RP1, RP2, RP3, TF) a partir de la
 * pestaña "Nomina" del Sheet de notas, instala el trigger onFormSubmit de ESTE
 * proyecto en cada uno y los mueve a la carpeta de Drive del curso.
 *
 * USO:  ejecutar construirFormsCoevaluacion()  VARIAS VECES.
 *   Cada corrida construye UN solo Form (el siguiente que falte) — así ninguna
 *   corrida se pasa del límite de 6 minutos de Apps Script. Además, en cada
 *   corrida verifica/repara el trigger de los Forms ya creados.
 *   El log dice cuántos faltan. Repetir hasta que diga "LISTO".
 *
 *   Si una corrida se corta a media construcción de un Form, ese Form queda
 *   en la RAÍZ de "Mi unidad" (no en la carpeta). Borrarlo a mano y volver a
 *   ejecutar.
 *
 * Regenerar en un semestre nuevo:
 *   1. Actualizar la pestaña Nomina (pegar el CSV nuevo de Moodle).
 *   2. Renombrar/archivar los Forms del semestre anterior.
 *   3. Ejecutar construirFormsCoevaluacion() las veces que haga falta.
 *
 * Depende de:  SHEET_ID, properCase() y COEVAL_Q1_TITULO definidos en Code.gs,
 * y de onFormSubmit() en el mismo proyecto.
 */

// Subcarpeta "GEP201 2026-2" en Drive (dentro de la carpeta personal "GEP201").
// En un semestre nuevo: crear "GEP201 20XX-Y" y poner aquí su id.
const COEVAL_FOLDER_ID = '1Ba5nB9av4-jJ8kMbUTaoDl1uT2jVdm0X';

const COEVAL_INSTRUMENTOS = [
  { tag: 'RP1', label: 'Reporte Parcial 1 (RP1)' },
  { tag: 'RP2', label: 'Reporte Parcial 2 (RP2)' },
  { tag: 'RP3', label: 'Reporte Parcial 3 (RP3)' },
  { tag: 'TF',  label: 'Trabajo Final (TF)' },
];

function coevalTitulo_(instr) {
  return 'Coevaluación de contribución grupal — ' + instr.label;
}

function coevalPortada_(label) {
  return [
    'Coevaluación de contribución grupal — ' + label,
    'GEP201 · Gestión Financiera del Estado',
    '',
    'Esta evaluación es anónima y forma parte del sistema de calificación individual del curso. Tu nombre no será visible para tus compañeros ni aparecerá asociado a tus respuestas en ningún reporte.',
    '',
    '¿Para qué sirve?',
    'La nota que el profesor asigna al trabajo grupal se ajusta individualmente según la contribución real de cada integrante. Esta evaluación es el mecanismo para hacer esa distinción de forma justa y transparente.',
    '',
    'Reglas importantes:',
    '• Asigna un puntaje distinto a cada compañero — no se permiten puntajes repetidos.',
    '• Evalúa solo a tus compañeros, no a ti mismo.',
    '• Tienes 48 horas desde la fecha de entrega para completar este formulario.',
    '• No completar el formulario en el plazo establecido tendrá consecuencias en tu propia calificación individual.',
  ].join('\n');
}

// Guía orientativa — visible en la pantalla donde el alumno asigna los puntajes.
const COEVAL_GUIA_SECCION = [
  'Asigna un puntaje de 0 a 100 a cada compañero según su contribución real. Usa un puntaje DISTINTO para cada persona.',
  '',
  '80–100 · Contribución excepcional — lideró el trabajo, entregó a tiempo, mejoró la calidad del resultado grupal',
  '60–79 · Contribución sólida — cumplió su parte con calidad y dentro de los plazos',
  '40–59 · Contribución irregular — cumplió parcialmente o con retrasos que afectaron al grupo',
  '0–39 · Contribución mínima o ausente — no cumplió con su parte o su aporte fue marginal',
].join('\n');

const COEVAL_DECLARACION_HELP =
  'Al enviar confirmo que: (1) asigné un puntaje distinto a cada compañero; ' +
  '(2) los puntajes reflejan honestamente la contribución real de cada persona; ' +
  '(3) entiendo que las evaluaciones falsas o coordinadas pueden ser detectadas y afectar mi propia calificación.';


// Roster desde Nomina, ordenado por grupo y luego por nombre (igual que el
// dropdown de 2026-1). Nombre en formato "Apellidos, Nombres".
function coevalRosterOrdenado_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const values = ss.getSheetByName('Nomina').getDataRange().getValues();
  const roster = [];
  for (let i = 1; i < values.length; i++) {
    if (!values[i][2]) continue;
    roster.push({
      nombre: properCase(values[i][1]) + ', ' + properCase(values[i][0]),
      grupo: String(values[i][4]).trim(),
    });
  }
  roster.sort(function(a, b) {
    if (a.grupo !== b.grupo) return a.grupo < b.grupo ? -1 : 1;
    return a.nombre < b.nombre ? -1 : (a.nombre > b.nombre ? 1 : 0);
  });
  return roster;
}

// Busca el Form de un instrumento dentro de la carpeta del curso.
function coevalBuscarForm_(folder, instr) {
  const it = folder.getFilesByName(coevalTitulo_(instr));
  return it.hasNext() ? it.next() : null;
}

// Garantiza que exista exactamente un trigger onFormSubmit para ese formId.
function coevalAsegurarTrigger_(formId) {
  const existe = ScriptApp.getProjectTriggers().some(function (t) {
    return t.getHandlerFunction() === 'onFormSubmit' && t.getTriggerSourceId() === formId;
  });
  if (existe) return 'trigger ok';
  ScriptApp.newTrigger('onFormSubmit').forForm(FormApp.openById(formId)).onFormSubmit().create();
  return 'trigger CREADO';
}

// Construye un Form completo y lo mueve a la carpeta del curso.
function coevalConstruirUnForm_(instr, roster, folder) {
  const numValidation = FormApp.createTextValidation()
    .setHelpText('Debe ser un número entre 0 y 100.')
    .requireNumberBetween(0, 100)
    .build();

  const form = FormApp.create(coevalTitulo_(instr));
  form.setTitle(coevalTitulo_(instr));   // título interno del Form (lo lee onFormSubmit)
  form.setDescription(coevalPortada_(instr.label));
  form.setProgressBar(true);
  try { form.setShowLinkToRespondAgain(false); } catch (x) {}
  try { form.setRequireLogin(true); } catch (x) {}
  try { form.setLimitOneResponsePerUser(true); } catch (x) {}
  try { form.setAllowResponseEdits(true); } catch (x) {}
  try {
    form.setEmailCollectionType(FormApp.EmailCollectionType.VERIFIED);
  } catch (x) {
    try { form.setCollectEmail(true); } catch (y) {}
  }

  // Q1 primero. Las opciones se agregan al final, cuando ya existen las
  // referencias a cada sección.
  const q1 = form.addListItem().setTitle(COEVAL_Q1_TITULO).setRequired(true);

  // Una sección (page break) por evaluador. Nunca ve su propio nombre.
  const seccionPorNombre = {};
  roster.forEach(function (ev, idx) {
    const pb = form.addPageBreakItem()
      .setTitle(ev.grupo)
      .setHelpText(COEVAL_GUIA_SECCION);

    // Al completar la sección anterior por avance lineal, saltar a "Enviar"
    // en vez de caer en la sección del siguiente evaluador.
    if (idx > 0) pb.setGoToPage(FormApp.PageNavigationType.SUBMIT);

    roster
      .filter(function (o) { return o.grupo === ev.grupo && o.nombre !== ev.nombre; })
      .forEach(function (mate) {
        form.addTextItem()
          .setTitle(mate.nombre)          // el título ES el nombre del evaluado
          .setRequired(true)
          .setValidation(numValidation);
      });

    form.addCheckboxItem()
      .setTitle('Declaración de honestidad')
      .setHelpText(COEVAL_DECLARACION_HELP)
      .setChoiceValues(['Confirmo'])
      .setRequired(true);

    seccionPorNombre[ev.nombre] = pb;
  });

  // Opciones de Q1: "GRUPO N: Apellidos, Nombres" → sección del evaluador.
  q1.setChoices(roster.map(function (ev) {
    return q1.createChoice(ev.grupo + ': ' + ev.nombre, seccionPorNombre[ev.nombre]);
  }));

  // Mover de la raíz de Drive a la carpeta del curso (recién ahora: si algo
  // falla antes, el Form queda en la raíz y es fácil de detectar y borrar).
  const file = DriveApp.getFileById(form.getId());
  folder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);
}


function construirFormsCoevaluacion() {
  const roster = coevalRosterOrdenado_();
  if (roster.length < 2) throw new Error('Nomina vacía o con menos de 2 alumnos.');
  const folder = DriveApp.getFolderById(COEVAL_FOLDER_ID);

  // 1. Construir el SIGUIENTE Form que falte (solo uno por corrida).
  let construido = null;
  for (let i = 0; i < COEVAL_INSTRUMENTOS.length; i++) {
    if (coevalBuscarForm_(folder, COEVAL_INSTRUMENTOS[i])) continue;
    coevalConstruirUnForm_(COEVAL_INSTRUMENTOS[i], roster, folder);
    construido = COEVAL_INSTRUMENTOS[i].tag;
    break;
  }

  // 2. Asegurar el trigger onFormSubmit en cada Form ya existente.
  const estado = [];
  let faltan = 0;
  COEVAL_INSTRUMENTOS.forEach(function (instr) {
    const file = coevalBuscarForm_(folder, instr);
    if (!file) { estado.push(instr.tag + ': FALTA'); faltan++; return; }
    const f = FormApp.openById(file.getId());
    if (f.getTitle() !== coevalTitulo_(instr)) f.setTitle(coevalTitulo_(instr));  // reparar título
    estado.push(instr.tag + ': ' + coevalAsegurarTrigger_(file.getId()) +
      '  ·  título="' + f.getTitle() + '"' +
      '  ·  responder: ' + f.getPublishedUrl());
  });

  Logger.log(
    (construido ? ('Form construido en esta corrida: ' + construido) : 'Ningún Form nuevo en esta corrida.') +
    '\n\n' + estado.join('\n') + '\n\n' +
    (faltan ? ('>>> Faltan ' + faltan + ' Form(s). Vuelve a ejecutar construirFormsCoevaluacion().')
            : '>>> LISTO: los 4 Forms existen y tienen trigger onFormSubmit.')
  );
}


/**
 * Utilitario: lista los triggers onFormSubmit instalados por este proyecto.
 */
function coevalListarTriggers() {
  const t = ScriptApp.getProjectTriggers()
    .filter(function (x) { return x.getHandlerFunction() === 'onFormSubmit'; });
  Logger.log('Triggers onFormSubmit: ' + t.length);
  t.forEach(function (x) { Logger.log('  ' + x.getUniqueId() + '  source=' + x.getTriggerSourceId()); });
}
