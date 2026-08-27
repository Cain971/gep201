const SHEET_ID = '1XGk13XxKvQJzPZFI7iV1LtBFMURh6HZV1O024r_zwAA'; // GEP201_2026-2_Notas (antes apuntaba a 2026-1)

// ── HTML DE EVALUACIÓN → REGISTRO_SESIONES ────────────────────────────────
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const action = data.action;

    if (action === 'registrar_sesion') {
      const sh = ss.getSheetByName('Registro_Sesiones');
      sh.appendRow([
        data.fecha, data.sesion, data.modo, data.instrumento,
        data.codigo, data.alumno, data.grupo, data.rol,
        data.ajuste, data.nivel, data.observacion
      ]);

      // El conteo de "veces participado" se actualiza aquí (cuando la
      // participación queda efectivamente calificada), no cuando se sortea.
      // Así cuenta por igual sorteo, selección manual y voluntarios, y no
      // cuenta un sorteo que nunca llegó a calificarse.
      if (data.modo === 'Modo 1') {
        incrementarConteoParticipacion(ss, data.codigo);
      }

      return ContentService.createTextOutput(
        JSON.stringify({status:'ok'})
      ).setMimeType(ContentService.MimeType.JSON);
    }

    // ── SORTEO → SOLO REGISTRA EL SORTEO (para el traspaso a evaluación) ───
    // Llamado por sorteo_gep201.html cada vez que sale un nombre. El conteo
    // de participación NO se toca aquí — se actualiza cuando se califica.
    if (action === 'sorteo_registrar') {
      const shEstado = ss.getSheetByName('Sorteo_Estado');
      shEstado.appendRow([
        new Date(), data.codigo, data.alumno, data.grupo, data.modoDestino || '', data.sesion || ''
      ]);
      return ContentService.createTextOutput(
        JSON.stringify({status:'ok'})
      ).setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService.createTextOutput(
      JSON.stringify({status:'error', msg:'accion no reconocida'})
    ).setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService.createTextOutput(
      JSON.stringify({status:'error', msg: err.toString()})
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

// Moodle entrega nombres/apellidos en MAYÚSCULAS. Esto los pasa a formato título
// (ej. "ARIANA ROSARIO" -> "Ariana Rosario"), respetando tildes y ñ.
function properCase(s) {
  return String(s).toLowerCase().replace(/(^|\s)([a-zñáéíóúü])/g, function(m, p1, p2) {
    return p1 + p2.toUpperCase();
  });
}

// Suma 1 al conteo de participación de un alumno (Sorteo_Conteo), sin importar
// si vino de sorteo, selección manual o voluntario.
function incrementarConteoParticipacion(ss, codigo) {
  const sh = ss.getSheetByName('Sorteo_Conteo');
  const values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(codigo)) {
      const fila = i + 1;
      sh.getRange(fila, 3).setValue((Number(values[i][2]) || 0) + 1);
      sh.getRange(fila, 4).setValue(new Date());
      break;
    }
  }
}

// ── LECTURAS (GET) — usadas por sorteo_gep201.html e index.html ───────────
function doGet(e) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const action = e.parameter.action;

    // sorteo_gep201.html e index.html consultan esto al cargar la página, en vez
    // de tener la lista de alumnos escrita a mano. Para actualizarla en un
    // semestre futuro: editar solo la hoja "Nomina" (pegar el CSV de Moodle).
    if (action === 'nomina') {
      const sh = ss.getSheetByName('Nomina');
      const values = sh.getDataRange().getValues();
      const alumnos = [];
      for (let i = 1; i < values.length; i++) {
        const codigo = values[i][2];
        if (!codigo) continue;
        alumnos.push({
          codigo: String(codigo),
          nombre: properCase(values[i][0]) + ' ' + properCase(values[i][1]),
          grupo: String(values[i][4]).trim()
        });
      }
      return ContentService.createTextOutput(JSON.stringify({status:'ok', alumnos: alumnos}))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // index.html consulta esto cada pocos segundos para saber si hubo sorteos nuevos.
    // Devuelve las últimas N filas (no solo la más reciente) para que una ronda de
    // "sortear todos los grupos" (varios nombres seguidos) no se pisen entre sí.
    if (action === 'sorteo_estado') {
      const sh = ss.getSheetByName('Sorteo_Estado');
      const last = sh.getLastRow();
      if (last < 2) {
        return ContentService.createTextOutput(JSON.stringify({status:'ok', sorteos:[]}))
          .setMimeType(ContentService.MimeType.JSON);
      }
      const N = 20;
      const desde = Math.max(2, last - N + 1);
      const filas = sh.getRange(desde, 1, last - desde + 1, 6).getValues();
      const sorteos = filas.map(function(row) {
        return { timestamp: row[0], codigo: row[1], alumno: row[2], grupo: row[3], modoDestino: row[4], sesion: row[5] };
      });
      return ContentService.createTextOutput(JSON.stringify({status:'ok', sorteos: sorteos}))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // sorteo_gep201.html consulta esto al cargar la página, para ponderar la selección
    if (action === 'sorteo_conteos') {
      const sh = ss.getSheetByName('Sorteo_Conteo');
      const values = sh.getDataRange().getValues();
      const conteos = {};
      for (let i = 1; i < values.length; i++) {
        conteos[String(values[i][0])] = Number(values[i][2]) || 0;
      }
      return ContentService.createTextOutput(JSON.stringify({status:'ok', conteos: conteos}))
        .setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService.createTextOutput(
      JSON.stringify({status:'error', msg:'accion no reconocida'})
    ).setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService.createTextOutput(
      JSON.stringify({status:'error', msg: err.toString()})
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

// ── COEVALUACIÓN → COEVALUACION (sin cambios) ──────────────────────────────
const INSTRUMENTO_FORM_MAP = {
  'Coevaluación de contribución grupal — RP1': 'RP1',
  'Coevaluación de contribución grupal — RP2': 'RP2',
  'Coevaluación de contribución grupal — RP3': 'RP3',
  'Coevaluación de contribución grupal — TF':  'TF_grupal',
};

// PENDIENTE: reconstruir con los 17 alumnos de 2026-2 cuando crees los nuevos
// Forms de coevaluación (cada alumno evalúa a sus compañeros de grupo). Se
// deja vacío por ahora para que el script sea válido y desplegable; onFormSubmit
// simplemente no encontrará match y no hará nada hasta que se complete esto.
const EVALUATOR_MAP = {};

function onFormSubmit(e) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const shCoev = ss.getSheetByName('Coevaluacion');

    const response = e.response;
    const itemResponses = response.getItemResponses();

    const evaluatorKey = itemResponses[0].getResponse().toString().trim();
    const evalInfo = EVALUATOR_MAP[evaluatorKey];
    if (!evalInfo) {
      Logger.log('NO MATCH — key not found in EVALUATOR_MAP');
      return;
    }

    const timestamp = response.getTimestamp();
    const formTitle = e.source.getTitle().trim();
    let instrumento = 'RP1';
    Object.entries(INSTRUMENTO_FORM_MAP).forEach(([title, instr]) => {
      if (formTitle.includes(title.trim()) || title.trim().includes(formTitle)) {
        instrumento = instr;
      }
    });

    const scores = [];
    const evaluatees = [];
    let scoreIdx = 0;
    for (let i = 1; i < itemResponses.length; i++) {
      const val = itemResponses[i].getResponse().toString().trim();
      const num = parseFloat(val);
      if (!isNaN(num) && scoreIdx < 4) {
        scores.push(num);
        evaluatees.push(evalInfo.evaluatees[scoreIdx]);
        scoreIdx++;
      }
    }

    scores.forEach((score, j) => {
      shCoev.appendRow([timestamp, instrumento, evalInfo.name, evalInfo.group, evaluatees[j], score]);
    });

  } catch(err) {
    Logger.log('ERROR: ' + err.toString());
  }
}
