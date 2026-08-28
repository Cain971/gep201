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
          nombres: properCase(values[i][0]),   // solo nombres (para el corte nombre/apellidos del sorteo)
          apellidos: properCase(values[i][1]), // solo apellidos
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

// ── COEVALUACIÓN → hoja "Coevaluacion" ────────────────────────────────────
// Los 4 Forms (RP1/RP2/RP3/TF) se generan con construir_forms_coevaluacion.gs.
// Estructura de cada Form:
//   Q1  "¿Cuál es tu nombre?"  — opción "GRUPO N: Apellidos, Nombres", ramifica
//        a una sección por evaluador.
//   Sección del evaluador — una casilla numérica 0-100 por cada compañero de su
//        grupo; el TÍTULO de esa casilla ES el nombre del evaluado
//        ("Apellidos, Nombres"). El evaluado se identifica por título, no por
//        posición (así funciona con grupos de 5 o de 6, sin tope fijo).
//
// Fila escrita en Coevaluacion: [Timestamp, Instrumento, Evaluador, Grupo,
// Evaluado, Puntaje]. "Evaluado" y "Grupo" salen exactamente en el formato que
// Ajustes_Coeval busca con AVERAGEIFS ("Apellidos, Nombres" / "GRUPO N").
// Instrumento ∈ {RP1, RP2, RP3, TF_grupal}.

const COEVAL_Q1_TITULO = '¿Cuál es tu nombre?';

// código → nombre "Apellidos, Nombres" + grupo, leído de la pestaña Nomina.
function coevalRoster_(ss) {
  const values = ss.getSheetByName('Nomina').getDataRange().getValues();
  const porNombre = {};
  for (let i = 1; i < values.length; i++) {
    if (!values[i][2]) continue;
    const nombre = properCase(values[i][1]) + ', ' + properCase(values[i][0]);
    porNombre[nombre] = { nombre: nombre, grupo: String(values[i][4]).trim() };
  }
  return porNombre;
}

// Deriva el instrumento del título del Form.
function coevalInstrumento_(titulo) {
  const t = String(titulo);
  if (/\bTF\b/.test(t) || /trabajo\s+final/i.test(t)) return 'TF_grupal';
  const m = t.match(/\bRP\s?([123])\b/);
  return m ? 'RP' + m[1] : null;
}

function onFormSubmit(e) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const shCoev = ss.getSheetByName('Coevaluacion');
    const roster = coevalRoster_(ss);

    const instrumento = coevalInstrumento_(e.source.getTitle());
    if (!instrumento) {
      Logger.log('Coeval: instrumento no derivable de "' + e.source.getTitle() + '"');
      return;
    }

    const respuestas = e.response.getItemResponses();
    const timestamp = e.response.getTimestamp();

    // Q1 → evaluador y grupo. Se busca por título (no por índice) por si la
    // recolección de correo u otro item se cuela primero.
    let q1 = '';
    for (let i = 0; i < respuestas.length; i++) {
      if (String(respuestas[i].getItem().getTitle()).trim() === COEVAL_Q1_TITULO) {
        q1 = String(respuestas[i].getResponse()).trim();
        break;
      }
    }
    const sep = q1.indexOf(': ');
    const evaluador = sep >= 0 ? q1.slice(sep + 2).trim() : q1;
    const grupo = (roster[evaluador] && roster[evaluador].grupo) ||
                  (sep >= 0 ? q1.slice(0, sep).trim() : '');
    if (!evaluador) { Logger.log('Coeval: Q1 vacía o no encontrada'); return; }

    const filas = [];
    for (let i = 0; i < respuestas.length; i++) {
      const titulo = String(respuestas[i].getItem().getTitle()).trim();
      if (!roster[titulo]) continue;                 // no es casilla de puntaje
      const puntaje = parseFloat(String(respuestas[i].getResponse()).trim());
      if (isNaN(puntaje)) continue;
      filas.push([timestamp, instrumento, evaluador, grupo, titulo, puntaje]);
    }

    if (filas.length) {
      shCoev.getRange(shCoev.getLastRow() + 1, 1, filas.length, 6).setValues(filas);
    } else {
      Logger.log('Coeval: 0 puntajes reconocidos — evaluador "' + evaluador + '" (' + instrumento + ')');
    }
  } catch (err) {
    Logger.log('ERROR onFormSubmit: ' + err.toString());
  }
}
