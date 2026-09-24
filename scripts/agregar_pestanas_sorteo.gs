function agregarPestanasSorteo() {
  var SHEET_ID = '1XGk13XxKvQJzPZFI7iV1LtBFMURh6HZV1O024r_zwAA';
  var ss = SpreadsheetApp.openById(SHEET_ID);

  // Nómina leída de la pestaña Nomina (CSV de PAIDEIA). Nunca escribir datos
  // de alumnos en este archivo: el repo es público.
  // Formato: [código, 'Apellidos, Nombres', 'GRUPO N'], ordenado por grupo y nombre
  // (el mismo orden y formato que usan los Forms de coevaluación).
  var ROSTER = (function() {
    function pc(x) {
      return String(x).toLowerCase().replace(/(^|\s)([a-zñáéíóúü])/g, function(m, a, b) { return a + b.toUpperCase(); });
    }
    return ss.getSheetByName('Nomina').getDataRange().getValues().slice(1)
      .filter(function(r) { return r[2]; })
      .map(function(r) { return [String(r[2]), pc(r[1]) + ', ' + pc(r[0]), String(r[4]).trim()]; })
      .sort(function(a, b) { return a[2] !== b[2] ? (a[2] < b[2] ? -1 : 1) : (a[1] < b[1] ? -1 : a[1] > b[1] ? 1 : 0); });
  })();

  function newSheet(name) {
    var existing = ss.getSheetByName(name);
    if (existing) ss.deleteSheet(existing);
    return ss.insertSheet(name);
  }

  // ---------- Sorteo_Estado: log de cada sorteo (para el traspaso sorteo -> evaluacion) ----------
  var sh = newSheet('Sorteo_Estado');
  sh.getRange(1, 1, 1, 6).setValues([[
    'Timestamp', 'Código', 'Alumno', 'Grupo', 'Modo destino', 'Sesión'
  ]]).setFontWeight('bold');
  sh.getRange('A2').setValue('⚠ Escrita automáticamente por sorteo_gep201.html. index.html lee la última fila. No editar.');

  // ---------- Sorteo_Conteo: una fila por alumno, cuantas veces le ha tocado en el semestre ----------
  sh = newSheet('Sorteo_Conteo');
  sh.getRange(1, 1, 1, 4).setValues([[
    'Código', 'Alumno', 'Veces_sorteado', 'Última_fecha'
  ]]).setFontWeight('bold');
  var rows = ROSTER.map(function(s) { return [s[0], s[1], 0, '']; });
  sh.getRange(2, 1, rows.length, 4).setValues(rows);
  sh.getRange(2, 1, rows.length, 1).setNumberFormat('@');

  // Reordenar: dejar estas dos pestañas justo después de Registro_Sesiones
  ss.setActiveSheet(ss.getSheetByName('Sorteo_Estado'));
  ss.moveActiveSheet(3);
  ss.setActiveSheet(ss.getSheetByName('Sorteo_Conteo'));
  ss.moveActiveSheet(4);

  Logger.log('Listo: Sorteo_Estado y Sorteo_Conteo creadas con ' + rows.length + ' alumnos.');
}
