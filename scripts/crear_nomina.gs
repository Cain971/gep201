function crearNomina() {
  var SHEET_ID = '1XGk13XxKvQJzPZFI7iV1LtBFMURh6HZV1O024r_zwAA';
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var existing = ss.getSheetByName('Nomina');
  if (existing) ss.deleteSheet(existing);
  var sh = ss.insertSheet('Nomina');

  // Estas son EXACTAMENTE las columnas que trae el CSV de participantes de
  // Moodle/PAIDEIA. Para actualizar en un semestre futuro: borra las filas de
  // datos (deja la fila 1 de encabezados) y pega ahí el CSV nuevo tal cual,
  // empezando en A2 — no hace falta transformar nada.
  // Nunca escribir datos de alumnos en este archivo: el repo es público.
  var datos = [
    ['Nombre', 'Apellido(s)', 'Código PUCP', 'Dirección de correo', 'Grupos']
  ];

  sh.getRange(1, 1, datos.length, 5).setValues(datos);
  sh.getRange(1, 1, 1, 5).setFontWeight('bold');
  sh.getRange('C2:C300').setNumberFormat('@'); // Código como texto

  // Dejarla justo después de INSTRUCCIONES
  ss.setActiveSheet(sh);
  ss.moveActiveSheet(2);

  Logger.log('Nomina creada (solo encabezados): pega el CSV de PAIDEIA en A2.');
}
