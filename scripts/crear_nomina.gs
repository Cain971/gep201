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
  var datos = [
    ['Nombre', 'Apellido(s)', 'Código PUCP', 'Dirección de correo', 'Grupos'],
    ['VALERIA CAROLINA', 'BACA MONTES', 20160993, 'carolina.bacam@pucp.edu.pe', 'GRUPO 2'],
    ['MILENE YADIRA', 'BECERRA SULLUCHUCO', 20214097, 'milene.becerra@pucp.edu.pe', 'GRUPO 1'],
    ['SEBASTIAN FABIO', 'BUENO HURTADO', 20224841, 'a20224841@pucp.edu.pe', 'GRUPO 1'],
    ['EDUARDO FRANCO', 'CALDERON SIPAN', 20191529, 'calderon.eduardo@pucp.edu.pe', 'GRUPO 2'],
    ['PATRICIA NUBIA', 'CAMPOS VASQUEZ', 20213680, 'n.campos@pucp.edu.pe', 'GRUPO 3'],
    ['CARLA MAYERLING', 'CARDENAS TORRES', 20191026, 'carla.cardenas@pucp.edu.pe', 'GRUPO 2'],
    ['ARIANA ROSARIO', 'GONZALES BENITES', 20216217, 'ariana.gonzalesb@pucp.edu.pe', 'GRUPO 1'],
    ['STEVE', 'GRIMALDO VICENTE', 20210437, 's.grimaldo@pucp.edu.pe', 'GRUPO 3'],
    ['MARIA FERNANDA', 'HUASASQUICHE UCULMANA', 20205930, 'a20205930@pucp.edu.pe', 'GRUPO 3'],
    ['ALEJANDRA LILY', 'MOREAU TEJADA', 20213311, 'amoreaut@pucp.edu.pe', 'GRUPO 2'],
    ['PABLO JESUS', 'OCAMPO HUAYLLAS', 20241271, 'a20241271@pucp.edu.pe', 'GRUPO 2'],
    ['CESAR ANDRE', 'ORDOÑEZ ACEVEDO', 20221937, 'ordonez.cesar@pucp.edu.pe', 'GRUPO 3'],
    ['BARBARA NICOLE', 'PACHECO FLORES', 20200613, 'barbara.pacheco@pucp.edu.pe', 'GRUPO 3'],
    ['LUCIANA CONCEPCION', 'ROSALES AGUILAR', 20211479, 'a20211479@pucp.edu.pe', 'GRUPO 1'],
    ['JOSEPH JACK', 'VALVERDE CIPRIANO', 20163080, 'jvalverdec@pucp.edu.pe', 'GRUPO 2'],
    ['DARIA LUCIANA', 'VILLANUEVA ALVAREZ', 20226646, 'a20226646@pucp.edu.pe', 'GRUPO 1'],
    ['MARICIELO', 'VILLAR RAZZETO', 20220858, 'villar.maricielo@pucp.edu.pe', 'GRUPO 3']
  ];

  sh.getRange(1, 1, datos.length, 5).setValues(datos);
  sh.getRange(1, 1, 1, 5).setFontWeight('bold');
  sh.getRange(2, 3, datos.length - 1, 1).setNumberFormat('@'); // Código como texto

  // Dejarla justo después de INSTRUCCIONES
  ss.setActiveSheet(sh);
  ss.moveActiveSheet(2);

  Logger.log('Nomina creada con ' + (datos.length - 1) + ' alumnos.');
}
