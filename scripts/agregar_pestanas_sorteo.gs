function agregarPestanasSorteo() {
  var SHEET_ID = '1XGk13XxKvQJzPZFI7iV1LtBFMURh6HZV1O024r_zwAA';
  var ss = SpreadsheetApp.openById(SHEET_ID);

  var ROSTER = [
    ['20214097','Becerra Sulluchuco, Milene Yadira','GRUPO 1'],
    ['20224841','Bueno Hurtado, Sebastian Fabio','GRUPO 1'],
    ['20216217','Gonzales Benites, Ariana Rosario','GRUPO 1'],
    ['20211479','Rosales Aguilar, Luciana Concepcion','GRUPO 1'],
    ['20226646','Villanueva Alvarez, Daria Luciana','GRUPO 1'],
    ['20160993','Baca Montes, Valeria Carolina','GRUPO 2'],
    ['20191529','Calderon Sipan, Eduardo Franco','GRUPO 2'],
    ['20191026','Cardenas Torres, Carla Mayerling','GRUPO 2'],
    ['20213311','Moreau Tejada, Alejandra Lily','GRUPO 2'],
    ['20241271','Ocampo Huayllas, Pablo Jesus','GRUPO 2'],
    ['20163080','Valverde Cipriano, Joseph Jack','GRUPO 2'],
    ['20213680','Campos Vasquez, Patricia Nubia','GRUPO 3'],
    ['20210437','Grimaldo Vicente, Steve','GRUPO 3'],
    ['20205930','Huasasquiche Uculmana, Maria Fernanda','GRUPO 3'],
    ['20221937','Ordonez Acevedo, Cesar Andre','GRUPO 3'],
    ['20200613','Pacheco Flores, Barbara Nicole','GRUPO 3'],
    ['20220858','Villar Razzeto, Maricielo','GRUPO 3']
  ];

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
