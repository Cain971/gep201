function limpiarPresentacionManualObsoleta() {
  var SHEET_ID = '1XGk13XxKvQJzPZFI7iV1LtBFMURh6HZV1O024r_zwAA';
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var N = 17;

  // 1) Borrar la columna TF_presentacion de Notas_Manuales (columna K, la 11).
  //    TF_grupal, que estaba en L, pasa a ocupar la columna K automáticamente.
  var shNM = ss.getSheetByName('Notas_Manuales');
  shNM.deleteColumn(11);

  // 2) Consolidacion, columna Z (TF_grupal base) leía Notas_Manuales!L — ahora
  //    debe leer Notas_Manuales!K, que es donde quedó TF_grupal tras el corrimiento.
  var shC = ss.getSheetByName('Consolidacion');
  var zFormulas = [];
  for (var i = 0; i < N; i++) {
    var r = 4 + i;
    zFormulas.push(['=IFERROR(Notas_Manuales!K' + r + ',"")']);
  }
  shC.getRange(4, 26, N, 1).setFormulas(zFormulas); // Z

  // 3) Las columnas W y X de Consolidacion (base manual y "ajuste oral" de la
  //    vieja Presentación) ya no las usa nada — Y ahora es autosuficiente. Se limpian.
  shC.getRange(4, 23, N, 2).clearContent(); // W, X

  Logger.log('Listo: Notas_Manuales sin la columna TF_presentacion; Consolidacion realineada (Z lee K); W/X limpiadas.');
}
