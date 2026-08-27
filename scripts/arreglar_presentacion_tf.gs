function arreglarPresentacionIndividual() {
  var SHEET_ID = '1XGk13XxKvQJzPZFI7iV1LtBFMURh6HZV1O024r_zwAA';
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName('Consolidacion');
  var N = 17;

  var yFormulas = [];
  for (var i = 0; i < N; i++) {
    var r = 4 + i;
    var cond = '(Registro_Sesiones!$C:$C="Modo 3")*(Registro_Sesiones!$E:$E=$A' + r + ')';
    var valor = 'SUMPRODUCT(' + cond + '*ISNUMBER(Registro_Sesiones!$I:$I)*IF(ISNUMBER(Registro_Sesiones!$I:$I),Registro_Sesiones!$I:$I,0))';
    var hay = 'SUMPRODUCT(' + cond + ')';
    // Si no hay ninguna fila de Modo 3 para este alumno todavía, queda en blanco
    // (pendiente de sustentación). Si ya la tiene, toma esa nota directo — ya no
    // depende de Notas_Manuales!K (esa columna queda sin uso).
    yFormulas.push(['=IFERROR(IF(' + hay + '=0,"",' + valor + '),"")']);
  }

  sh.getRange(4, 25, N, 1).setFormulas(yFormulas); // Y: TF Presentación Individual

  Logger.log('Presentación individual del TF ahora se toma directo de Modo 3 en Registro_Sesiones (columna Notas_Manuales K queda sin uso).');
}
