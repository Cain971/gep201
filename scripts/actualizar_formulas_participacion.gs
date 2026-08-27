function actualizarFormulasParticipacion() {
  var SHEET_ID = '1XGk13XxKvQJzPZFI7iV1LtBFMURh6HZV1O024r_zwAA';
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName('Consolidacion');
  var N = 17;

  function sinDominioCheck(instrumento, r) {
    return 'COUNTIFS(Registro_Sesiones!$D:$D,"' + instrumento + '",Registro_Sesiones!$E:$E,$A' + r + ',Registro_Sesiones!$J:$J,"Sin dominio")>0';
  }

  var gFormulas = [], kFormulas = [], oFormulas = [], sFormulas = [], aaFormulas = [];

  for (var i = 0; i < N; i++) {
    var r = 4 + i;
    var acR = r - 1;

    // RP1/RP2/RP3 individual: si hubo "Sin dominio" registrado para ese alumno
    // en ese reporte, la nota individual queda fija en 10 (ignora grupal + coeval + oral).
    var baseG = 'IFERROR(IF(D' + r + '="","",MIN(20,MAX(0,IF(ISNUMBER(E' + r + '),D' + r + '+E' + r + ',D' + r + ')+F' + r + '))),"")';
    gFormulas.push(['=IF(' + sinDominioCheck('RP1', r) + ',10,' + baseG + ')']);

    var baseK = 'IFERROR(IF(H' + r + '="","",MIN(20,MAX(0,IF(ISNUMBER(I' + r + '),H' + r + '+I' + r + ',H' + r + ')+J' + r + '))),"")';
    kFormulas.push(['=IF(' + sinDominioCheck('RP2', r) + ',10,' + baseK + ')']);

    var baseO = 'IFERROR(IF(L' + r + '="","",MIN(20,MAX(0,IF(ISNUMBER(M' + r + '),L' + r + '+M' + r + ',L' + r + ')+N' + r + '))),"")';
    oFormulas.push(['=IF(' + sinDominioCheck('RP3', r) + ',10,' + baseO + ')']);

    // Control de Lectura: promedio de las 2 mejores de 3, SIN ajuste oral
    // (el CL no lleva mecanismo de participación/sorteo).
    sFormulas.push(['=IFERROR(IF(COUNTA(P' + r + ':R' + r + ')<2,"",MIN(20,MAX(0,(SUM(P' + r + ':R' + r + ')-MINIFS(P' + r + ':R' + r + ',P' + r + ':R' + r + ',"<>"))/2))),"")']);

    // Ajuste_Individual_TF = Ajuste_Coevaluación_TF + Ajuste_Participación, tope [-4,+3].
    // Ajuste_Participación = promedio redondeado de los eventos de Participación
    // (Modo 1) del semestre; si tiene menos de 2 registros, el ajuste es 0.
    var coevalTF = 'IF(Ajustes_Coeval!S' + acR + '="",0,Ajustes_Coeval!S' + acR + ')';
    var countPart = 'COUNTIFS(Registro_Sesiones!$C:$C,"Modo 1",Registro_Sesiones!$E:$E,$A' + r + ')';
    var avgPart = 'ROUND(IFERROR(AVERAGEIFS(Registro_Sesiones!$I:$I,Registro_Sesiones!$C:$C,"Modo 1",Registro_Sesiones!$E:$E,$A' + r + '),0),0)';
    var ajustePart = 'IF(' + countPart + '<2,0,' + avgPart + ')';
    aaFormulas.push(['=MAX(-4,MIN(3,(' + coevalTF + ')+(' + ajustePart + ')))']);
  }

  sh.getRange(4, 7, N, 1).setFormulas(gFormulas);   // G: RP1 Individual
  sh.getRange(4, 11, N, 1).setFormulas(kFormulas);  // K: RP2 Individual
  sh.getRange(4, 15, N, 1).setFormulas(oFormulas);  // O: RP3 Individual
  sh.getRange(4, 19, N, 1).setFormulas(sFormulas);  // S: CL Promedio 2
  sh.getRange(4, 27, N, 1).setFormulas(aaFormulas); // AA: Ajuste_Individual_TF (antes solo Coeval)

  // El conteo ya no es "solo sorteos" (cuenta toda participación) — renombrar encabezado.
  var shConteo = ss.getSheetByName('Sorteo_Conteo');
  if (shConteo) shConteo.getRange('C1').setValue('Veces_participado');

  Logger.log('Consolidacion actualizada: RP con piso 10 (Sin dominio), CL sin ajuste oral, TF_grupal = Coevaluación + Participación (tope -4/+3).');
}
