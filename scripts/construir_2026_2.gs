function construirHoja2026_2() {
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
  var N = ROSTER.length;

  function newSheet(name) {
    var existing = ss.getSheetByName(name);
    if (existing) ss.deleteSheet(existing);
    return ss.insertSheet(name);
  }

  // ---------- INSTRUCCIONES ----------
  var sh = newSheet('INSTRUCCIONES');
  var instrRows = [
    [2, 'GEP201 · Gestión Financiera del Estado · 2026-2', ''],
    [3, 'Sistema de Registro y Consolidación de Notas', ''],
    [5, 'ESTRUCTURA DE HOJAS', ''],
    [6, 'Registro_Sesiones', 'Recibe automáticamente los ajustes orales desde el HTML. NO editar.'],
    [7, 'Notas_Manuales', 'Ingresa aquí las notas grupales de RP1-3, notas de CL1-3, memo individual, presentación TF y TF grupal.'],
    [8, 'Coevaluacion', 'Recibe respuestas del Google Form automáticamente. NO editar.'],
    [9, 'Ajustes_Coeval', 'Calcula automáticamente desviación y nivel de coevaluación por alumno. Solo lectura.'],
    [10, 'Consolidacion', 'Integra todas las fuentes y calcula la nota final del curso. Solo lectura.'],
    [11, 'Vista_PAIDEIA', 'Lista limpia código + nombre + nota final para subir al sistema institucional.'],
    [13, 'FÓRMULA DEL CURSO', ''],
    [14, 'Reportes Parciales (30%)', 'RP1 + RP2 + RP3 · 10% cada uno · Nota grupal ± coevaluación ± ajuste oral'],
    [15, 'Controles de Lectura (20%)', 'CL1, CL2, CL3 · Se elimina la menor · Las dos restantes pesan 10% c/u · ± ajuste oral'],
    [16, 'Trabajo Final (50%)', 'Memo individual 20% + Presentación individual 20% + TF grupal individual 10%'],
    [18, 'REGLAS DE CÁLCULO', ''],
    [19, 'Coevaluación', 'Desviación = puntaje recibido − promedio del grupo. Sobresaliente ≥+10 → +2 | Normal −5 a +9 → 0 | Baja −15 a −6 → −2 | Free rider ≤−16 → −4'],
    [20, 'Ajuste oral', 'Registrado por el HTML. Puede aplicar a cualquier instrumento. Se suma a la nota individual después de coevaluación.'],
    [21, 'No entrega coeval', 'Penalización automática de −2 en el instrumento correspondiente.'],
    [22, 'Escala', '0–20. Aprobación con 11. Fracción ≥0.5 sube al entero superior (Reglamento FGAD Art. 45).']
  ];
  instrRows.forEach(function(r) {
    sh.getRange(r[0], 2).setValue(r[1]);
    if (r[2]) sh.getRange(r[0], 3).setValue(r[2]);
  });
  sh.getRange('B2').setFontWeight('bold').setFontSize(13);
  sh.getRange('B5').setFontWeight('bold');
  sh.getRange('B13').setFontWeight('bold');
  sh.getRange('B18').setFontWeight('bold');

  // ---------- Registro_Sesiones ----------
  sh = newSheet('Registro_Sesiones');
  sh.getRange(1, 1, 1, 11).setValues([[
    'Fecha','Sesión','Modo','Instrumento','Código PUCP','Alumno','Grupo','Rol','Ajuste','Nivel','Observación'
  ]]).setFontWeight('bold');
  sh.getRange('A2').setValue('⚠ Esta hoja es escritura automática del HTML de evaluación. No editar manualmente.');

  // ---------- Notas_Manuales ----------
  sh = newSheet('Notas_Manuales');
  sh.getRange('A1').setValue('NOTAS MANUALES · Ingresa solo las celdas azules. Las demás hojas se calculan automáticamente.').setFontWeight('bold');
  sh.getRange('D2').setValue('REPORTES PARCIALES (nota grupal base)');
  sh.getRange('G2').setValue('CONTROLES DE LECTURA (nota individual)');
  sh.getRange('J2').setValue('TRABAJO FINAL');
  sh.getRange(3, 1, 1, 12).setValues([[
    'Código','Alumno','Grupo','RP1_grupal','RP2_grupal','RP3_grupal','CL1','CL2','CL3','TF_memo','TF_presentacion','TF_grupal'
  ]]).setFontWeight('bold');
  var nmRows = ROSTER.map(function(s){ return [s[0], s[1], s[2], '', '', '', '', '', '', '', '', '']; });
  sh.getRange(4, 1, N, 12).setValues(nmRows);
  sh.getRange(4, 1, N, 1).setNumberFormat('@'); // codigo como texto

  // ---------- Coevaluacion ----------
  sh = newSheet('Coevaluacion');
  sh.getRange(1, 1, 1, 6).setValues([[
    'Timestamp','Instrumento','Evaluador','Grupo','Evaluado','Puntaje'
  ]]).setFontWeight('bold');
  sh.getRange('A2').setValue('⚠ Esta hoja recibe respuestas del Google Form automáticamente. No editar manualmente.');

  // ---------- Metodología_Coeval ----------
  sh = newSheet('Metodología_Coeval');
  var meto = [
    [1, 'METODOLOGÍA DEL AJUSTE POR COEVALUACIÓN — GEP201'],
    [2, 'Hoja de referencia — no editar durante el semestre'],
    [4, '1. Qué corrige este ajuste'],
    [5, 'La nota grupal de cada entregable (RP1, RP2, RP3, TF grupal) no distingue el aporte individual dentro del equipo. Este ajuste corrige la nota grupal según cómo cada estudiante fue evaluado por sus propios compañeros, para que la nota individual refleje contribución real y no solo pertenencia al grupo.'],
    [7, '2. Cómo se calcula la desviación'],
    [8, 'Desviación = puntaje recibido por el estudiante (0-100) − promedio de los puntajes que recibieron todos los integrantes de su grupo en ese instrumento'],
    [9, 'Se usa la posición relativa dentro del propio grupo, no el puntaje absoluto, porque distintos evaluadores puntúan con "temperaturas" distintas (unos generosos, otros exigentes). Comparar a cada estudiante contra el promedio de su propio grupo neutraliza ese sesgo de escala.'],
    [11, '3. Tabla de equivalencias: desviación → ajuste'],
    [18, '4. Por qué estos valores de desviación (y no otros)'],
    [19, 'Banda "Normal" amplia y centrada ligeramente bajo cero: la regla de no repetir puntajes obliga a diferenciar incluso en grupos donde todos trabajaron parejo. La banda normal absorbe esa dispersión mínima "de ruido" sin penalizar a nadie injustamente.'],
    [20, 'Asimetría entre premio y castigo: el tramo positivo exige una desviación mayor (+10) que el negativo (−6) para activarse, y el castigo máximo (−4) duplica al premio máximo (+2). El objetivo central es disuadir el free riding, no premiar generosamente el desempeño sobresaliente, que ya se refleja en otros componentes de la nota (TF individual, sustentación oral).'],
    [21, 'Corte de "free rider" en −16: calibrado para activarse solo con una desviación clara y consistente, evitando falsos positivos por un solo evaluador atípico, ya que el ajuste se basa en el promedio de 3-4 evaluadores por compañero.'],
    [23, '5. Penalización por no completar el formulario'],
    [24, 'No enviar la coevaluación dentro del plazo (48 horas) equivale automáticamente a nivel "Baja contribución" (−2) en ese instrumento, sin importar los puntajes recibidos. Esto evita que un estudiante evite dar una evaluación honesta (que hundiría a un compañero) simplemente absteniéndose sin costo.'],
    [26, '6. Limitaciones a tener en cuenta'],
    [27, 'Con grupos de 5 a 6, cada estudiante recibe entre 4 y 5 evaluaciones — muestra pequeña; un voto atípico por conflicto personal puede mover el promedio del grupo de forma no despreciable.'],
    [28, 'Los umbrales (±10, ±6, ±16) son un diseño razonado, no validados aún con datos reales del curso. Revisar con los datos acumulados de RP1-RP3 al cierre del semestre.']
  ];
  meto.forEach(function(r) { sh.getRange(r[0], 2).setValue(r[1]); });
  sh.getRange('B1').setFontWeight('bold').setFontSize(13);
  sh.getRange(12, 2, 1, 4).setValues([['Desviación respecto al promedio del grupo','Nivel','Ajuste','Qué significa']]).setFontWeight('bold');
  sh.getRange(13, 2, 4, 4).setValues([
    ['≥ +10', 'Sobresaliente', '+2', 'Contribución claramente por encima del resto del equipo'],
    ['Entre −5 y +9', 'Normal', '0', 'Dentro del rango esperado de variación entre compañeros'],
    ['Entre −15 y −6', 'Baja contribución', '−2', 'Aporte notoriamente menor al del resto del equipo'],
    ['≤ −16', 'Free rider', '−4', 'Aporte mínimo o ausente, muy por debajo del grupo']
  ]);

  // ---------- Ajustes_Coeval ----------
  sh = newSheet('Ajustes_Coeval');
  sh.getRange('A1').setValue('AJUSTES COEVALUACIÓN · Calculado automáticamente desde hoja Coevaluacion').setFontWeight('bold');
  sh.getRange(2, 1, 1, 19).setValues([[
    'Código','Alumno','Grupo',
    'RP1·Prom_recibido','Prom_grupo','Desviación','Ajuste',
    'RP2·Prom_recibido','Prom_grupo','Desviación','Ajuste',
    'RP3·Prom_recibido','Prom_grupo','Desviación','Ajuste',
    'TF_grupal·Prom_recibido','Prom_grupo','Desviación','Ajuste'
  ]]).setFontWeight('bold');

  function acFormulas(r) {
    var dev = function(avgCell, groupAvgCell) { return '=IFERROR(IF(' + avgCell + '="","",' + avgCell + '-' + groupAvgCell + '),"")'; };
    var adj = function(devCell) {
      return '=IFERROR(IF(' + devCell + '="","",IF(' + devCell + '>=-16,IF(' + devCell + '>=-6,IF(' + devCell + '>=10,2,0),-2),-4)),"--")';
    };
    var recv = function(instr) { return '=IFERROR(AVERAGEIFS(Coevaluacion!$F:$F,Coevaluacion!$E:$E,$B' + r + ',Coevaluacion!$B:$B,"' + instr + '"),"")'; };
    var grpAvg = function(instr) { return '=IFERROR(AVERAGEIFS(Coevaluacion!$F:$F,Coevaluacion!$D:$D,$C' + r + ',Coevaluacion!$B:$B,"' + instr + '"),"")'; };
    return [
      recv('RP1'), grpAvg('RP1'), dev('D'+r,'E'+r), adj('F'+r),
      recv('RP2'), grpAvg('RP2'), dev('H'+r,'I'+r), adj('J'+r),
      recv('RP3'), grpAvg('RP3'), dev('L'+r,'M'+r), adj('N'+r),
      recv('TF_grupal'), grpAvg('TF_grupal'), dev('P'+r,'Q'+r), adj('R'+r)
    ];
  }
  var acData = [];
  for (var i = 0; i < N; i++) {
    var r = 3 + i;
    acData.push(ROSTER[i].concat(acFormulas(r)));
  }
  sh.getRange(3, 1, N, 19).setValues(acData);
  sh.getRange(3, 1, N, 1).setNumberFormat('@');

  // ---------- Consolidacion ----------
  sh = newSheet('Consolidacion');
  sh.getRange('A1').setValue('CONSOLIDACIÓN · Solo lectura. Integra Notas_Manuales + Ajustes_Coeval + Registro_Sesiones').setFontWeight('bold');
  sh.getRange(2, 4).setValue('RP1 · 10%');
  sh.getRange(2, 8).setValue('RP2 · 10%');
  sh.getRange(2, 12).setValue('RP3 · 10%');
  sh.getRange(2, 16).setValue('CL · 20%');
  sh.getRange(2, 20).setValue('TF memo · 20%');
  sh.getRange(2, 23).setValue('TF present. · 20%');
  sh.getRange(2, 26).setValue('TF grupal · 10%');
  sh.getRange(2, 30).setValue('NOTA FINAL');
  sh.getRange(3, 1, 1, 30).setValues([[
    'Código','Alumno','Grupo',
    'Grupal','Coeval','Oral','Individual',
    'Grupal','Coeval','Oral','Individual',
    'Grupal','Coeval','Oral','Individual',
    'CL1','CL2','CL3','Promedio 2',
    'Nota','Oral','Individual',
    'Nota','Oral','Individual',
    'Grupal','Coeval','Oral','Individual',
    'Nota Final'
  ]]).setFontWeight('bold');

  function sump(r, instr) {
    var rr = 'Registro_Sesiones!$E$3:$E$499', dd = 'Registro_Sesiones!$D$3:$D$499', ii = 'Registro_Sesiones!$I$3:$I$499';
    return 'IFERROR(SUMPRODUCT((' + rr + '=$A' + r + ')*(' + dd + '="' + instr + '")*ISNUMBER(' + ii + ')*IF(ISNUMBER(' + ii + '),' + ii + ',0)),0)';
  }
  function consFormulas(r, acR) {
    var f = {};
    f.D = '=IFERROR(Notas_Manuales!D' + r + ',"")';
    f.E = '=IFERROR(Ajustes_Coeval!G' + acR + ',0)';
    f.F = '=' + sump(r, 'RP1');
    f.G = '=IFERROR(IF(D' + r + '="","",MIN(20,MAX(0,IF(ISNUMBER(E' + r + '),D' + r + '+E' + r + ',D' + r + ')+F' + r + '))),"")';
    f.H = '=IFERROR(Notas_Manuales!E' + r + ',"")';
    f.I = '=IFERROR(Ajustes_Coeval!K' + acR + ',0)';
    f.J = '=' + sump(r, 'RP2');
    f.K = '=IFERROR(IF(H' + r + '="","",MIN(20,MAX(0,IF(ISNUMBER(I' + r + '),H' + r + '+I' + r + ',H' + r + ')+J' + r + '))),"")';
    f.L = '=IFERROR(Notas_Manuales!F' + r + ',"")';
    f.M = '=IFERROR(Ajustes_Coeval!O' + acR + ',0)';
    f.N = '=' + sump(r, 'RP3');
    f.O = '=IFERROR(IF(L' + r + '="","",MIN(20,MAX(0,IF(ISNUMBER(M' + r + '),L' + r + '+M' + r + ',L' + r + ')+N' + r + '))),"")';
    f.P = '=IFERROR(Notas_Manuales!G' + r + ',"")';
    f.Q = '=IFERROR(Notas_Manuales!H' + r + ',"")';
    f.R = '=IFERROR(Notas_Manuales!I' + r + ',"")';
    f.S = '=IFERROR(IF(COUNTA(P' + r + ':R' + r + ')<2,"",MIN(20,MAX(0,(SUM(P' + r + ':R' + r + ')-MINIFS(P' + r + ':R' + r + ',P' + r + ':R' + r + ',"<>"))/2+(' +
          sump(r, 'CL1') + '+' + sump(r, 'CL2') + '+' + sump(r, 'CL3') + ')/3))),"")';
    f.T = '=IFERROR(Notas_Manuales!J' + r + ',"")';
    f.U = '=' + sump(r, 'TF_memo');
    f.V = '=IFERROR(IF(T' + r + '="","",MIN(20,MAX(0,T' + r + '+U' + r + '))),"")';
    f.W = '=IFERROR(Notas_Manuales!K' + r + ',"")';
    f.X = '=' + sump(r, 'TF_presentacion');
    f.Y = '=IFERROR(IF(W' + r + '="","",MIN(20,MAX(0,W' + r + '+X' + r + '))),"")';
    f.Z = '=IFERROR(Notas_Manuales!L' + r + ',"")';
    f.AA = '=IFERROR(Ajustes_Coeval!S' + acR + ',0)';
    f.AB = '=' + sump(r, 'TF_grupal');
    f.AC = '=IFERROR(IF(Z' + r + '="","",MIN(20,MAX(0,IF(ISNUMBER(AA' + r + '),Z' + r + '+AA' + r + ',Z' + r + ')+AB' + r + '))),"")';
    f.AD = '=IFERROR(IF(OR(G' + r + '="",K' + r + '="",O' + r + '="",S' + r + '="",V' + r + '="",Y' + r + '="",AC' + r + '=""),"Incompleta",' +
           'IF(MOD(G' + r + '*0.1+K' + r + '*0.1+O' + r + '*0.1+S' + r + '*0.2+V' + r + '*0.2+Y' + r + '*0.2+AC' + r + '*0.1,1)>=0.5,' +
           'INT(G' + r + '*0.1+K' + r + '*0.1+O' + r + '*0.1+S' + r + '*0.2+V' + r + '*0.2+Y' + r + '*0.2+AC' + r + '*0.1)+1,' +
           'INT(G' + r + '*0.1+K' + r + '*0.1+O' + r + '*0.1+S' + r + '*0.2+V' + r + '*0.2+Y' + r + '*0.2+AC' + r + '*0.1))),"Error")';
    var order = ['D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z','AA','AB','AC','AD'];
    return order.map(function(c){ return f[c]; });
  }
  var consData = [];
  for (var i2 = 0; i2 < N; i2++) {
    var r2 = 4 + i2, acR2 = r2 - 1;
    consData.push(ROSTER[i2].concat(consFormulas(r2, acR2)));
  }
  sh.getRange(4, 1, N, 30).setValues(consData);
  sh.getRange(4, 1, N, 1).setNumberFormat('@');

  // ---------- Vista_PAIDEIA ----------
  sh = newSheet('Vista_PAIDEIA');
  sh.getRange('A1').setValue('VISTA PAIDEIA · Una sección por acta. Cada sección exportable como CSV (código + nota entera, sin encabezados).').setFontWeight('bold');
  var actas = [
    [1,'CONTROL DE LECTURA 1'],[5,'CONTROL DE LECTURA 2'],[9,'CONTROL DE LECTURA 3'],
    [13,'REPORTE 1'],[17,'REPORTE 2'],[21,'REPORTE 3'],[25,'TRABAJO FINAL']
  ];
  actas.forEach(function(a) {
    sh.getRange(2, a[0]).setValue('ACTA · ' + a[1]);
    sh.getRange(3, a[0], 1, 3).setValues([['Código PUCP','Apellidos y Nombres','Nota']]).setFontWeight('bold');
    sh.getRange(4, a[0]).setValue('CSV: cols A+C de esta sección');
  });

  function clFormula(consR, notasCol, instr) {
    var ee = 'Registro_Sesiones!$E$3:$E$499', dd = 'Registro_Sesiones!$D$3:$D$499', ii = 'Registro_Sesiones!$I$3:$I$499';
    var base = 'IFERROR(IF(Notas_Manuales!' + notasCol + consR + '="","",MIN(20,MAX(0,Notas_Manuales!' + notasCol + consR +
      '+IFERROR(SUMPRODUCT((' + ee + '=Consolidacion!$A' + consR + ')*(' + dd + '="' + instr + '")*IF(ISNUMBER(' + ii + '),' + ii + ',0)),0)))),"")';
    return '=IFERROR(IF(' + base + '="","",IF(MOD(' + base + ',1)>=0.5,INT(' + base + ')+1,INT(' + base + '))),"—")';
  }
  function rpFormula(consR, consCol) {
    return '=IFERROR(IF(Consolidacion!' + consCol + consR + '="","",IF(MOD(Consolidacion!' + consCol + consR + ',1)>=0.5,' +
      'INT(Consolidacion!' + consCol + consR + ')+1,INT(Consolidacion!' + consCol + consR + '))),"—")';
  }
  function tfFormula(consR) {
    var base = 'IFERROR(IF(OR(Consolidacion!V' + consR + '="",Consolidacion!Y' + consR + '="",Consolidacion!AC' + consR + '=""),"",' +
      'Consolidacion!V' + consR + '*0.4+Consolidacion!Y' + consR + '*0.4+Consolidacion!AC' + consR + '*0.2),"—")';
    return '=IFERROR(IF(' + base + '="","",IF(MOD(' + base + ',1)>=0.5,INT(' + base + ')+1,INT(' + base + '))),"—")';
  }

  for (var i3 = 0; i3 < N; i3++) {
    var r3 = 5 + i3, consR3 = r3 - 1;
    var codigo = ROSTER[i3][0], nombre = ROSTER[i3][1];
    var blocks = [
      [1, clFormula(consR3, 'G', 'CL1')],
      [5, clFormula(consR3, 'H', 'CL2')],
      [9, clFormula(consR3, 'I', 'CL3')],
      [13, rpFormula(consR3, 'G')],
      [17, rpFormula(consR3, 'K')],
      [21, rpFormula(consR3, 'O')],
      [25, tfFormula(consR3)]
    ];
    blocks.forEach(function(b) {
      sh.getRange(r3, b[0], 1, 3).setValues([[codigo, nombre, b[1]]]);
    });
  }
  sh.getRange(5, 1, N, 1).setNumberFormat('@');
  sh.getRange(5, 5, N, 1).setNumberFormat('@');
  sh.getRange(5, 9, N, 1).setNumberFormat('@');
  sh.getRange(5, 13, N, 1).setNumberFormat('@');
  sh.getRange(5, 17, N, 1).setNumberFormat('@');
  sh.getRange(5, 21, N, 1).setNumberFormat('@');
  sh.getRange(5, 25, N, 1).setNumberFormat('@');

  // Reordenar y limpiar hoja default
  var order = ['INSTRUCCIONES','Registro_Sesiones','Notas_Manuales','Coevaluacion','Metodología_Coeval','Ajustes_Coeval','Consolidacion','Vista_PAIDEIA'];
  order.forEach(function(name, idx) { ss.setActiveSheet(ss.getSheetByName(name)); ss.moveActiveSheet(idx + 1); });
  var def = ss.getSheetByName('Hoja 1') || ss.getSheetByName('Sheet1');
  if (def) ss.deleteSheet(def);

  Logger.log('Listo: ' + N + ' alumnos cargados en ' + order.length + ' hojas.');
}
