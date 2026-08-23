/**
 * Google Apps Script para el examen de PLN.
 *
 *  - doGet()  -> LEE las preguntas desde la pestaña "Estructura" y las devuelve como JSON.
 *  - doPost() -> GUARDA los resultados del estudiante en la pestaña con el nombre del examen.
 */

// ================= LEER PREGUNTAS / SECCIONES =================
function doGet(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // --- Lectura de la pestaña "Secciones" (base de NRC y docentes EN VIVO) ---
    if (e && e.parameter && e.parameter.type === 'secciones') {
      return getSecciones(ss);
    }

    // Primero intentar leer la estructura JSON centralizada desde BD_Examenes
    var sheetData = ss.getSheetByName('BD_Examenes');
    if (sheetData) {
      var vals = sheetData.getDataRange().getValues();
      var jsonStr = '';
      for(var i = 0; i < vals.length; i++){ jsonStr += vals[i][0]; }
      if (jsonStr) {
        var parsed = JSON.parse(jsonStr);
        return jsonOut({ ok: true, data: parsed, source: 'BD_Examenes' });
      }
    }

    // Fallback: modo antiguo leyendo "Estructura" fila por fila
    var sheet = ss.getSheetByName('Estructura');
    if (!sheet) {
      return jsonOut({ ok: false, error: 'No existe la pestaña "BD_Examenes" ni "Estructura".' });
    }

    var data = sheet.getDataRange().getValues();
    var questions = [];
    var current = null;

    for (var i = 0; i < data.length; i++) {
      var b = String(data[i][1] == null ? '' : data[i][1]).trim(); // Columna B
      var c = String(data[i][2] == null ? '' : data[i][2]).trim(); // Columna C
      var d = data[i][3];                                          // Columna D (checkbox)

      if (/^PREGUNTA/i.test(b)) {
        if (current) questions.push(current);
        current = { title: b, text: '', options: [] };   // title = nombre (col B)
      } else if (current && current.text === '' && b === '' && c !== '') {
        current.text = c;                     // primera línea de texto = enunciado (col C)
      } else if (current && /^[a-zA-Z]\s*\)/.test(b)) {
        current.options.push({
          letter: b.charAt(0).toLowerCase(),
          text: c,
          correct: (d === true || String(d).toUpperCase() === 'TRUE' || d === 'x' || d === 'X')
        });
      }
    }
    if (current) questions.push(current);

    // Filtra preguntas incompletas
    questions = questions.filter(function (q) { return q.options.length > 0; });

    return jsonOut({ ok: true, questions: questions });

  } catch (err) {
    return jsonOut({ ok: false, error: err.toString() });
  }
}

// ================= GUARDAR RESULTADOS O EXÁMENES =================
function doPost(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var d = JSON.parse(e.postData.contents);
    
    // --- Acción: Guardar estructura de exámenes ---
    if (d.action === 'saveExams') {
      var sheetData = ss.getSheetByName('BD_Examenes') || ss.insertSheet('BD_Examenes');
      sheetData.clear();
      var jsonStr = JSON.stringify(d.data);
      // Dividir el JSON en bloques de 40000 caracteres para evitar límite por celda
      var chunks = [];
      for(var i = 0; i < jsonStr.length; i += 40000) {
        chunks.push([jsonStr.substring(i, i + 40000)]);
      }
      sheetData.getRange(1, 1, chunks.length, 1).setValues(chunks);
      return jsonOut({ ok: true });
    }

    // --- Acción: Cargar estructura de exámenes ---
    if (d.action === 'loadExams') {
      var sheetData = ss.getSheetByName('BD_Examenes');
      if(!sheetData) return jsonOut({ ok: true, data: { exams: [] } });
      var values = sheetData.getDataRange().getValues();
      var jsonStr = '';
      for(var i = 0; i < values.length; i++){ jsonStr += values[i][0]; }
      return jsonOut({ ok: true, data: JSON.parse(jsonStr) });
    }

    // --- Acción por defecto: Guardar resultados del estudiante ---
    var sheetName = d.examName ? d.examName : 'Resultados';
    sheetName = sheetName.replace(/[\[\]\*\\\?\:]/g, "").substring(0, 95);
    
    var hoja = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);

    if (hoja.getLastRow() === 0) {
      var headers = [
        'Fecha', 'Nombres', 'Apellidos', 'NRC',
        'Aciertos', 'Total', 'Porcentaje', 'Enviado'
      ];
      
      if (d.respuestasDetalladas && d.respuestasDetalladas.length > 0) {
        for (var i = 0; i < d.respuestasDetalladas.length; i++) {
          var p = d.respuestasDetalladas[i];
          var titulo = p.pregunta || ('P' + (i+1));
          headers.push(titulo + ' (Respuesta)');
          headers.push(titulo + ' (Bien/Mal)');
        }
      }
      hoja.appendRow(headers);
      hoja.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#f3f3f3");
    }

    var rowData = [
      d.fecha, d.nombres, d.apellidos, d.nrc,
      d.aciertos, d.total, d.porcentaje + '%', d.enviado
    ];
    
    if (d.respuestasDetalladas && d.respuestasDetalladas.length > 0) {
      for (var i = 0; i < d.respuestasDetalladas.length; i++) {
        var p = d.respuestasDetalladas[i];
        rowData.push(p.elegida);
        rowData.push(p.esCorrecta ? 'Bien' : 'Mal');
      }
    }

    hoja.appendRow(rowData);

    return jsonOut({ ok: true });

  } catch (err) {
    return jsonOut({ ok: false, error: err.toString() });
  }
}

// ================= LEER PESTAÑA "SECCIONES" =================
// Devuelve todas las filas como objetos {Header: valor}, tomando SIEMPRE
// los datos actuales de la pestaña "Secciones" del spreadsheet activo.
function getSecciones(ss) {
  var sheet = ss.getSheetByName('Secciones');
  if (!sheet) {
    return jsonOut({ ok: false, error: 'No existe la pestaña "Secciones".' });
  }
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) {
    return jsonOut({ ok: true, secciones: [] });
  }
  var headers = values[0].map(function (h) { return String(h).trim(); });
  var rows = [];
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    // Omitir filas totalmente vacías
    var vacia = row.every(function (c) { return String(c).trim() === ''; });
    if (vacia) continue;
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      obj[headers[j]] = row[j] == null ? '' : String(row[j]).trim();
    }
    rows.push(obj);
  }
  return jsonOut({ ok: true, secciones: rows });
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
