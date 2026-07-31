/**
 * Google Apps Script para el examen de PLN.
 *
 *  - doGet()  -> LEE las preguntas desde la pestaña "Estructura" y las devuelve como JSON.
 *  - doPost() -> GUARDA los resultados del estudiante en la pestaña con el nombre del examen.
 */

// ================= LEER PREGUNTAS =================
function doGet(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Estructura');
    if (!sheet) {
      return jsonOut({ ok: false, error: 'No existe la pestaña "Estructura".' });
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

// ================= GUARDAR RESULTADOS =================
function doPost(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var d = JSON.parse(e.postData.contents);
    
    // Usar el nombre del examen para la pestaña, por defecto "Resultados" si viene vacío
    var sheetName = d.examName ? d.examName : 'Resultados';
    // Limpiar nombre de hoja para que sea válido en Google Sheets (max 100 chars, sin caracteres raros)
    sheetName = sheetName.replace(/[\[\]\*\\\?\:]/g, "").substring(0, 95);
    
    var hoja = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);

    // Si la hoja está vacía, preparamos las cabeceras dinámicamente
    if (hoja.getLastRow() === 0) {
      var headers = [
        'Fecha', 'Nombres', 'Apellidos', 'NRC',
        'Aciertos', 'Total', 'Porcentaje', 'Enviado'
      ];
      
      // Agregamos dos columnas por cada pregunta (Respuesta y si está Bien/Mal)
      if (d.respuestasDetalladas && d.respuestasDetalladas.length > 0) {
        for (var i = 0; i < d.respuestasDetalladas.length; i++) {
          var p = d.respuestasDetalladas[i];
          var titulo = p.pregunta || ('P' + (i+1));
          headers.push(titulo + ' (Respuesta)');
          headers.push(titulo + ' (Bien/Mal)');
        }
      }
      hoja.appendRow(headers);
      
      // Darle formato a las cabeceras
      hoja.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#f3f3f3");
    }

    var rowData = [
      d.fecha, d.nombres, d.apellidos, d.nrc,
      d.aciertos, d.total, d.porcentaje + '%', d.enviado
    ];
    
    // Agregar las respuestas detalladas a la fila
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

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
