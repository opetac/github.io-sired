function doGet(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Respuestas de formulario 1") || SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    return ContentService.createTextOutput(JSON.stringify({ status: "success", records: [] })).setMimeType(ContentService.MimeType.JSON);
  }
  
  var headers = data[0];
  var records = [];
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (row.join("").trim() === "") continue;
    var item = {};
    for (var j = 0; j < headers.length; j++) {
      item[headers[j]] = row[j];
    }
    records.push(item);
  }
  
  return ContentService.createTextOutput(JSON.stringify({
    status: "success",
    records: records.slice(-50).reverse() // Los 50 registros más recientes
  })).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Respuestas de formulario 1") || SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    var body = JSON.parse(e.postData.contents);
    
    // Gestión de subida de fotografía a Google Drive
    var photoUrl = "";
    if (body.fotoBase64) {
      try {
        var folder;
        var folders = DriveApp.getFoldersByName("FOTOS_DETENIDOS_APP");
        if (folders.hasNext()) {
          folder = folders.next();
        } else {
          folder = DriveApp.createFolder("FOTOS_DETENIDOS_APP");
        }
        
        var splitBase = body.fotoBase64.split(",");
        var contentType = splitBase[0].split(":")[1].split(";")[0];
        var bytes = Utilities.base64Decode(splitBase[1]);
        var blob = Utilities.newBlob(bytes, contentType, "DETENIDO_" + (body.nombre || "S_N") + "_" + new Date().getTime() + ".jpg");
        var file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        photoUrl = file.getUrl();
      } catch (errFoto) {
        photoUrl = "Error al guardar foto: " + errFoto.toString();
      }
    }

    var timestamp = new Date();
    
    // Inserción en el orden exacto de columnas del Google Sheet
    var rowData = [
      timestamp,                             // Marca temporal
      body.motivoDetencion || "",           // MOTIVO DE DETENCION
      body.nombre || "",                    // NOMBRE
      body.fechaDetencion || "",            // FECHA DETENCION
      body.horaDetencion || "",             // HORA DETENCION
      body.calle || "",                     // CALLE
      body.numero || "",                    // N°
      body.entreCalle || "",                // ENTRE CALLE
      body.colonia || "",                   // COLONIA
      body.municipio || "",                 // MUNICIPIO
      body.codigoPostal || "",              // CODIGO POSTAL
      body.estado || "NUEVO LEON",          // ESTADO
      body.telefono || "",                  // TELEFONO
      photoUrl,                             // FOTOGRAFIA DETENIDO
      photoUrl,                             // ENLACE
      body.apellidoPaterno || "",           // APELLIDO PATERNO
      body.apellidoMaterno || "",           // APELLIDO MATERNO
      body.fechaNacimiento || "",           // FECHA DE NACIMIENTO
      body.curp || "",                      // CURP
      body.nombrePadre || "",               // NOMBRE PADRE
      body.nombreMadre || "",               // NOMBRE MADRE
      "", "",                               // HERMANOS, HERMANOS
      body.esposaConyuge || "",             // ESPOS@ y/o CONYUGE
      "", "",                               // HIJOS, HIJOS
      body.folio || "",                     // FOLIO
      body.edad || "",                      // EDAD
      body.ocupacion || "",                 // OCUPACION
      body.gradoEstudios || "",             // GRADO DE ESTUDIOS
      body.originario || "",                // ORIGINARIO
      body.estadoCivil || "",               // ESTADO CIVIL
      body.lugarDetencion || "",            // LUGAR DE LA DETENCION
      body.objetoAsegurado || "",           // OBJETO ASEGURADO
      body.indicios || "",                  // INDICIOS
      body.alias || "",                     // ALIAS
      body.latitud || "",                   // LATITUD
      body.longitud || "",                  // LONGITUD
      body.email || "",                     // Dirección de correo electrónico
      body.entrevista || ""                 // ENTREVISTA
    ];

    sheet.appendRow(rowData);

    return ContentService.createTextOutput(JSON.stringify({ status: "success" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
