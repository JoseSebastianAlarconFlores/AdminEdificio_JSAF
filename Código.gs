/**
 * =========================================================================
 * CÓDIGO INTEGRADO PARA GESTIÓN EDIFICIO JSAF
 * =========================================================================
 */

function doGet(e) {
  var datos = obtenerDatosEdificio();
  var output = JSON.stringify(datos);
  return ContentService.createTextOutput(output)
    .setMimeType(ContentService.MimeType.JSON);
}

function obtenerDatosEdificio() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('BD edificio JSAF') || ss.getSheets()[0];
  var data = sheet.getDataRange().getValues();
  var objetoEdificio = {};

  var limpiarNum = function (valor) {
    if (valor === undefined || valor === null || valor === "") return 0;
    if (typeof valor === 'number') return valor;
    var texto = String(valor).trim().replace(/[^0-9,.-]/g, '');
    if (texto.indexOf(',') !== -1 && texto.indexOf('.') === -1) {
      texto = texto.replace(',', '.');
    } else if (texto.indexOf(',') !== -1 && texto.indexOf('.') !== -1) {
      texto = texto.replace(',', '');
    }
    return parseFloat(texto) || 0;
  };

  for (var i = 1; i < data.length; i++) {
    var fila = data[i];
    var idDepto = String(fila[1]).trim(); // Columna B
    if (idDepto && idDepto !== "" && !isNaN(idDepto)) {
      objetoEdificio[idDepto] = {
        inquilino: String(fila[2] || "Sin Nombre"),
        fono: String(fila[4] || "").replace(/[^0-9+]/g, ''),
        alquiler: limpiarNum(fila[5]),
        mantenimiento: limpiarNum(fila[6]),
        limpieza: limpiarNum(fila[7]),
        seguridad: limpiarNum(fila[8]),
        agua: limpiarNum(fila[9]),
        electricity: limpiarNum(fila[10]),
        totalMora: limpiarNum(fila[11]),    // Col L
        cancelado: limpiarNum(fila[13]),    // Col N
        diferencia: limpiarNum(fila[14]),   // Col O
        estadoSheet: String(fila[15] || "Al día").trim() // Col P
      };
    }
  }
  return objetoEdificio;
}

function registrarPagoConArchivo(e) {
  try {
    if (!e.datosBase64 || !e.idDepto) {
      throw new Error("Datos incompletos para el registro.");
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName('BD edificio JSAF') || ss.getSheets()[0];
    var datos = hoja.getDataRange().getValues();
    var filaEncontrada = -1;

    for (var i = 1; i < datos.length; i++) {
      if (String(datos[i][1]).trim() === String(e.idDepto).trim()) {
        filaEncontrada = i + 1;
        break;
      }
    }

    if (filaEncontrada === -1) throw new Error("Depto '" + e.idDepto + "' no encontrado.");

    var idCarpetaVouchers = "1MgqL4E9_MbSgkRO4OH2HvwxRQPV_y5GP";
    var carpeta = DriveApp.getFolderById(idCarpetaVouchers);

    var bytes = Utilities.base64Decode(e.datosBase64);
    var blob = Utilities.newBlob(bytes, e.tipoMime, e.nombreArchivo);
    var archivo = carpeta.createFile(blob);

    try {
      archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (err) {
      Logger.log("Permiso público omitido: " + err.message);
    }

    var celdaRecibido = hoja.getRange(filaEncontrada, 14);
    var montoAnterior = parseFloat(celdaRecibido.getValue()) || 0;
    var montoNuevo = parseFloat(e.montoEnviado) || 0;
    celdaRecibido.setValue(montoAnterior + montoNuevo);

    hoja.getRange(filaEncontrada, 13).setValue(archivo.getUrl());

    SpreadsheetApp.flush();
    var nuevaDiferencia = parseFloat(hoja.getRange(filaEncontrada, 15).getValue()) || 0;

    if (Math.abs(nuevaDiferencia) <= 0.02) {
      hoja.getRange(filaEncontrada, 16).setValue("Al día");
    }

    return {
      exito: true,
      mensaje: "¡Pago registrado correctamente!"
    };

  } catch (err) {
    throw new Error(err.message);
  }
}
function doPost(e) {
  var output = ContentService.createTextOutput();

  try {
    var json = JSON.parse(e.postData.contents);
    var resultado = registrarPagoConArchivo(json);

    output.setContent(JSON.stringify({
      "status": "éxito",
      "mensaje": resultado.mensaje
    }));

  } catch (err) {
    output.setContent(JSON.stringify({
      "status": "error",
      "mensaje": err.toString()
    }));
  }
  output.setMimeType(ContentService.MimeType.JSON);
  output.setHeader("Access-Control-Allow-Origin", "*");
  return output;
}