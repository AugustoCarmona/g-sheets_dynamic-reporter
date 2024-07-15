function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Herramientas Matteucci')
      .addItem('Generar Reporte', 'procesarDatos')
      .addToUi();
  }
  
  function procesarDatos() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var outputSheet = ss.getSheetByName('output');
    if (!outputSheet) {
      outputSheet = ss.insertSheet('output');
    } else {
      outputSheet.clear(); // limpia la hoja si ya existe
    }
  
    var sheet = ss.getSheetByName('Nueva Plantilla');
    if (!sheet) {
      SpreadsheetApp.getUi().alert('La hoja "Nueva Plantilla" no se encuentra.');
      return;
    }

    // datos del cliente -----------------------------------------------------------------------------------------------------------------------
    var datosCliente = {
      "Nombre": sheet.getRange('B6').getValue(),
      "Calle": sheet.getRange('B7').getValue(),
      "Teléfono": sheet.getRange('B8').getValue(),
      "Celular": sheet.getRange('B9').getValue(),
      "Mail": sheet.getRange('B10').getValue(),
      "Fecha": sheet.getRange('G6').getValue(),
      "Lote": sheet.getRange('G7').getValue(),
      "Manzana": sheet.getRange('G8').getValue(),
      "Barrio": sheet.getRange('G9').getValue(),
      "Arquitecto": sheet.getRange('G10').getValue()
    };

    // datos del pedido ------------------------------------------------------------------------------------------------------------------------
    let pedidosPuertas = obtenerRangoDinamico(sheet, "Puertas", "Sub Total");
    let pedidosAccesorios = obtenerRangoDinamico(sheet, "Accesorios", "Sub Total");
    let pedidosPintura = obtenerRangoDinamico(sheet, "Pintura", "Sub Total");
  
    let headers = [
      ["Nombre", "Calle", "Telefono", "Celular", "Mail", "Lote", "Manzana", "Barrio", "Arquitecto",
       "Cantidad Puertas", "Cp Puerta", "Tipo Puerta", "Modelo Puerta", "Acabado Puerta", "Precio Unit Puerta", "Precio Total Puerta", "Subtotal Puertas",
       "Cantidad Accesorio", "Modelo Accesorio", "Precio Unit Accesorio", "Precio Total Accesorio", "Subtotal Accesorios",
       "M2 Pintura", "Tipo Pintura", "Color Pintura", "Precio Unit Pintura", "Precio Total Pintura", "Subtotal Pintura"]
    ];
  
    let dataTabla = [];
  
    pedidosPuertas.forEach(row => {
      if (row.some(cell => cell !== "")) { // verifica si la fila no esta vacia
        dataTabla.push([
          datosCliente.Nombre,
          datosCliente.Calle,
          datosCliente.Teléfono,
          datosCliente.Celular,
          datosCliente.Mail,
          datosCliente.Lote,
          datosCliente.Manzana,
          datosCliente.Barrio,
          datosCliente.Arquitecto,
          row[0], // Cantidad_Puertas
          row[1], // Cp_Puerta
          row[2], // Tipo_Puerta
          row[3], // Modelo_Puerta
          row[4], // Acabado_Puerta
          row[6], // Precio_Unit_Puerta
          row[7], // Precio_Total_Puerta
          "", // Subtotal_Puertas (calculamos después)
          "", // Cantidad_Accesorio
          "", // Modelo_Accesorio
          "", // Precio_Unit_Accesorio
          "", // Precio_Total_Accesorio
          "", // Subtotal_Accesorios
          "", // M2_Pintura
          "", // Tipo_Pintura
          "", // Color_Pintura
          "", // Precio_Unit_Pintura
          "", // Precio_Total_Pintura
          ""  // Subtotal_Pintura
        ]);
      }
    });
  
    pedidosAccesorios.forEach(row => {
      if (row.some(cell => cell !== "")) {  // verifica si la fila no esta vacia
        dataTabla.push([
          datosCliente.Nombre,
          datosCliente.Calle,
          datosCliente.Teléfono,
          datosCliente.Celular,
          datosCliente.Mail,
          datosCliente.Lote,
          datosCliente.Manzana,
          datosCliente.Barrio,
          datosCliente.Arquitecto,
          "", // Cantidad_Puertas
          "", // Cp_Puerta
          "", // Tipo_Puerta
          "", // Modelo_Puerta
          "", // Acabado_Puerta
          "", // Precio_Unit_Puerta
          "", // Precio_Total_Puerta
          "", // Subtotal_Puertas
          row[0], // Cantidad_Accesorio
          row[3], // Modelo_Accesorio
          row[6], // Precio_Unit_Accesorio
          row[7], // Precio_Total_Accesorio
          "", // Subtotal_Accesorios
          "", // M2_Pintura
          "", // Tipo_Pintura
          "", // Color_Pintura
          "", // Precio_Unit_Pintura
          "", // Precio_Total_Pintura
          ""  // Subtotal_Pintura
        ]);
      }
    });
  
    pedidosPintura.forEach(row => {
      if (row.some(cell => cell !== "")) {  // verifica si la fila no esta vacia
        dataTabla.push([
          datosCliente.Nombre,
          datosCliente.Calle,
          datosCliente.Teléfono,
          datosCliente.Celular,
          datosCliente.Mail,
          datosCliente.Lote,
          datosCliente.Manzana,
          datosCliente.Barrio,
          datosCliente.Arquitecto,
          "", // Cantidad_Puertas
          "", // Cp_Puerta
          "", // Tipo_Puerta
          "", // Modelo_Puerta
          "", // Acabado_Puerta
          "", // Precio_Unit_Puerta
          "", // Precio_Total_Puerta
          "", // Subtotal_Puertas
          "", // Cantidad_Accesorio
          "", // Modelo_Accesorio
          "", // Precio_Unit_Accesorio
          "", // Precio_Total_Accesorio
          "", // Subtotal_Accesorios
          row[0], // M2_Pintura
          row[1], // Tipo_Pintura
          row[2], // Color_Pintura
          row[6], // Precio_Unit_Pintura
          row[7], // Precio_Total_Pintura
          ""  // Subtotal_Pintura
        ]);
      }
    });
  
    let tablaCompleta = headers.concat(dataTabla);
  
    //------------------------------------------------------------------------------------------------------------------------------------------
    // pega los datos
    try {
      outputSheet.getRange(1, 1, tablaCompleta.length, tablaCompleta[0].length).setValues(tablaCompleta);
    } catch (e) {
      SpreadsheetApp.getUi().alert('Error al copiar los datos: ' + e.message);
      return;
    }
  
    // registro de actividad
    Logger.log("Datos pegados en la hoja 'output'.");
    SpreadsheetApp.getUi().alert("Datos pegados en la hoja 'output'.");
  }
  
  function obtenerRangoDinamico(sheet, seccion, finSeccion) {
    /**
     * itera sobre las filas de la hoja y devuelve un rango de filas fraccionado por el inicio y el fin de la seccion
     */
    var data = sheet.getDataRange().getValues();
    var startRow, endRow;
  
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] === seccion) {
        startRow = i + 2;
      }
      if (data[i][5] === finSeccion && startRow !== undefined) {
        endRow = i;
        break;
      }
    }
  
    if (startRow !== undefined && endRow !== undefined) {
      return sheet.getRange(startRow + 1, 1, endRow - startRow, 8).getValues();
    } else {
      SpreadsheetApp.getUi().alert('No se encontró la sección "' + seccion + '" o el final de la sección.');
      return [];
    }
  }
  