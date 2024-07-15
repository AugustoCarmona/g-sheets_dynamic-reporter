# Google Sheets - Generación de Reportes Dinámicos

Este script de Google Apps Script agrega un menú personalizado en una hoja de cálculo de Google Sheets para crear un reporte basado en los datos de una plantilla específica. Utiliza la función `obtenerRangoDinamico` para fragmentar una tabla en multiples secciones, permitiendo al usuario de la plantilla ingresar cuantas filas sean necesarias en la hoja de cálculo.

## Función Principal

### 1. `obtenerRangoDinamico`
Esta función ayuda a obtener un rango dinámico de datos de la hoja basada en las secciones especificadas.
- Dentro de la iteración, `startRow = i + 2` suma dos filas ya que una vez que encuentra la fila con el nombre de la sección (por ejemplo "Puertas"), necesita saltear dicha fila y la fila siguiente que contiene los headers de la tabla.
- Cada sección debe tener un identificador de inicio y de fin, en este caso se identifica el inicio con el nombre propio de la sección el cual se ubica en la columna "A" y el fin con el dato "Subtotal" (ubicado en la columna F), motivo por el cual se observa en el iterador que `data[i][5] === finSeccion`. Si en contraparte necesitamos observar el indicador de fin en la columna "C", se lo ubicará por `data[i][2] === finSeccion` ya que Google Sheets indexa las columnas desde 0.
```javascript
function obtenerRangoDinamico(sheet, seccion, finSeccion) {
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
```

## Cómo Usar
1. Abre tu hoja de cálculo de Google Sheets.
2. Haz clic en "Extensiones" y selecciona "Apps Script".
3. Copia y pega el código en el editor y guarda el proyecto.
4. Recarga la hoja de cálculo.
5. Verás un nuevo menú llamado "Herramientas Matteucci" (agregá el nombre que corresponda para tu proyecto). Hacé clic en él y selecciona "Generar Reporte" para crear el reporte en una hoja nueva llamada "Reporte".
