function onEdit(e) {
    var sheetName = "Hoja1";  // Aquí debemos indicar el nombre exacto entre comillas de la hoja de cálculo sobre la que aplicaremos la función.
    var ss = e.source;  // Obtener la hoja de cálculo activa
    var range = e.range; // Obtener el rango editado (No es necesario con la configuración actual, pero mantenemos la variable)
    var editedColumn = range.getColumn(); // Obtener el número de columna editada
    var oldValue;
    var newValue;
    var activeCell = ss.getActiveCell();
    
    // Definir las columnas en las que vamos aplicar la funcionalidad de lista desplegable con selección múltiple
    var startColumn = 26; // Desde columna 26 (columna Z)
    var endColumn = 34;   // Hasta columna 34 (columna AH)
    var excludedColumns = [27, 30, 32, 33]; // Columnas que vamos a excluir del rango anterior: AA, AD, AF, AG
    
    // Aplicación de la función en base a condicionales.
    if (ss.getActiveSheet().getName() === sheetName && editedColumn >= startColumn && editedColumn <= endColumn && !excludedColumns.includes(editedColumn)) {
        newValue=e.value;
        oldValue=e.oldValue;
        if(!e.value) {
            activeCell.setValue("");
        } else {
            if (!e.oldValue) {
                activeCell.setValue(newValue);
            } else {
                activeCell.setValue(oldValue+', '+newValue);
            }
        }
    }
}
