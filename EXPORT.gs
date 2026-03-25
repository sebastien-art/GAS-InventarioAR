/**
 * @OnlyCurrentDoc
 * Exporta los datos filtrados desde "Naves" al archivo fijo de propuestas.
 * Crea una pestaña con el nombre del cliente (sin fecha) y actualiza el Menú.
 */

function exportarFilasVisiblesAFijo() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaOrigen = ss.getSheetByName("Naves");
  if (!hojaOrigen) return ui.alert("❌ No se encontró la hoja 'Naves'.");

  const filtro = hojaOrigen.getFilter();
  if (!filtro) return ui.alert("❌ Aplica un filtro antes de exportar.");

  // 🔹 DESTINO: Archivo “Propuestas comerciales”
  const ID_DESTINO = "1vodBL6mNl6rkrjuNdQEnIktC8TlRzHS-J446piiKbJE";

  const columnasDeseadas = [
    "Intermediario", "Operación", "Ficha", "REF", "Estado", "Zona Principal", "Sub Zona",
    "Desarrollador", "Parque", "Nave", "M2 de construcción", "M2 de terreno",
    "Asking price /m2", "Mantenimiento / m2", "Disponibilidad","Energía (kVAs)", "Comentarios", "Coordenadas", "Ubicación", "Altura libre", "Altura máxima"
   ];

  try {
    const respuesta = ui.prompt("Nombre de la propuesta", "Escribe el nombre del cliente:", ui.ButtonSet.OK_CANCEL);
    if (respuesta.getSelectedButton() !== ui.Button.OK) return ui.alert("Operación cancelada.");
    const baseName = respuesta.getResponseText().trim() || "Propuesta sin nombre";

    ss.toast("Copiando datos filtrados...", "Paso 1/3", -1);

    // Crear hoja temporal con los datos visibles
    const rangoTotal = hojaOrigen.getDataRange();
    const tempSheet = ss.insertSheet("TEMP_FILTER_EXPORT");
    rangoTotal.copyTo(tempSheet.getRange(1, 1), { contentsOnly: true });

    ss.toast("Procesando columnas...", "Paso 2/3", -1);

    const allData = tempSheet.getDataRange().getValues();
    const encabezados = allData[0];

    const indices = columnasDeseadas.map(col => {
      const i = encabezados.indexOf(col);
      if (i === -1) throw new Error(`Falta columna "${col}"`);
      return i;
    });

    const salida = [["ENVIAR"].concat(columnasDeseadas)];

    for (let i = 1; i < allData.length; i++) {
      const fila = allData[i];
      if (fila.some(v => v !== "" && v !== null)) {
        salida.push([""].concat(indices.map(ix => fila[ix] ?? "")));
      }
    }

    ss.deleteSheet(tempSheet);

    if (salida.length <= 1) return ui.alert("⚠️ No hay datos para exportar.");

    ss.toast("Creando en archivo destino...", "Paso 3/3", -1);
    const ssDestino = SpreadsheetApp.openById(ID_DESTINO);

    // Evitar duplicados
    let nombreHoja = baseName;
    let contador = 1;
    while (ssDestino.getSheetByName(nombreHoja)) nombreHoja = `${baseName} (${contador++})`;

    const hojaNueva = ssDestino.insertSheet(nombreHoja);
    hojaNueva.getRange(1, 1, salida.length, salida[0].length).setValues(salida);

    // ===== Formato =====
    hojaNueva.setFrozenRows(1);
    hojaNueva.setFrozenColumns(1);
    const numColsFinal = salida[0].length;

    hojaNueva.getRange(1, 1, 1, numColsFinal)
      .setBackground("#b6d7a8")
      .setFontWeight("bold")
      .setFontSize(10)
      .setHorizontalAlignment("center");

    // Anchos: 50 px para ENVIAR, Intermediario, Operación, Ficha, REF, Nave; 100 px para el resto
    const SMALL = new Set(["ENVIAR", "Intermediario", "Operación", "Ficha", "REF", "Nave"]);
    const headersFinal = salida[0]; // ["ENVIAR", ...columnasDeseadas]
    for (let c = 0; c < headersFinal.length; c++) {
      const header = headersFinal[c];
      hojaNueva.setColumnWidth(c + 1, SMALL.has(header) ? 50 : 100);
    }
    // ====================

    actualizarMenu(ssDestino, hojaNueva);

    SpreadsheetApp.flush();
    ss.toast("¡Completado!", "✅", 2);

    ui.alert(`✅ Exportado`, `"${nombreHoja}" creado con ${salida.length - 1} filas.`, ui.ButtonSet.OK);

  } catch (err) {
    try {
      const temp = ss.getSheetByName("TEMP_FILTER_EXPORT");
      if (temp) ss.deleteSheet(temp);
    } catch (e) {}
    ui.alert("❌ Error", err.message, ui.ButtonSet.OK);
  }
}

function actualizarMenu(ssDestino, hojaNueva) {
  let hojaMenu = ssDestino.getSheetByName("Menú");
  if (!hojaMenu) hojaMenu = ssDestino.insertSheet("Menú", 0);

  if (hojaMenu.getLastRow() < 1) {
    hojaMenu.getRange("A1:B1").setValues([["Propuestas/Reportes", "Navegar"]])
      .setBackground("#b6d7a8").setFontWeight("bold");
    hojaMenu.setFrozenRows(1).setColumnWidth(1, 250).setColumnWidth(2, 90);
  }

  const nombre = hojaNueva.getName();
  const fila = hojaMenu.getLastRow() + 1;

  hojaMenu.getRange(fila, 1).setValue(nombre);
  hojaMenu.getRange(fila, 2)
    .setFormula(`=HYPERLINK("#gid=${hojaNueva.getSheetId()}", "VER DATOS")`)
    .setBackground("#007bff")
    .setFontColor("white")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
}
