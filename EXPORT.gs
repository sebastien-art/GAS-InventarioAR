/**
 * @OnlyCurrentDoc
 * Exporta los datos filtrados desde "Naves" al archivo fijo de propuestas.
 * Ajustado para mantener Hipervínculos (Ficha) y Formatos Visuales usando la lógica de hoja temporal.
 */
function exportarFilasVisiblesAFijo() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaOrigen = ss.getSheetByName("Naves");
  if (!hojaOrigen) return ui.alert("❌ No se encontró la hoja 'Naves'.");

  const filtro = hojaOrigen.getFilter();
  if (!filtro) return ui.alert("❌ Aplica un filtro antes de exportar.");

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

    const rangoTotal = hojaOrigen.getDataRange();
    const tempSheet = ss.insertSheet("TEMP_FILTER_EXPORT");
    tempSheet.hideSheet();

    // Copiamos todo a la temporal para respetar el filtrado visual inicial (Lógica Rápida)
    rangoTotal.copyTo(tempSheet.getRange(1, 1));

    ss.toast("Procesando columnas e hipervínculos...", "Paso 2/3", -1);

    const rangeTemp = tempSheet.getDataRange();
    const allData = rangeTemp.getDisplayValues(); 
    const allFormulas = rangeTemp.getFormulas(); 
    const encabezados = allData[0];

    const indices = columnasDeseadas.map(col => {
      const i = encabezados.indexOf(col);
      if (i === -1) throw new Error(`Falta columna "${col}"`);
      return i;
    });

    const salida = [["ENVIAR"].concat(columnasDeseadas)];

    for (let i = 1; i < allData.length; i++) {
      const filaValores = allData[i];
      const filaFormulas = allFormulas[i];

      if (filaValores.some(v => v !== "" && v !== null)) {
        const nuevaFila = indices.map(ix => {
          const formula = filaFormulas[ix];
          const valorVisual = filaValores[ix];
          return (formula && formula.startsWith("=")) ? formula : (valorVisual || "");
        });
        salida.push([""].concat(nuevaFila));
      }
    }

    ss.deleteSheet(tempSheet);

    if (salida.length <= 1) return ui.alert("⚠️ No hay datos para exportar.");

    ss.toast("Creando en archivo destino...", "Paso 3/3", -1);
    const ssDestino = SpreadsheetApp.openById(ID_DESTINO);

    let nombreHoja = baseName;
    let contador = 1;
    while (ssDestino.getSheetByName(nombreHoja)) nombreHoja = `${baseName} (${contador++})`;

    const hojaNueva = ssDestino.insertSheet(nombreHoja);
    hojaNueva.getRange(1, 1, salida.length, salida[0].length).setValues(salida);

    // 🎨 Formato
    hojaNueva.setFrozenRows(1);
    hojaNueva.setFrozenColumns(1);
    const numColsFinal = salida[0].length;

    hojaNueva.getRange(1, 1, 1, numColsFinal)
      .setBackground("#b6d7a8")
      .setFontWeight("bold")
      .setFontSize(10)
      .setHorizontalAlignment("center");

    // --- FILA DE FECHA AMARILLA BAJO "OPERACIÓN" ---
    const ultimaFilaData = hojaNueva.getLastRow();
    const filaFecha = ultimaFilaData + 1;
    const fechaHoy = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "dd/MM/yyyy");
    
    const celdaAmarilla = hojaNueva.getRange(filaFecha, 3); // Columna 3 es "Operación"
    celdaAmarilla.setValue(fechaHoy)
                 .setBackground("yellow")
                 .setFontWeight("bold")
                 .setHorizontalAlignment("center");

    // === Anchos de columnas ===
    const SMALL = new Set(["ENVIAR", "Intermediario", "Operación", "Ficha", "REF", "Nave"]);
    const headersFinal = salida[0];

    for (let c = 0; c < headersFinal.length; c++) {
      const header = headersFinal[c];
      hojaNueva.setColumnWidth(c + 1, SMALL.has(header) ? 50 : 130);
    }

    actualizarMenu(ssDestino, hojaNueva);

    SpreadsheetApp.flush();
    ss.toast("¡Completado!", "✅", 2);

    ui.alert(`✅ Exportado`, `"${nombreHoja}" creado con éxito.`, ui.ButtonSet.OK);

  } catch (err) {
    const temp = ss.getSheetByName("TEMP_FILTER_EXPORT");
    if (temp) ss.deleteSheet(temp);
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