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
    "Asking price /m2", "Precio total", "Mantenimiento / m2", "Disponibilidad","Energía (kVAs)", "Comentarios", "Coordenadas", "Ubicación", "Altura libre", "Altura máxima"
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
    const allRichTexts = rangeTemp.getRichTextValues();
    const encabezados = allData[0];

    const indices = columnasDeseadas.map(col => {
      const i = encabezados.indexOf(col);
      if (i === -1) throw new Error(`Falta columna "${col}"`);
      return i;
    });

    const salida = [["ENVIAR"].concat(columnasDeseadas)];
    const richTextSalida = [[null].concat(columnasDeseadas.map(() => null))];

    for (let i = 1; i < allData.length; i++) {
      const filaValores = allData[i];
      const filaFormulas = allFormulas[i];
      const filaRichTexts = allRichTexts[i];

      if (filaValores.some(v => v !== "" && v !== null)) {
        const nuevaFilaRichText = [null];
        const nuevaFila = indices.map(ix => {
          const formula = filaFormulas[ix];
          const valorVisual = filaValores[ix];
          const richTextVal = filaRichTexts[ix];
          
          if (formula && formula.startsWith("=")) {
            nuevaFilaRichText.push(null);
            return formula;
          } else if (richTextVal && richTextVal.getLinkUrl()) {
            nuevaFilaRichText.push(richTextVal);
            return valorVisual || "";
          } else {
            nuevaFilaRichText.push(null);
            return (valorVisual || "");
          }
        });
        salida.push([""].concat(nuevaFila));
        richTextSalida.push(nuevaFilaRichText);
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
    
    for (let r = 0; r < salida.length; r++) {
      for (let c = 0; c < salida[r].length; c++) {
        const celdaDestino = hojaNueva.getRange(r + 1, c + 1);
        const valor = salida[r][c];
        const rt = richTextSalida[r][c];
        
        if (valor && valor.toString().startsWith("=")) {
          celdaDestino.setFormula(valor);
        } else if (rt) {
          celdaDestino.setRichTextValue(rt);
        } else {
          celdaDestino.setValue(valor);
        }
      }
    }

    // 🎨 Formato
    hojaNueva.setFrozenRows(1);
    hojaNueva.setFrozenColumns(1);
    const numColsFinal = salida[0].length;

    hojaNueva.getRange(1, 1, 1, numColsFinal)
      .setBackground("#b6d7a8")
      .setFontWeight("bold")
      .setFontSize(10)
      .setHorizontalAlignment("center");

// --- FECHA DE CREACIÓN (VISIBLE + OCULTA) ---
const ultimaFilaData = hojaNueva.getLastRow();
const filaFecha = ultimaFilaData + 1;

// Fecha completa con hora (para el archivado)
const fechaCreacion = new Date();

// Visible para el usuario (igual que hoy)
const fechaVisible = Utilities.formatDate(
  fechaCreacion,
  ss.getSpreadsheetTimeZone(),
  "dd/MM/yyyy"
);

// Valor interno con fecha y hora
const fechaInterna = Utilities.formatDate(
  fechaCreacion,
  ss.getSpreadsheetTimeZone(),
  "yyyy-MM-dd HH:mm:ss"
);

// Celda amarilla visible
hojaNueva.getRange(filaFecha, 3)
  .setValue(fechaVisible)
  .setBackground("yellow")
  .setFontWeight("bold")
  .setHorizontalAlignment("center");

// Guardar fecha interna en AA1 (columna oculta)
hojaNueva.getRange("AA1").setValue(fechaInterna);

// Ocultar la columna AA únicamente la primera vez
if (!hojaNueva.isColumnHiddenByUser(27)) {
  hojaNueva.hideColumns(27);
}

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
  const hojaMenu = ssDestino.getSheetByName("Menú");
  if (!hojaMenu) return;

  const ultimaFila = hojaMenu.getLastRow();
  if (ultimaFila < 1) return;

  // 1. Leer encabezados existentes en la Fila 1
  const ultimaColumna = hojaMenu.getLastColumn();
  const encabezados = hojaMenu.getRange(1, 1, 1, ultimaColumna).getValues()[0];

  // 2. Buscar dinámicamente las posiciones de las columnas por encabezado
  const idxNombre = encabezados.indexOf("Propuestas/Reportes");
  const idxNavegar = encabezados.indexOf("Navegar");
  const idxFecha = encabezados.indexOf("Fecha de creación");

  const filaNueva = ultimaFila + 1;
  const nombre = hojaNueva.getName();

  // 3. Escribir Nombre si existe la columna "Propuestas/Reportes" (o columna 1 si no la halla)
  const colNombre = idxNombre !== -1 ? idxNombre + 1 : 1;
  hojaMenu.getRange(filaNueva, colNombre).setValue(nombre);

  // 4. Escribir Enlace si existe la columna "Navegar" (o columna 2 si no la halla)
  const colNavegar = idxNavegar !== -1 ? idxNavegar + 1 : 2;
  hojaMenu.getRange(filaNueva, colNavegar)
    .setFormula(`=HYPERLINK("#gid=${hojaNueva.getSheetId()}", "VER DATOS")`)
    .setBackground("#007bff")
    .setFontColor("white")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  // 5. Escribir Fecha si existe la columna "Fecha de creación"
  if (idxFecha !== -1) {
    const fechaHoy = Utilities.formatDate(
      new Date(),
      ssDestino.getSpreadsheetTimeZone(),
      "dd/MM/yyyy"
    );

    hojaMenu.getRange(filaNueva, idxFecha + 1)
      .setValue(fechaHoy)
      .setHorizontalAlignment("center");
  }
}