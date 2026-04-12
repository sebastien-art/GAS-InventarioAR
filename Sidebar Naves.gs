/**
 * Lanza el Sidebar del buscador.
 */
function showNaveSidebar() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('SidebarHTML')
      .setTitle('Buscador de Naves IEM')
      .setWidth(350);
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (e) {
    SpreadsheetApp.getUi().alert("Error al abrir el panel: " + e.toString());
  }
}

/**
 * BUSCA EL LINK DIRECTO EN LA COLUMNA "Ficha" DE LA HOJA LOCAL.
 * Como importas los datos diario, esta función solo extrae el link de la celda.
 */
function obtenerLinkFichaDesdeCelda(ref) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Naves");
  if (!sh) return "";

  const data = sh.getDataRange().getDisplayValues(); 
  const formulas = sh.getDataRange().getFormulas();
  const headers = data[0];
  
  const colRefIdx = headers.indexOf("REF");
  const colFichaIdx = headers.indexOf("Ficha");
  
  if (colRefIdx === -1 || colFichaIdx === -1) return "";

  const target = ref ? ref.toString().trim().toUpperCase() : "";
  if (!target) return "";

  for (let i = 1; i < data.length; i++) {
    if (data[i][colRefIdx].trim().toUpperCase() === target) {
      const formula = formulas[i][colFichaIdx];
      // Extrae la URL de la fórmula =HYPERLINK("URL", "OK")
      if (formula && formula.includes("HYPERLINK")) {
        const matches = formula.match(/"([^"]+)"/);
        return matches ? matches[1] : "";
      }
      return ""; 
    }
  }
  return "";
}

/**
 * BUSCAR DATOS TÉCNICOS EN EL ARCHIVO MAESTRO
 */
function buscarNaveEnServidor(ref) {
  if (!ref) return "Error: No escribiste nada";
  try {
    const id = "1jHh3SUkVrQtOZPQ2T2iOhcXF_FbjdXivt3uuBYfYHz0"; // ID de tu base maestra
    const target = ref.toString().trim().toUpperCase();
    const ss = SpreadsheetApp.openById(id);
    const sh = ss.getSheetByName("Naves");
    const dataRange = sh.getDataRange().getValues();
    const headers = dataRange[0];
    const colRefIdx = headers.indexOf("REF");

    let rowIndex = -1;
    for (let i = 1; i < dataRange.length; i++) {
      if (dataRange[i][colRefIdx] && dataRange[i][colRefIdx].toString().trim().toUpperCase() === target) {
        rowIndex = i;
        break;
      }
    }

    if (rowIndex === -1) return "No se encontró la REF: " + target;
    
    const rowData = dataRange[rowIndex];
const campos = ["Intermediario","Operación","Ficha","REF","Estado","Zona Principal","Sub Zona","Desarrollador","Parque","Nave","M2 de construcción","M2 de terreno","M2 mínimos rentables","Asking price /m2","Mantenimiento / m2","Energía (kVAs)","Disponibilidad","Comentarios","Renta total","Mantenimiento total","Coordenadas","Ubicación","Andenes de carga","Rampas","A piso", "Resistencia de piso (espesor, resistencia tonelada por m2)","Altura libre","Altura máxima","Tipo de construcción","Tipo de techo","% Skylight","Seguridad 24/7","Oficinas (m2 o %)","Moneda del contrato","Año de construcción","Protección contra incendios","Plazo mínimo de contrato","Gas natural","Caseta de seguridad privada","ID de carpeta de fotos"];

    let res = {};
    campos.forEach(c => {
      let idx = headers.indexOf(c);
      if (idx !== -1) {
        let valor = rowData[idx];
        res[c] = (valor !== "" && valor !== null) ? valor : "---";
      }
    });
    return JSON.stringify(res);
  } catch (err) {
    return "Fallo en servidor: " + err.message;
  }
}