/**
 * ACTUALIZACIÓN: Importación con formatos visuales e Hipervínculos.
 * Mantiene símbolos de moneda, decimales, fechas y links (HYPERLINK).
 */

/************* CONFIG *************/
const ID_ORIGEN  = "1jHh3SUkVrQtOZPQ2T2iOhcXF_FbjdXivt3uuBYfYHz0"; // BACKOFFICE
const ID_DESTINO = "1TNw7t5_kog5WeSgbVByKQa4kGDNNdnXhCw8ft3SRGe4"; // INVENTARIO IEM

/************* ENCABEZADOS DESTINO *************/
const ENCAB_NAVES = [
  "Fecha","Intermediario","Operación","Ficha","REF","Estado","Zona Principal","Sub Zona",
  "M2 de construcción","M2 de terreno","Asking price /m2", "Precio total", "Mantenimiento / m2",
  "Desarrollador","Parque","Nave","Disponibilidad","Comentarios","Energía (kVAs)",
  "Renta total","Mantenimiento total","Coordenadas","Ubicación","Rango m2",
  "M2 mínimos rentables","Plazo mínimo de contrato","Comisión","Link","Altura libre",
  "Altura máxima","Patio de maniobras","Crossdock","Espacio entre columnas / tamaño de bahía",
  "Oficinas (m2 o %)","Andenes de carga","Seguridad 24/7","Costo de los derechos de kVAs",
  "% Skylight","Iluminación artificial","Certificación","Extracción / cambios de aire",
  "Protección contra incendios","Moneda del contrato","Cajones de estacionamiento",
  "Suministro de agua","Gas natural","Instalación de grúa","Rampas","Andenes",
  "Resistencia de piso (espesor, resistencia tonelada por m2)","Tipo de construcción",
  "Tipo de techo","Aire acondicionado oficinas","Fibra optica / telecomunicaciones",
  "Caseta de seguridad privada","Dimensión del edificio","Año de construcción",
  "Altura de andén","Estacionamiento parar trailers","Espacio en mezzanine"
];

const ENCAB_TERRENOS = [
  "INTERMEDIARIO","OPERACION","REF","ESTADO","ZONA","HECTÁREAS","PRECIO M2",
  "INFORMACIÓN","USD 1","USD 2","MXN","MXN 2","CONCATENAR TODOS PRECIOS",
  "USOSUELOCAT","PROPIEDAD","DESCRIPCION ESQUIVOCADA","USO SUELO","LEGAL",
  "SERVICIOS","UBICACIÓN","COORDENADAS","ANEXOS","COMISIÓN %","PRECIO TOTAL",
  "Contacto","Factibilidad de energia","Detalles de aportación"
];

/************* FUNCIONES PRINCIPALES *************/
function copiarDatosInventario() {
  const t0 = Date.now();

  actualizarNaves();
  SpreadsheetApp.flush();

  actualizarTerrenos();
  SpreadsheetApp.flush();

  const ssDest = SpreadsheetApp.openById(ID_DESTINO);
  const timestamp = Utilities.formatDate(new Date(), "America/Mexico_City", "dd/MM HH:mm");
  ssDest.rename(`INVENTARIO IEM (Actualizado: ${timestamp})`);

  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(`✅ Inventario actualizado (${((Date.now()-t0)/1000).toFixed(1)}s)`);
  } catch(e) {
    console.log(`Inventario actualizado en ${((Date.now()-t0)/1000).toFixed(1)}s`);
  }
}

function actualizarNaves()    { copiarHojaSegura_("Naves", ENCAB_NAVES); }
function actualizarTerrenos() { copiarHojaSegura_("Terrenos", ENCAB_TERRENOS); }

function probarConexion() {
  try {
    SpreadsheetApp.openById(ID_ORIGEN);
    SpreadsheetApp.openById(ID_DESTINO);
    SpreadsheetApp.getActive().toast("✅ Conexión OK");
  } catch(e) {
    SpreadsheetApp.getActive().toast("❌ Error de conexión: " + e.message);
  }
}

/************* COPIA ORDENADA CON FORMATO VISUAL E HIPERVÍNCULOS *************/
function copiarHojaSegura_(nombreHoja, headersDestino) {
  const libroOrigen  = SpreadsheetApp.openById(ID_ORIGEN);
  SpreadsheetApp.flush();
  Utilities.sleep(3000); // ✅ fuerza sincronización antes de leer

  const libroDestino = SpreadsheetApp.openById(ID_DESTINO);
  const hojaOrigen   = libroOrigen.getSheetByName(nombreHoja);
  const hojaDestino  = libroDestino.getSheetByName(nombreHoja);

  if (!hojaOrigen || !hojaDestino)
    throw new Error(`❌ Falta hoja ${nombreHoja}`);

  if (hojaOrigen.getFilter()) hojaOrigen.getFilter().remove();

  const rangeOrigen   = hojaOrigen.getDataRange();
  const displayValues = rangeOrigen.getDisplayValues();
  const rawValues     = rangeOrigen.getValues();
  const richTexts     = rangeOrigen.getRichTextValues(); // ✅ solo para Ficha

  // ✅ Fórmulas SOLO de la columna Ficha, no de todo el rango
  const encabezadosOrigen = displayValues[0].map(h => String(h).trim());
  const idxFichaOrigen = encabezadosOrigen.findIndex(h => normalize_(h) === normalize_("Ficha"));
  
  let formulasFicha = []; // array de strings, una por fila (sin encabezado)
  if (idxFichaOrigen !== -1) {
    const colFicha = idxFichaOrigen + 1; // getRange usa base 1
    const numFilas = rangeOrigen.getNumRows() - 1;
    formulasFicha = hojaOrigen
      .getRange(2, colFicha, numFilas, 1)
      .getFormulas()
      .map(r => r[0]);
  }

  console.log(`[${nombreHoja}] Rango: ${rangeOrigen.getNumRows()} filas × ${rangeOrigen.getNumColumns()} cols`);
  console.log(`[${nombreHoja}] idx Ficha en origen: ${idxFichaOrigen}`);

  const bodyDisplay = displayValues.slice(1);
  const bodyRaw     = rawValues.slice(1);
  const bodyRich    = richTexts.slice(1);

  const mapaOrigen = {};
  encabezadosOrigen.forEach((h, i) => mapaOrigen[normalize_(h)] = i);

  const idxs = headersDestino.map(h => {
    const n = normalize_(h);
    return n in mapaOrigen ? mapaOrigen[n] : null;
  });

  const cuerpoOrdenado = bodyDisplay
    .filter((fila, i) => {
      return fila.some(c => c !== "") ||
             bodyRaw[i].some(c => c !== "" && c !== false);
    })
    .map((fila, i) => {
      return idxs.map(idxOriginal => {
        if (idxOriginal === null) return "";

        // ✅ Solo Ficha recibe tratamiento de fórmula/hyperlink
        if (idxOriginal === idxFichaOrigen) {

          // 1. Fórmula =HYPERLINK en Ficha
          const formula = formulasFicha[i] || "";
          if (formula.startsWith("=")) return formula;

          // 2. Rich Text con URL en Ficha
          try {
            const rt = bodyRich[i][idxOriginal];
            if (rt) {
              const runs = rt.getRuns();
              for (const run of runs) {
                const url = run.getLinkUrl();
                if (url) {
                  const texto = rt.getText() || url;
                  return `=HYPERLINK("${url}","${texto.replace(/"/g,'')}")`;
                }
              }
            }
          } catch(e) {}

          // 3. Si no hay fórmula ni link, valor display
          return fila[idxOriginal];
        }

        // ✅ Todas las demás columnas: solo valor display (sin fórmulas)
        const rawVal = bodyRaw[i][idxOriginal];
        if (rawVal === true)  return "OK";
        if (rawVal === false) return "";

        return fila[idxOriginal];
      });
    });

  console.log(`[${nombreHoja}] Filas a escribir: ${cuerpoOrdenado.length}`);

  // Fill-Down Fecha
  const idxDestFecha = headersDestino.findIndex(h => normalize_(h) === "fecha");
  if (idxDestFecha !== -1) {
    let ultimaFechaValida = "";
    for (let i = 0; i < cuerpoOrdenado.length; i++) {
      if (cuerpoOrdenado[i][idxDestFecha] !== "") {
        ultimaFechaValida = cuerpoOrdenado[i][idxDestFecha];
      } else {
        cuerpoOrdenado[i][idxDestFecha] = ultimaFechaValida;
      }
    }
  }

  const salida = [headersDestino, ...cuerpoOrdenado];
  hojaDestino.clearContents();
  hojaDestino.getRange(1, 1, salida.length, headersDestino.length).setValues(salida);

  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(`✅ ${nombreHoja}: ${cuerpoOrdenado.length} filas copiadas`);
  } catch(e) {
    console.log(`${nombreHoja}: ${cuerpoOrdenado.length} filas copiadas`);
  }
}

/************* HELPERS *************/
function normalize_(s) {
  return String(s || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ")
    .trim();
}