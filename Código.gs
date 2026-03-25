/**
 * ACTUALIZACIÓN: Importación con formatos visuales (getDisplayValues)
 * Mantiene símbolos de moneda, decimales y formatos de fecha del origen.
 */

/************* CONFIG *************/
const ID_ORIGEN  = "1jHh3SUkVrQtOZPQ2T2iOhcXF_FbjdXivt3uuBYfYHz0"; // BACKOFFICE
const ID_DESTINO = "1TNw7t5_kog5WeSgbVByKQa4kGDNNdnXhCw8ft3SRGe4"; // INVENTARIO IEM

/************* ENCABEZADOS DESTINO *************/
const ENCAB_NAVES = [
  "Fecha","Intermediario","Operación","Ficha","REF","Estado","Zona Principal","Sub Zona",
  "M2 de construcción","M2 de terreno","Asking price /m2","Mantenimiento / m2",
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
  actualizarTerrenos();
  
  const ssDest = SpreadsheetApp.openById(ID_DESTINO);
  const timestamp = Utilities.formatDate(new Date(), "America/Mexico_City", "dd/MM HH:mm");
  ssDest.rename(`INVENTARIO IEM (Actualizado: ${timestamp})`);
  
  SpreadsheetApp.getActive().toast(`✅ Inventario actualizado (${((Date.now()-t0)/1000).toFixed(1)}s)`);
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

/************* COPIA ORDENADA CON FORMATO VISUAL *************/
function copiarHojaSegura_(nombreHoja, headersDestino) {
  const libroOrigen  = SpreadsheetApp.openById(ID_ORIGEN);
  const libroDestino = SpreadsheetApp.openById(ID_DESTINO);
  const hojaOrigen   = libroOrigen.getSheetByName(nombreHoja);
  const hojaDestino  = libroDestino.getSheetByName(nombreHoja);
  
  if (!hojaOrigen || !hojaDestino)
    throw new Error(`❌ Falta hoja ${nombreHoja} en origen o destino`);

  if (hojaOrigen.getFilter()) hojaOrigen.getFilter().remove();

  // OBTENEMOS VALORES VISUALES (Tal cual se ven en la pantalla)
  const rangeOrigen = hojaOrigen.getDataRange();
  const displayValues = rangeOrigen.getDisplayValues(); 
  
  const encabezadosOrigen = displayValues[0].map(h => String(h).trim());
  const bodyDisplay = displayValues.slice(1);

  // Mapa normalizado de encabezados origen para búsqueda rápida
  const mapaOrigen = {};
  encabezadosOrigen.forEach((h, i) => mapaOrigen[normalize_(h)] = i);

  // Índices de columnas según encabezado destino
  const idxs = headersDestino.map(h => {
    const n = normalize_(h);
    return n in mapaOrigen ? mapaOrigen[n] : null;
  });

  // Log de advertencia si faltan columnas
  const noEncontrados = headersDestino.filter((h, i) => idxs[i] === null);
  if (noEncontrados.length)
    Logger.log(`⚠️ [${nombreHoja}] No encontrados: ${noEncontrados.join(", ")}`);

  // CONSTRUCCIÓN DEL CUERPO (Copiando el formato visual de todas las celdas)
  const cuerpoOrdenado = bodyDisplay
    .filter(fila => fila.some(celda => celda !== "")) // Quitar filas vacías
    .map(fila => {
      return idxs.map(idxOriginal => (idxOriginal === null ? "" : fila[idxOriginal]));
    });

  // Lógica de Fill-Down para la columna "Fecha" (si existe en el destino)
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

  // Pegar en destino
  hojaDestino.clearContents();
  hojaDestino.getRange(1, 1, salida.length, headersDestino.length).setValues(salida);

  SpreadsheetApp.getActive().toast(`✅ ${nombreHoja}: ${cuerpoOrdenado.length} filas con formato copiadas`);
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