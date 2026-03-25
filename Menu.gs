/************* MENÚ *************/
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("⚙️ Acciones")
    .addItem("Actualizar inventario ahora", "copiarDatosInventario")
    .addSeparator()
    .addItem("Actualizar solo Naves", "actualizarNaves")
    .addItem("Actualizar solo Terrenos", "actualizarTerrenos")
    .addItem("Probar conexión", "probarConexion")
    .addSeparator() // --- Línea separadora ---
    .addItem("Exportar a Propuestas Comerciales", "exportarFilasVisiblesAFijo")
    .addSeparator() // --- Línea separadora ---    
    .addItem('Sidebar NAVE', 'showNaveSidebar')
   .addToUi();
}
