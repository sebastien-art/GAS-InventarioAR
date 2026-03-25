function onOpen(e) {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('⚙️ Acciones')
.addItem('Exportar filtrados a Propuestas Comerciales', 'exportarFilasVisiblesAFijo')
   .addSeparator() // --- Línea separadora ---    
  .addItem('Sidebar NAVE', 'showNaveSidebar')
    .addToUi();
}
