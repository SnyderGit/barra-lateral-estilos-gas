var estilos_sheet = PropertiesService.getDocumentProperties();

function onOpen() {

  SpreadsheetApp.getUi().createMenu('Aulaenlanube')
    .addItem('Mostrar barra lateral','mostrarBarralateral')
    .addToUi();
  
}

function mostrarBarralateral()
{
  var barra = HtmlService.createHtmlOutputFromFile('BarraLateral').setTitle('Barra lateral Aulaenlanube');
  SpreadsheetApp.getUi().showSidebar(barra);
}

function aplicarEstilo1()
{
  var hojaActual = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var celdas = hojaActual.getActiveRange();

  celdas.setFontColor(estilos_sheet.getProperty('color'))
        .setBackground(estilos_sheet.getProperty('colorFondo'))
        .setFontSize(estilos_sheet.getProperty('size'));
}

function guardarEstilo1()
{
  var celda = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveCell();

  estilos_sheet.setProperty('color', celda.getFontColor())
               .setProperty('colorFondo', celda.getBackground())
               .setProperty('size', celda.getFontSize()+'');

  return {  colorFondo: estilos_sheet.getProperty('colorFondo'),
            colorLetra: estilos_sheet.getProperty('color')};

}

function borrarEstilos()
{
  SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveRange().clear({formatOnly: true});
}

function borrarTodo()
{
  SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveRange().clear();
}

























