function emailOnFormSubmit (e){
  var nombre = e.values[1];
  var email = e.values[3];
  var idioma = e.values[5];
  var asunto = "Sorteo viajes";
  var cuerpo;
    
  if (idioma == "Español") {
    cuerpo = "Hola " + nombre + " este es un correo que confirma que tu solicitud nos ha llegado con exito.";
  }else if (idioma == "Ingles") {
    cuerpo = "Hello " + nombre + " este es un correo que confirma que tu solicitud nos ha llegado con exito.";
  }else if (idioma == "Frances") {
    cuerpo = "Bonjour " + nombre + " este es un correo que confirma que tu solicitud nos ha llegado con exito.";
  }else {cuerpo = "Hallo " + nombre + " este es un correo que confirma que tu solicitud nos ha llegado con exito.";
  }
  
  MailApp.sendEmail(email, asunto, cuerpo);
}

function borrar(){
  var sps = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = sps.getSheets();
  var datos = sheet[0].getDataRange().getValues();
  var ultimafila = sheet[0].getLastRow();
  var total = ultimafila - 1;
  
  Logger.log("Valor de ultimafila es" + ultimafila);
  
  if(ultimafila <= 1){
    Browser.msgBox("No hay concursantes para borrar.");
  }else{
    var rangoborrar = sheet[0].getRange(2, 1, total, 6);
    rangoborrar.clear();
  }  
}

function bordear(){
  var sps = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = sps.getSheets();
  var datos = sheet[0].getDataRange().getValues();
  var ultimafila = sheet[0].getLastRow();
  var total = ultimafila - 1;
   
  if(ultimafila <= 1){
    Browser.msgBox("No hay concursantes para establecerles un borde.");
  }else{
    var rangoborde = sheet[0].getRange(2, 1, total, 6);  
    rangoborde.setBorder(true, true, true, true, true, true);
  }    
}

function desbordear(){
  var sps = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = sps.getSheets();
  var datos = sheet[0].getDataRange().getValues();
  var ultimafila = sheet[0].getLastRow();
  var total = ultimafila - 1;
      
  if(ultimafila <= 1){
    Browser.msgBox("No hay concursantes quitarles el borde.");
  }else{
    var rangoborde = sheet[0].getRange(2, 1, total, 6);  
    rangoborde.setBorder(false, false, false, false, false, false);
  }
}

function elegir(){
  var sps = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = sps.getSheets();
  var datos = sheet[0].getDataRange().getValues();
  var ultimafila = sheet[0].getLastRow();
  var total = ultimafila;
  
  if(ultimafila <= 1){
    Browser.msgBox("No hay concursantes para elegir.");
  }else{
    do{
      var winner = Math.floor(Math.random() * total + 1);
    }while (winner == 1);    
    
    var nombre = sheet[0].getRange(winner, 2).getValue();
    var apellidos = sheet[0].getRange(winner, 3).getValue();
    var email = sheet[0].getRange(winner, 4).getValue();
    var destino = sheet[0].getRange(winner, 5).getValue();
    var ganador = sheet[0].getRange(2, 8);
    
    sheet[0].getRange(2, 8).setValue(nombre);
    sheet[1].getRange(2, 1).setValue(nombre);
    sheet[1].getRange(2, 2).setValue(apellidos);
    sheet[1].getRange(2, 3).setValue(email);
    sheet[1].getRange(2, 4).setValue(destino);
    Browser.msgBox("El ganador es " + nombre+ " " + apellidos);
    ganador.setBorder(true, true, true, true, false, false);
  }
}

function eliminarganador(){
  var sps = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = sps.getSheets();
  var datos = sheet[0].getDataRange().getValues();
  var ganador = sheet[0].getRange(2,8);
  
  ganador.clear();
}

function grafico() {
  var sps = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = sps.getSheets();
  var datos = sheet[0].getDataRange().getValues();
  var ultimafila = sheet[0].getLastRow();
  var total = ultimafila;
  
  var chart = datos.newChart()
  .setChartType(Charts.ChartType.BAR)
  .addRange(datos.getRange(2, 5, datos, 5))
  .setPosition(2, 9, 0, 0)
  .build();
  
  sheet.insertChart(chart);
}

function onOpen() {
  grafico();
  var ui = SpreadsheetApp.getUi();
  
  ui.createMenu('Opciones avanzadas')
      .addItem('Elegir ganador', 'elegir')
      .addItem('Eliminar ganador', 'eliminarganador')
      .addItem('Eliminar los concursantes', 'borrar')
      .addItem('Bordear concursantes', 'bordear')  
      .addItem('Quitar Bordes concursantes', 'desbordear')  
      .addToUi();
}




















