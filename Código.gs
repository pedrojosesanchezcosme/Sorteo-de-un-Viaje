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




