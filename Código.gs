//Esta función es la encargada de enviar los correos a cada concursante una vez enviada la solicitud
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
  }else if (idioma == "Aleman") {
    cuerpo = "Hallo " + nombre + " este es un correo que confirma que tu solicitud nos ha llegado con exito.";
  }else {cuerpo = "Ciao" + nombre + " este es un correo que confirma que tu solicitud nos ha llegado con exito.";
  }
  
  MailApp.sendEmail(email, asunto, cuerpo);
}

//Con esta función borramos todos los concursates
function borrar(){
  var sps = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = sps.getSheets();
  var datos = sheet[0].getDataRange().getValues();
  var ultimafila = sheet[0].getLastRow();
  var total = ultimafila - 1;
  
  Logger.log("Valor de ultimafila es" + ultimafila);
  
  //Este if es el encargado de que en el caso de que no haya concursantes indique un mensaje de que no hay concursantes para borrar.
  //Además de que si la condicio es que siempre el valor de ultima fila va a ser 1 minimo debido a que es la fila que empleamos para nombrar.
  if(ultimafila <= 1){
    Browser.msgBox("No hay concursantes para borrar.");
  }else{
    var rangoborrar = sheet[0].getRange(2, 1, total, 6);
    rangoborrar.clear();
  }  
}

//Con esta función bordeamos todas las celdas de los concursantes(la funcion if realiza los mismo en todos las funciones)
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

//Con esta función eliminamos los bordes
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

//Con esta función elegimos al ganador
function elegir(){
  var sps = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = sps.getSheets();
  var datos = sheet[0].getDataRange().getValues();
  var ultimafila = sheet[0].getLastRow();
  var total = ultimafila;
  
  
  if(ultimafila <= 1){
    Browser.msgBox("No hay concursantes para elegir.");
  }else{
  
    //Con esta repetitiva evitamos que uno de los resultados posible sea la primera fila, que es donde se situa la descripcion de las columnas
    do{
      var winner = Math.floor(Math.random() * total + 1);
    }while (winner == 1);    
    
    var nombre = sheet[0].getRange(winner, 2).getValue();
    var apellidos = sheet[0].getRange(winner, 3).getValue();
    var email = sheet[0].getRange(winner, 4).getValue();
    var destino = sheet[0].getRange(winner, 5).getValue();
    var ganador = sheet[0].getRange(2, 8);
    
    
    //Con este comando ponemos el nombre del ganador en una fila aaparte 
    sheet[0].getRange(2, 8).setValue(nombre);
    
    //Con esta serie de comando pasamos las caracteristicas de el ganador a una hoja aparte para claridad de los datos del ganador.
    sheet[1].getRange(2, 1).setValue(nombre);
    sheet[1].getRange(2, 2).setValue(apellidos);
    sheet[1].getRange(2, 3).setValue(email);
    sheet[1].getRange(2, 4).setValue(destino);
    
    //Con este comando mostramos un mensaje por pantalla en el que indicamos el nombre y los apellidos del ganador.
    Browser.msgBox("El ganador es " + nombre+ " " + apellidos);
    
    //Aqui empleamos el setBorder para bordear la celda en la que se encuentra el nombre que previamente hemos elegido.
    ganador.setBorder(true, true, true, true, false, false);
  }
}

//Esta función es la encargada de eliminar el ganador para volver a elegirlo
function eliminarganador(){
  var sps = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = sps.getSheets();
  var datos = sheet[0].getDataRange().getValues();
  var ganador = sheet[0].getRange(2,8);
  
  ganador.clear();
}

//Con la funcion opOpen creamos el menu donde estan todas las funciones, por si no queremos usar los botones.
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  
  ui.createMenu('Opciones avanzadas')
      .addItem('Elegir ganador', 'elegir')
      .addItem('Eliminar ganador', 'eliminarganador')
      .addItem('Eliminar los concursantes', 'borrar')
      .addItem('Bordear concursantes', 'bordear')  
      .addItem('Quitar Bordes concursantes', 'desbordear')  
      .addToUi();
}




















