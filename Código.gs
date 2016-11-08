/*
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
 
*/