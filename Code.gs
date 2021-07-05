/***********************Generar Menú****************************** */

function onOpen(e) {
  var startingsheet = SpreadsheetApp.getActiveSpreadsheet();
  SpreadsheetApp.getUi().createMenu("Camionetas")
  .addItem('Ejecutar', 'mostrarBarra')
  .addToUi();
}

//*Barra lateral
function mostrarBarra(){
  var html = HtmlService.createHtmlOutputFromFile('barraLateral')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setTitle("Informes camionetas")
      .setWidth(300);
  SpreadsheetApp.getUi()
      .showSidebar(html);
}

/***************************************************************** /

/*
function tres(){
global("MORELOS","18key47Qg16etOJ5utQLJXfj84juO7SK5");
}
*/



function global(argumentoTerritorio,argumentoIdCarpeta){
 /************************Traer Datos Base Conjunta*******************************************************/  
  
  //traer la hoja de cálculo de donde salen los datos por su id
  var archivoExterno =SpreadsheetApp.openById("1sX-TPywUPCllQV_OhSZlf-h6R6PwwZEtF5lGMo-OVls");
  
  // traer las hojas del archivo externo
  var hojaConjunta= archivoExterno.getSheetByName("base_conjunta");
  var hojaMatch= archivoExterno.getSheetByName("IDs");


  //traer las ultimas filas y columnas base conjunta
  var ultimaFilaConjunta= hojaConjunta.getLastRow();
  var ultimaColumnaConjunta= hojaConjunta.getLastColumn();

  //traer las ultimas filas y columnas IDs
  var ultimaFilaMatch= hojaMatch.getLastRow();
  var ultimaColumnaMatch= hojaMatch.getLastColumn();

  /*************************Hacer Match Filtrado***********************************************************/
  //traer todos los valores
  var arregloMatchCompleto= hojaMatch.getRange(1,1, ultimaFilaMatch,2).getValues();

  // condicionar solo traer de acuerdo a territorio
  var arregloMatchSemi= arregloMatchCompleto.filter(function(item){
  return item[1]==argumentoTerritorio; // Iteracion
  });

  // traer solo la primer columna de ese arreglo semi
  var arregloMatch=[];
  
  for(var z=0;z<= arregloMatchSemi.length-1;z++){
    var unoPorUno= arregloMatchSemi[z][0];
    arregloMatch.push(unoPorUno);
  }



  /***********************************  crear folders y subfolder **************************/
    var folderTraido=DriveApp.getFolderById(argumentoIdCarpeta); 
    var subFolder= folderTraido.createFolder(argumentoTerritorio); //Crear subfolder nombre del subfolder *
    var idSubFolder= subFolder.getId();

  /************************* Hacer filtro de base Conjunta**************************************************** */ 
  var datos_originales= hojaConjunta.getRange(1,1,ultimaFilaConjunta,ultimaColumnaConjunta).getValues();

    for (i=0; i<=arregloMatch.length-1; i++){

      /*hacer el filtro mediante ciertas condiciones*/
      var datos_filtrados= datos_originales.filter(function(item){
      return item[1]==arregloMatch[i]; // Iteracion
      });
      /**/


      /***************************copia del archivo*********************************************** */
        nombreCopia=(datos_filtrados[0][3]);
      
      
        documentoCopiado= DriveApp.getFileById("1HGwuqgbpvKfwJEk6VPuyIuacJcWm8h4g6WxrAboVmDY").makeCopy(nombreCopia);
  
        var idNuevoDocumento = (documentoCopiado.getId());

      /**/

      /*********************filtrar solo las columnas que me interesan slice push******************/ //TUTORIAL
        var nuevoArreglo=[];
        for(var k=0;k<= datos_filtrados.length-1;k++){
        var unoPorUno= datos_filtrados[k].slice(6,11);
        nuevoArreglo.push(unoPorUno);
        }
      /**/

      /************************añadir un dia a la fecha******************************** */

        for (w=0;w<=nuevoArreglo.length-1;w++) {
        columna=0;
        var fechaAnterior= new Date(nuevoArreglo[w][columna]);

        //SUMARLE 24 HORAS
        var milisegundosUnDia = 1000 * 60 * 60 * 24;
        var nuevaFecha = new Date(fechaAnterior.getTime() + milisegundosUnDia);
        
        //Cambiarle formato
        //var nuevaFecha = Utilities.formatDate(nuevaFecha, 'America/Chicago', 'dd/MM/yyyy');

        //reemplazar
        nuevoArreglo[w].splice(0,1,nuevaFecha);
  
        }
        
      /********************************Traer la hoja************************************************* */
        //traer la hoja de cálculo Plantilla por su id
        var archivoPlantilla =SpreadsheetApp.openById(idNuevoDocumento);

        // traer las hojas del archivo Plantilla
        var hojaPlantilla= archivoPlantilla.getSheetByName("Hoja2");
      /**/


      /************************************pegar valores *************************************************/
  
        //Pegar la tabla
        var rangoAPegar= hojaPlantilla.getRange(12,1, nuevoArreglo.length,nuevoArreglo[0].length);
        rangoAPegar.setValues(nuevoArreglo);

        //pegar la marca
        var rangoAPegar= hojaPlantilla.getRange(7,2);
        // arreglo[0][3] [fila] [columna]
        rangoAPegar.setValue(datos_filtrados[0][11]);

        //pegar el territorio
        var rangoAPegar= hojaPlantilla.getRange(7,5);
        // arreglo[0][3] [fila] [columna]
        rangoAPegar.setValue(datos_filtrados[0][4]);
        
        //pegar placas
        var rangoAPegar= hojaPlantilla.getRange(8,2);
        // arreglo[0][3] [fila] [columna]
        rangoAPegar.setValue(datos_filtrados[0][5]);

        //pegar monedero
        var rangoAPegar= hojaPlantilla.getRange(9,2);
        // arreglo[0][3] [fila] [columna]
        rangoAPegar.setValue(datos_filtrados[0][2]);

        //pegar el resguardante
        var rangoAPegar= hojaPlantilla.getRange(29,4);
        // arreglo[0][3] [fila] [columna]
        rangoAPegar.setValue(datos_filtrados[0][3]);


        // hacer la suma de litros y pegarla //TUTORIAL
        var suma = 0;
        for (l=0;l<=nuevoArreglo.length-1;l++) {
        //var suma_total_dos= suma_total_dos+ nuevoArreglo[i][1]
        columna=2;
        var valorASumar= nuevoArreglo[l][columna];
        var suma = suma + valorASumar;
        }

        //pegar la suma de litros
        var rangoAPegar= hojaPlantilla.getRange(24,3);
        rangoAPegar.setValue(suma);


        // hacer la suma del importe y pegarla //TUTORIAL
        var suma = 0;
        for (j=0;j<=nuevoArreglo.length-1;j++) {
        //var suma_total_dos= suma_total_dos+ nuevoArreglo[i][1]
        columna=3;
        var valorASumar= nuevoArreglo[j][columna];
        var suma = suma + valorASumar;
        }

        //pegar la suma de importe
        var rangoAPegar= hojaPlantilla.getRange(24,4);
        rangoAPegar.setValue(suma);
      
      /*************************************************Mover archivo */


      //var archivo = DriveApp.getFileById("1ITt4o6ePYun2-iyezxQQeIaDDNJcVuWeMAEJ_dw1vAE"); //mover archivo
      DriveApp.getFolderById(idSubFolder).addFile(documentoCopiado);

      /**/

      /***************************************PONER LOS DATOS EN LA HOJA */

      var libro =SpreadsheetApp.getActive();
      var hojaDeTrabajo= libro.getSheetByName('Sheet1');
      hojaDeTrabajo.appendRow([nombreCopia,datos_filtrados[0][4],arregloMatch[i],new Date()]);

    } //aquì termina el for

imprimir();

}// aqui termina la funcion global


function imprimir(){
	Browser.msgBox("Fin de la función");

}


/*******************************funcion  enviar a correo ****************************/

function enviarCorreo(argumentoCarpetaEnviar,argumentoMail) {
   var folder = DriveApp.getFolderById(argumentoCarpetaEnviar);
   var contents = folder.getFiles();
   

   var contador = 0;
   var file;


   var nuevoArreglo=[]
   /*Traer los ID*/

   while (contents.hasNext()) {
    var file = contents.next();
    contador++;

       data = [
            file.getName(),
            file.getId(),
        ];

        /*imprimir los ID*/
        
        //console.log(data[1]);
        //nuevoArreglo.push(data[1]);
        
        var archivo1 = DriveApp.getFileById(data[1]);/*compartido*/
        nuevoArreglo.push(archivo1);
    };

    //console.log(nuevoArreglo);

    GmailApp.sendEmail(argumentoMail, "Asunto", "mensaje", {attachments:nuevoArreglo});

   };












