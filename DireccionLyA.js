function copiar(){
  var archivo = SpreadsheetApp.getActiveSpreadsheet();

  var copyss = archivo.getSheetByName("Datos");
  var copyApostilla = copyss.getRange('M2');
  var copyLegalizacion = copyss.getRange('N2');
  var refApostilla = copyss.getRange('M3');
  var refLegalizacion = copyss.getRange('N3');

  var ss = archivo.getSheetByName("PersonaTramite");
  var activa = ss.getActiveCell();
  var valor = activa.getValue();
  var filaActiva = activa.getRow();
  var colActiva = activa.getColumn();


  if(filaActiva>=2 && colActiva==1 && archivo.getActiveSheet().getName() == "PersonaTramite"){
    switch(valor){
      case "APOSTILLA":
        refApostilla.copyTo(activa.offset(0,1));
        copyApostilla.copyTo(activa.offset(0,15));
      break;

      case "LEGALIZACION":
        refLegalizacion.copyTo(activa.offset(0,1));
        copyLegalizacion.copyTo(activa.offset(0,15));
      break;

      default:
        activa.offset(0,1).setValue("ERROR");
        activa.offset(0,15).setValue("ERROR");
      break
    }
  }

  
}

function personatramite(){
  var archivo = SpreadsheetApp.getActiveSpreadsheet();
  var ss = archivo.getSheetByName("PersonaTramite");
  var activa = ss.getActiveCell();
  var valor = activa.getValue();
  var filaActiva = activa.getRow();
  var colActiva = activa.getColumn();



  if(filaActiva>=2 && colActiva==3 && archivo.getActiveSheet().getName() =="PersonaTramite"){
      activa.offset(0,14).setValue(new Date());
    
  }

  if(filaActiva>=2 && colActiva==10 && archivo.getActiveSheet().getName() =="PersonaTramite"){
      activa.offset(0,15).setValue(new Date());

  }
}

function listaCargoFuncionario(){
  var archivo = SpreadsheetApp.getActiveSpreadsheet();
  var ss = archivo.getSheetByName("LEGALIZACION");
  var ssf = archivo.getSheetByName("FUNCIONARIOS");
  var arregloFuncionarios = ssf.getDataRange().getValues();

  
  var func = ss.getActiveCell();
  var valor = func.getValue();
  var filaActiva = func.getRow();
  var colActiva = func.getColumn();

  if(filaActiva>=2 && colActiva==8 && archivo.getActiveSheet().getName() =="LEGALIZACION"){
        var arregloCargo =  new Array();
        var j=0;
        arregloFuncionarios.forEach(fila=>{
            var nombre = fila[1];
            var cargo = fila[2];
            var area =  fila[3];
            console.log(nombre);

            if(valor == nombre){
              arregloCargo[j] = cargo;
              console.log("Cargo " +j+" :"+arregloCargo[j]);
              j++;
            }
           
        })
      }
      
    var reglaValid =SpreadsheetApp.newDataValidation().requireValueInList(arregloCargo,true).build();
    var celda =  SpreadsheetApp.getActive().getActiveRange().offset(0,1).clearDataValidations().clearContent();
    celda.setDataValidation(reglaValid);  
}

function horaModificacion() {
  var archivo = SpreadsheetApp.getActiveSpreadsheet();
  var ssa = archivo.getSheetByName("BoletasApostillas");
  var ssl = archivo.getSheetByName("BoletasLegalizaciones");
  var activa = ssa.getActiveCell();
  var actival = ssl.getActiveCell();
  var valor = activa.getValue();
  var valorl = actival.getValue();
  var filaActiva = activa.getRow();
  var filaActival = actival.getRow();
  var colActiva = activa.getColumn();
  var colActival = actival.getColumn();

  if (filaActiva >= 2 && colActiva == 3 && archivo.getActiveSheet().getName() == "BoletasApostillas") {
    if (activa.offset(0, 2).getValue() == "") {
      activa.offset(0, 2).setValue(new Date());
    }
    else {
      activa.offset(0, 4).setValue(new Date());
    }

  }

  if (filaActiva >= 2 && colActiva == 3 && archivo.getActiveSheet().getName() == "BoletasLegalizaciones") {
    if (activa.offset(0, 2).getValue() == "") {
      activa.offset(0, 2).setValue(new Date());
    }
    else {
      activa.offset(0, 4).setValue(new Date());
    }

  }
}


function recibido(){
   var archivo = SpreadsheetApp.getActiveSpreadsheet();
  var ssa = archivo.getSheetByName("APOSTILLA");
  var ssl = archivo.getSheetByName("LEGALIZACION");

  var activa = ssa.getActiveCell();
  var actival = ssl.getActiveCell();

  var valor = activa.getValue();
  var valorl = actival.getValue();

  var filaActiva = activa.getRow();
  var filaActival = actival.getRow();
  var colActiva = activa.getColumn();
  var colActival = actival.getColumn();

  var celdanom = ssl.getRange(2,13);
  var celdanoma = ssa.getRange(2,13);

  for(var i=2;i<100000;i=i+1){
    var celdanom = ssl.getRange(i,13);
    var celdanoma = ssa.getRange(i,13);
    var nombre = celdanom.getValue();
    var nombrea = celdanoma.getValue();
    if(nombre!=""){
      if(ssl.getRange(i,16).getValue()==""){
          ssl.getRange(i,15).setValue("RECIBIDO");
          ssl.getRange(i,16).setValue(new Date());
      } 
    }
    if(nombrea!=""){
      if(ssa.getRange(i,16).getValue()==""){
          ssa.getRange(i,15).setValue("RECIBIDO");
          ssa.getRange(i,16).setValue(new Date());
      } 
    }
  }

  
}

function enproceso(){
  var archivo = SpreadsheetApp.getActiveSpreadsheet();
  var ssa = archivo.getSheetByName("APOSTILLA");
  var ssl = archivo.getSheetByName("LEGALIZACION");
  var activa = ssa.getActiveCell();
  var actival = ssl.getActiveCell();
  var valor = activa.getValue();
  var valorl = actival.getValue();
  var filaActiva = activa.getRow();
  var filaActival = actival.getRow();
  var colActiva = activa.getColumn();
  var colActival = actival.getColumn();

  if(colActival==3 && valorl!="" && actival.offset(0, 12).getValue()=="RECIBIDO"){
    actival.offset(0, 12).setValue("EN PROCESO");
    actival.offset(0, 14).setValue(new Date());

  }
  if(colActiva==3 && valor!="" && activa.offset(0, 12).getValue()=="RECIBIDO"){
    activa.offset(0, 12).setValue("EN PROCESO");
    activa.offset(0, 14).setValue(new Date());

  }
}

function entregado(){
  var archivo = SpreadsheetApp.getActiveSpreadsheet();
  var ssa = archivo.getSheetByName("APOSTILLA");
  var ssl = archivo.getSheetByName("LEGALIZACION");

  var celdanom = ssl.getRange(2,14);
  var celdanoma = ssa.getRange(2,14);
  

  var colTrami = celdanom.getColumn;
  for(var i=2;i<100000;i=i+1){
    var celdanom = ssl.getRange(i,14);
    var celdanoma = ssa.getRange(i,14);
   

    var nombre = celdanom.getValue();
    var nombrea = celdanoma.getValue();
  

    if(nombre!=""){
      if(ssl.getRange(i,18).getValue()==""){
          ssl.getRange(i,15).setValue("ENTREGADO");
          ssl.getRange(i,18).setValue(new Date());
      } 
    }
    if(nombrea!=""){
      if(ssa.getRange(i,18).getValue()==""){
          ssa.getRange(i,15).setValue("ENTREGADO");
          ssa.getRange(i,18).setValue(new Date());
      } 
    }
  }

}
