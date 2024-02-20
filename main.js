
//Funcion para Generar Alertas, Menus Personalizados, etc.
var ui = SpreadsheetApp.getUi();
function onOpen(){
  ui.createMenu('Menu Personalizado').addItem('Duplicar Los Datos', 'duplicarDatos').addToUi();
}

function vaciarDatosSalida() {
  const sheet = SpreadsheetApp.openById('1pPWE_pS5tNcRabGIIymprfMo2TNWswZKNv5ovOZ95dY');
  const p_datos_creacion_tarjetas = sheet.getSheetByName('Datos_Creacion_Tarjetas');

  const ultimaFila = p_datos_creacion_tarjetas.getLastRow();
  p_datos_creacion_tarjetas.getRange('A2:J' + ultimaFila).clearContent();
}

function duplicarDatos() {
    //Conectar Sheets a AppScript
    const sheet = SpreadsheetApp.openById('1pPWE_pS5tNcRabGIIymprfMo2TNWswZKNv5ovOZ95dY');
    //Conectar Hojas especificas
    const p_respuestas_form_1 = sheet.getSheetByName('Respuestas_Form_1');
    const p_datos_creacion_tarjetas = sheet.getSheetByName('Datos_Creacion_Tarjetas');
    
    const ultimaFila = p_respuestas_form_1.getLastRow();
    
    for (let fila = 2; fila <= ultimaFila; fila++) 
    {
      Logger.log('Fila= ' + fila);
      const fecha = p_respuestas_form_1.getRange('A' + fila).getValue();
      const descripcion = p_respuestas_form_1.getRange('F' + fila).getValue();
      const descripcionDetallada = p_respuestas_form_1.getRange('G' + fila).getValue();
      const tituloTarjeta = p_respuestas_form_1.getRange('I' + fila).getValue() + '' +
                            p_respuestas_form_1.getRange('J' + fila).getValue() + '' +
                            p_respuestas_form_1.getRange('K' + fila).getValue() + '' +
                            p_respuestas_form_1.getRange('L' + fila).getValue() + '' +
                            p_respuestas_form_1.getRange('N' + fila).getValue();
      const nombrePersona = p_respuestas_form_1.getRange('D' + fila).getValue();
      const lugar = p_respuestas_form_1.getRange('C' + fila).getValue();
      const grupoPlanificador = p_respuestas_form_1.getRange('O' + fila).getValue();
      const prioridad = p_respuestas_form_1.getRange('T' + fila).getValue();
      const tipoRiesgo = p_respuestas_form_1.getRange('H' + fila).getValue();
    
      if (fecha !== '') 
      {
        Logger.log("Entre a imprimir")
        p_datos_creacion_tarjetas.appendRow([fecha, descripcion, descripcionDetallada, tituloTarjeta, nombrePersona, lugar, grupoPlanificador, prioridad, tipoRiesgo]);
      }
    }
  }