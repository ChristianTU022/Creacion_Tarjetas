
//Funcion para Generar Alertas, Menus Personalizados, etc.
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Menú Personalizado')
    .addItem('Duplicar Datos', 'duplicarDatos')
    .addItem('Limpiar Datos de Entrada', 'limpiarDatosEntrada')
    .addItem('Limpiar Datos de Salida', 'limpiarDatosSalida')
    .addToUi();
}

//Funcion que permite Limpiar los datos del formulario sheets la hoja "Datos_Entrada_CT"
function limpiarDatosEntrada() {
  const sheet = SpreadsheetApp.openById('1pPWE_pS5tNcRabGIIymprfMo2TNWswZKNv5ovOZ95dY');
  const p_Datos_Salida_CT = sheet.getSheetByName('Datos_Entrada_CT');
  const ultimaFila = p_Datos_Salida_CT.getLastRow();
  p_Datos_Salida_CT.getRange('A2:AK' + ultimaFila).clearContent();
}
//Funcion que permite Limpiar los datos del formulario sheets la hoja "Datos_Salida_CT"
function limpiarDatosSalida() {
  const sheet = SpreadsheetApp.openById('1pPWE_pS5tNcRabGIIymprfMo2TNWswZKNv5ovOZ95dY');
  const p_Datos_Salida_CT = sheet.getSheetByName('Datos_Salida_CT');

  const ultimaFila = p_Datos_Salida_CT.getLastRow();
  p_Datos_Salida_CT.getRange('A2:K' + ultimaFila).clearContent();
}


function duplicarDatos() {
    //Conectar Sheets a AppScript
    const sheet = SpreadsheetApp.openById('1pPWE_pS5tNcRabGIIymprfMo2TNWswZKNv5ovOZ95dY');
    //Conectar Hojas especificas
    const p_Datos_Entrada_CT = sheet.getSheetByName('Datos_Entrada_CT');
    const p_Datos_Salida_CT = sheet.getSheetByName('Datos_Salida_CT');
    
    const ultimaFila = p_Datos_Entrada_CT.getLastRow();
    
    //Saca Los datos especificos del Archivo/hoja de Entrada al archivo/hoja de salida
    for (let fila = 2; fila <= ultimaFila; fila++) 
    {
      const fecha = p_Datos_Entrada_CT.getRange('A' + fila).getValue();
      const descripcion = p_Datos_Entrada_CT.getRange('F' + fila).getValue();
      const descripcionDetallada = p_Datos_Entrada_CT.getRange('G' + fila).getValue();
      const tituloTarjeta = p_Datos_Entrada_CT.getRange('I' + fila).getValue() + '' +
                            p_Datos_Entrada_CT.getRange('J' + fila).getValue() + '' +
                            p_Datos_Entrada_CT.getRange('K' + fila).getValue() + '' +
                            p_Datos_Entrada_CT.getRange('L' + fila).getValue() + '' +
                            p_Datos_Entrada_CT.getRange('N' + fila).getValue();
      const nombrePersona = p_Datos_Entrada_CT.getRange('D' + fila).getValue();
      const lugar = p_Datos_Entrada_CT.getRange('C' + fila).getValue();
      const grupoPlanificador = p_Datos_Entrada_CT.getRange('O' + fila).getValue();
      const prioridad = p_Datos_Entrada_CT.getRange('T' + fila).getValue();
      const tipoRiesgo = p_Datos_Entrada_CT.getRange('H' + fila).getValue();
       

      if (fecha !== '') 
      {
        p_Datos_Salida_CT.appendRow([fecha, descripcion, descripcionDetallada, tituloTarjeta, nombrePersona, lugar, grupoPlanificador, prioridad, tipoRiesgo]);
      }
      let valor_Cod_Titulo = p_Datos_Salida_CT.getRange('D' + fila).getValue();
      Logger.log(valor_Cod_Titulo + " Valor Tomado");
      if (valor_Cod_Titulo === 'Condición básica' || valor_Cod_Titulo === 'Condiciones básicas') 
      {
        p_Datos_Salida_CT.getRange('K' + fila).setValue('CB');
      } else if (valor_Cod_Titulo === 'Condición insegura') {
        p_Datos_Salida_CT.getRange('K' + fila).setValue('CI');
      } else if (valor_Cod_Titulo === 'Incidente') {
        p_Datos_Salida_CT.getRange('K' + fila).setValue('I');
      } else if (valor_Cod_Titulo === 'Acto inseguro' || valor_Cod_Titulo === 'Actos Inseguros') {
        p_Datos_Salida_CT.getRange('K' + fila).setValue('AI');
      } else if (valor_Cod_Titulo === 'Incidentes ambientales') {
        p_Datos_Salida_CT.getRange('K' + fila).setValue('IA');
      } else if (valor_Cod_Titulo === 'Acto Inseguro ambientales') {
        p_Datos_Salida_CT.getRange('K' + fila).setValue('AIA');
      } else if (valor_Cod_Titulo === 'Defecto') {
        p_Datos_Salida_CT.getRange('K' + fila).setValue('DF');
      } else if (valor_Cod_Titulo === 'Acto y/o comportamiento') {
        p_Datos_Salida_CT.getRange('K' + fila).setValue('AC');
      } else if (valor_Cod_Titulo === 'Condición de operación') {
        p_Datos_Salida_CT.getRange('K' + fila).setValue('CO');
      }    
    }
  }

