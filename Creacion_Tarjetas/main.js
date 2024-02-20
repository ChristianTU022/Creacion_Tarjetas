
//Funcion para Generar Alertas, Menus Personalizados, etc.
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Menú Personalizado')
    .addItem('Duplicar Datos', 'duplicarDatos')
    .addItem('Limpiar Datos de Entrada', 'confirmarLimpiarDatosEntrada')
    .addItem('Limpiar Datos de Salida', 'confirmarLimpiarDatosSalida')
    .addToUi();
}

function confirmarLimpiarDatosEntrada() {
  const ui = SpreadsheetApp.getUi();
  const respuesta = ui.alert(
    'Confirmación',
    '¿Está seguro de que desea limpiar los datos de entrada? Este proceso limpiará cualquier tipo de dato.',
    ui.ButtonSet.YES_NO);

  if (respuesta == ui.Button.YES) {
    limpiarDatosEntrada();
  }
}

function confirmarLimpiarDatosSalida() {
  const ui = SpreadsheetApp.getUi();
  const respuesta = ui.alert(
    'Confirmación',
    '¿Está seguro de que desea limpiar los datos de Salida? Este proceso limpiará cualquier tipo de dato.',
    ui.ButtonSet.YES_NO);

  if (respuesta == ui.Button.YES) {
    limpiarDatosSalida();
  }
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

      //Sacar Codigo o iniciales del Titulo de la Tarjeta
      let valor_Cod_Titulo = p_Datos_Salida_CT.getRange('D' + fila).getValue();
      Logger.log(valor_Cod_Titulo + " Valor Tomado");

      switch (valor_Cod_Titulo) {
        case 'Condición básica':
        case 'Condiciones básicas':
          p_Datos_Salida_CT.getRange('K' + fila).setValue('CB');
          break;
        case 'Condición insegura':
          p_Datos_Salida_CT.getRange('K' + fila).setValue('CI');
          break;
        case 'Incidente':
          p_Datos_Salida_CT.getRange('K' + fila).setValue('I');
          break;
        case 'Acto inseguro':
        case 'Actos Inseguros':
          p_Datos_Salida_CT.getRange('K' + fila).setValue('AI');
          break;
        case 'Incidentes ambientales':
          p_Datos_Salida_CT.getRange('K' + fila).setValue('IA');
          break;
        case 'Acto Inseguro ambientales':
          p_Datos_Salida_CT.getRange('K' + fila).setValue('AIA');
          break;
        case 'Defecto':
          p_Datos_Salida_CT.getRange('K' + fila).setValue('DF');
          break;
        case 'Acto y/o comportamiento':
          p_Datos_Salida_CT.getRange('K' + fila).setValue('AC');
          break;
        case 'Condición de operación':
          p_Datos_Salida_CT.getRange('K' + fila).setValue('CO');
          break;
        default:
          break;
      }
    }
  }

