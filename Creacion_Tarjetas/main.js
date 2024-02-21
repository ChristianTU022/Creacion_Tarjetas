
//Funcion para Generar Alertas, Menus Personalizados, etc.
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Menú Personalizado')
    .addItem('Duplicar Datos', 'duplicateData')
    .addItem('Limpiar Datos de Entrada', 'confirmClearDataEntry')
    .addItem('Limpiar Datos de Salida', 'confirmClearDataOutput')
    .addToUi();
}

function confirmClearDataEntry() {
  const ui = SpreadsheetApp.getUi();
  const respuesta = ui.alert(
    'Confirmación',
    '¿Está seguro de que desea limpiar los datos de entrada? Este proceso limpiará cualquier tipo de dato.',
    ui.ButtonSet.YES_NO);

  if (respuesta == ui.Button.YES) {
    cleanDataInput();
  }
}

function confirmClearDataOutput() {
  const ui = SpreadsheetApp.getUi();
  const respuesta = ui.alert(
    'Confirmación',
    '¿Está seguro de que desea limpiar los datos de Salida? Este proceso limpiará cualquier tipo de dato.',
    ui.ButtonSet.YES_NO);

  if (respuesta == ui.Button.YES) {
    cleanDataOutput();
  }
}


//Funcion que permite Limpiar los datos del formulario sheets la hoja "CT_Input_Data"
function cleanDataInput() {
  const sheet = SpreadsheetApp.openById('1pPWE_pS5tNcRabGIIymprfMo2TNWswZKNv5ovOZ95dY');
  const p_CT_Output_Data = sheet.getSheetByName('CT_Input_Data');
  const lastRow = p_CT_Output_Data.getLastRow();
  p_CT_Output_Data.getRange('A2:AK' + lastRow).clearContent();
}
//Funcion que permite Limpiar los datos del formulario sheets la hoja "CT_Output_Data"
function cleanDataOutput() {
  const sheet = SpreadsheetApp.openById('1pPWE_pS5tNcRabGIIymprfMo2TNWswZKNv5ovOZ95dY');
  const p_CT_Output_Data = sheet.getSheetByName('CT_Output_Data');

  const lastRow = p_CT_Output_Data.getLastRow();
  p_CT_Output_Data.getRange('A2:K' + lastRow).clearContent();
}


function duplicateData() {
    //Conectar Sheets a AppScript
    const sheet = SpreadsheetApp.openById('1pPWE_pS5tNcRabGIIymprfMo2TNWswZKNv5ovOZ95dY');
    //Conectar Hojas especificas
    const p_CT_Input_Data = sheet.getSheetByName('CT_Input_Data');
    const p_CT_Output_Data = sheet.getSheetByName('CT_Output_Data');
    
    const lastRow = p_CT_Input_Data.getLastRow();
    
    //Saca Los datos especificos del Archivo/hoja de Entrada al archivo/hoja de salida
    for (let row = 2; row <= lastRow; row++) 
    {
      const date = p_CT_Input_Data.getRange('A' + row).getValue();
      const short_description = p_CT_Input_Data.getRange('F' + row).getValue();
      const long_description = p_CT_Input_Data.getRange('G' + row).getValue();
      const card_title = p_CT_Input_Data.getRange('I' + row).getValue() + '' +
                            p_CT_Input_Data.getRange('J' + row).getValue() + '' +
                            p_CT_Input_Data.getRange('K' + row).getValue() + '' +
                            p_CT_Input_Data.getRange('L' + row).getValue() + '' +
                            p_CT_Input_Data.getRange('N' + row).getValue();
      const person_name = p_CT_Input_Data.getRange('D' + row).getValue();
      const place = p_CT_Input_Data.getRange('C' + row).getValue();
      const plannerGroup = p_CT_Input_Data.getRange('O' + row).getValue();
      const priority = p_CT_Input_Data.getRange('T' + row).getValue();
      const riskType = p_CT_Input_Data.getRange('H' + row).getValue();
       

      if (date !== '') 
      {
        p_CT_Output_Data.appendRow([date, short_description, long_description, card_title, person_name, place, plannerGroup, priority, riskType]);
      }

      //Sacar Codigo o iniciales del Titulo de la Tarjeta
      let valor_Cod_Titulo = p_CT_Output_Data.getRange('D' + row).getValue();
      Logger.log(valor_Cod_Titulo + " Valor Tomado");

      switch (valor_Cod_Titulo) {
        case 'Condición básica':
        case 'Condiciones básicas':
          p_CT_Output_Data.getRange('K' + row).setValue('CB');
          break;
        case 'Condición insegura':
          p_CT_Output_Data.getRange('K' + row).setValue('CI');
          break;
        case 'Incidente':
          p_CT_Output_Data.getRange('K' + row).setValue('I');
          break;
        case 'Acto inseguro':
        case 'Actos Inseguros':
          p_CT_Output_Data.getRange('K' + row).setValue('AI');
          break;
        case 'Incidentes ambientales':
          p_CT_Output_Data.getRange('K' + row).setValue('IA');
          break;
        case 'Acto Inseguro ambientales':
          p_CT_Output_Data.getRange('K' + row).setValue('AIA');
          break;
        case 'Defecto':
          p_CT_Output_Data.getRange('K' + row).setValue('DF');
          break;
        case 'Acto y/o comportamiento':
          p_CT_Output_Data.getRange('K' + row).setValue('AC');
          break;
        case 'Condición de operación':
          p_CT_Output_Data.getRange('K' + row).setValue('CO');
          break;
        default:
          break;
      }
    }
  }

