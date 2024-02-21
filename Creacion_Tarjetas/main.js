//Funcion para Generar Alertas, Menus Personalizados, etc.
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Menú Personalizado')
    .addItem('Duplicar Datos', 'duplicateData')
    .addItem('Limpiar Datos de Entrada', 'confirmClearDataEntry')
    .addItem('Limpiar Datos de Salida', 'confirmClearDataOutput')
    .addToUi();
}

//Funcion para Conectarse al Sheet
function conectionSheets() {
   //Conectar Sheets a AppScript
  const sheetId = '1pPWE_pS5tNcRabGIIymprfMo2TNWswZKNv5ovOZ95dY';
  const sheet = SpreadsheetApp.openById(sheetId);
   //Conectar Hojas especificas
  const p_CT_Input_Data = sheet.getSheetByName('CT_Input_Data');
  const p_CT_Output_Data = sheet.getSheetByName('CT_Output_Data');
  
  return { sheet, p_CT_Input_Data, p_CT_Output_Data };
}

//Funcion para Generar Alertas, Menus Personalizados, etc.
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Menú Personalizado')
    .addItem('Duplicar Datos', 'duplicateData')
    .addItem('Limpiar Datos de Entrada', 'confirmClearDataEntry')
    .addItem('Limpiar Datos de Salida', 'confirmClearDataOutput')
    .addToUi();
}
//Funcion Para Confirmar Limpieza de Datos de Entrada
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

//Funcion Para Confirmar Limpieza de Datos de Salida
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
  const {p_CT_Input_Data} = conectionSheets();
  const lastRow = p_CT_Input_Data.getLastRow();
  p_CT_Input_Data.getRange('A2:AK' + lastRow).clearContent();
}

//Funcion que permite Limpiar los datos del formulario sheets la hoja "CT_Output_Data"
function cleanDataOutput() {
  const {p_CT_Output_Data} = conectionSheets();
  const lastRow = p_CT_Output_Data.getLastRow();
  p_CT_Output_Data.getRange('A2:Q' + lastRow).clearContent();
}

//Funcion Para Duplicar los Datos
function duplicateData() {
  const { p_CT_Input_Data, p_CT_Output_Data } = conectionSheets();

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

    //Se llama a La funcion get_Cod_Tittle
    const cod_Title_Value = get_Cod_Title(p_CT_Output_Data, row);
    p_CT_Output_Data.getRange('K' + row).setValue(cod_Title_Value);

    //Se llama a la funcion concatenateColumnsTitle
    concatenateColumnsTitle(p_CT_Output_Data, row);
  }
}

//Funcion que Saca Codigo o Iniciales para Columna COD_CARD_TITLE
function get_Cod_Title(sheet, row) {
  let cod_Title_Value = sheet.getRange('D' + row).getValue();
  Logger.log(cod_Title_Value + " Valor Tomado");
  switch (cod_Title_Value) 
  {
    case 'Condición básica':
    case 'Condiciones básicas':
      return 'CB';
    case 'Condición insegura':
      return 'CI';
    case 'Incidente':
      return 'I';
    case 'Acto inseguro':
    case 'Actos Inseguros':
      return 'AI';
    case 'Incidentes ambientales':
      return 'IA';
    case 'Acto Inseguro ambientales':
      return 'AIA';
    case 'Defecto':
      return 'DF';
    case 'Acto y/o comportamiento':
      return 'AC';
    case 'Condición de operación':
      return 'CO';
    default:
      return '';
  }
}

//Funcion que Concatena 2 Columnas para Obtener el titulo completo de la Tarjeta = Columna COD_SHORT_DESC_TITLE
function concatenateColumnsTitle(sheet, row) {
  const columnKValue = sheet.getRange('K' + row).getValue();
  const columnBValue = sheet.getRange('B' + row).getValue();
  const concatenatedValue = columnKValue + ' ' + columnBValue;
  sheet.getRange('L' + row).setValue(concatenatedValue);
}