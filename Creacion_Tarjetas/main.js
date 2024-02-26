//Funcion para Generar Alertas, Menus Personalizados, etc.
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Menú Personalizado')
    .addItem('Duplicar Datos', 'duplicateData')
    .addItem('Limpiar Datos de Entrada', 'confirmClearDataEntry')
    .addItem('Limpiar Datos de Salida', 'confirmClearDataOutput')
    .addItem('Limpiar Todo', 'cleanAll')
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

//Funcion Para Confirmar Limpieza de Datos
function confirmAndCleanData(sheetName, confirmationMessage, lastColumn) {
  const ui = SpreadsheetApp.getUi();
  const respuesta = ui.alert(
    'Confirmación',
    confirmationMessage,
    ui.ButtonSet.YES_NO);

  if (respuesta == ui.Button.YES) {
    const { sheet } = conectionSheets();
    const targetSheet = sheet.getSheetByName(sheetName);
    const lastRow = targetSheet.getLastRow();
    const range = 'A2:' + lastColumn + lastRow;
    targetSheet.getRange(range).clearContent();
  }
}

//Funcion que permite Limpiar los datos del formulario sheets la hoja "CT_Input_Data"
function confirmClearDataEntry() {
  //Se debe especificar hasta el numero de Columna que se desea eliminar (ultimo parametro)
  confirmAndCleanData('CT_Input_Data', '¿Está seguro de que desea limpiar los datos de "Entrada"?\n\nEste proceso limpiará cualquier tipo de dato', 'AK');
}

//Funcion que permite Limpiar los datos del formulario sheets la hoja "CT_Output_Data"
function confirmClearDataOutput() {
  //Se debe especificar hasta el numero de Columna que se desea eliminar (ultimo parametro)
  confirmAndCleanData('CT_Output_Data', '¿Está seguro de que desea limpiar los datos de "Salida"?\n\nEste proceso limpiará cualquier tipo de dato.', 'Q');
}



//Funcion para Limpiar los dos archivos
function cleanAll() {
  const ui = SpreadsheetApp.getUi();
  const respuesta = ui.alert(
    'Confirmación',
    '¿Está seguro de que desea limpiar todas las hojas?\n\nEste proceso limpiará cualquier tipo de dato en las hojas "CT_Input_Data" y "CT_Output_Data".',
    ui.ButtonSet.YES_NO
  );

  if (respuesta == ui.Button.YES) {
    const { sheet } = conectionSheets();
    const inputSheet = sheet.getSheetByName('CT_Input_Data');
    const outputSheet = sheet.getSheetByName('CT_Output_Data');
    const lastRowInput = inputSheet.getLastRow();
    const lastRowOutput = outputSheet.getLastRow();

    // Limpiar hoja de Input
    if (lastRowInput > 1) { // Verificar que haya datos en la hoja de entrada
      inputSheet.getRange('A2:AK' + lastRowInput).clearContent();
    }

    // Limpiar hoja de Output
    if (lastRowOutput > 1) { // Verificar que haya datos en la hoja de salida
      outputSheet.getRange('A2:Q' + lastRowOutput).clearContent();
    }
  }
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

    //Se llama a la funcion assignCodesPlace
    assignCodesPlace(p_CT_Output_Data, row);

    // Se llama a la función assignCodesPG
    assignCodesPG(p_CT_Output_Data, row);

    // Se llama a la función assignCodesRisk
    assignCodesRisk(p_CT_Output_Data, row);
    
    // Se llama a la función assignCodesPriority
    assignCodesPriority(p_CT_Output_Data, row);
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
      return 'Error_Cod_Title';
  }
}

//Funcion que Concatena 2 Columnas para Obtener el titulo completo de la Tarjeta = Columna COD_SHORT_DESC_TITLE
function concatenateColumnsTitle(sheet, row) {
  const columnKValue = sheet.getRange('K' + row).getValue();
  const columnBValue = sheet.getRange('B' + row).getValue();
  const concatenatedValue = columnKValue + ' ' + columnBValue;
  sheet.getRange('L' + row).setValue(concatenatedValue);
}

//Funcion Para Asociar Los lugares a un Codigo cod_place y imprimirlo en la columna especifica
function assignCodesPlace(sheet) {
  const descriptions = [
    ["Logística - Materias Primas", "DR15-ALMG"/*, mas codigos y especificos*/],
    ["Logística -  Almacén General", "DR15-ALMG"],
    ["Manufactura - Molino", "DR15-MOL1"],
    ["Manufactura - Pastificio", "DR15-PAST"],
    ["Manufactura - Empaque", "DR15-EMPA"],
    ["Edificio Información Manufactura", "DR15-PAST-OEPA"],
    ["Ingeniería y Montajes", "DR15-TMTO-INGE"],
    ["Servicios Técnicos", "DR15-TMTO"],
    ["SDM", "DR15-PAST-SDME"],
    ["Empaques especiales (CEMPA)", "DR15-EMPA-EESP"],
    ["Logística CEDI A", "DR15-OPER-CEDI"],
    ["Logística CEDI B", "DR15-PAST-CD_B"],
    ["Calidad Integral", "DR15-PAST-OCAL"],
    ["Laboratorio de Calidad", "DR15-LABS-LCAL"],
    ["Laboratorio I+D", "DR15-LABS-INV"],
    ["Edificio Administrativo", "DR15-EADM"],
    ["Mercadeo", "DR15-EADM-OEAD"],
    ["Exteriores", "DR15-EXTE"],
    ["Plantas de tratamiento de aguas (PTAR - PTAP)", "DR15-PTAR"],
    ["Bodega de excedentes industriales", "DR15-CRES"],
    ["Zona de contratistas", "DR15-ZCNT"],
    ["Portería", "DR15-PORT"],
    ["Casino", "DR15-EADM-CSNO"],
    ["Cuarto de Baterías", "DR15-OPER-CEDI"],
    ["Cuarto Venta de Empleados", "DR15-OPER-CEDI"],
  ];

  const lastRow = sheet.getLastRow();

  for (let row = 2; row <= lastRow; row++) {
    const description = sheet.getRange('F' + row).getValue();
    const codeInfo = descriptions.find(entry => entry[0] == description);
    if (codeInfo) {
      sheet.getRange('M' + row).setValue(codeInfo[1]);
      // if (codeInfo.length > 2) {
      //   sheet.getRange('N' + row).setValue(codeInfo[2]);
      // }
    }
  }
}

//Funcion Para Asociar Los grupos de planeacion a un Codigo "COD_PLANNER_GROUP" y imprimirlo en las columnas especificadas 
function assignCodesPG(sheet, row) {
  const valueG = sheet.getRange('G' + row).getValue();
  let codeN = '';
  let codeO = '';

  switch (valueG) {
    case 'Jefe de ingeniería y montajes':
      codeN = 'M12';
      break;
    case 'Obras civiles':
      codeN = 'M06';
      codeO = 'CONTCVIL';
      break;
    case 'Jefe Aseguramiento de calidad':
      codeN = 'M07';
      codeO = 'ANLICAL';
      break;
    case 'Coordinador de gestión ambiental':
      codeN = 'M07';
      codeO = 'JEFEGAMB';
      break;
    case 'Equipo SST':
      codeN = 'M08';
      codeO = 'COORSST';
      break;
    case 'Jefe servicios administrativos':
      codeN = 'M12';
      break;
    case 'Jefe de empaque':
      codeN = 'M03';
      break;
    case 'Jefe de pastificio':
      codeN = 'M02';
      break;
    case 'Jefe de molino':
      codeN = 'M01';
      break;
    case 'Jefe CEDI':
      codeN = 'M10';
      codeO = 'JEFECEDI';
      break;
    case 'Jefe materias primas':
      codeN = 'M13';
      break;
    case 'Jefe de almacén general':
      codeN = 'M14';
      break;
    case 'Metrología':
      codeN = 'M05';
      break;
    case 'Servicios Industriales':
      codeN = 'M04';
      break;
    case 'Autónomo':
      codeN = 'M09';
      break;
    case 'Técnico eléctrico':
      codeN = 'M12';
      codeO = 'JEFIYM02';
      break;
    case 'Técnico mecánico':
      codeN = 'M12';
      codeO = 'JEFIYM01';
      break;
    case 'Sistema de Dosificación y Mezclas':
      codeN = 'M11';
      break;
    default:
      return 'Error_Cod_PG';
  }

  sheet.getRange('N' + row).setValue(codeN);
  sheet.getRange('O' + row).setValue(codeO);
}

//Funcion Para Asociar tipos de riesgo un Codigo "COD_RISK" y imprimirlo en la columna especificada
function assignCodesRisk(sheet, row) {
  const valueI = sheet.getRange('I' + row).getValue();
  let codeQ = '';

  switch (valueI) {
    case 'Riesgo De Falla':
    case 'Riesgo de falla de equipo':
      codeQ = '0010';
      break;
    case 'Riesgo De Calidad':
    case 'Riesgo de calidad':
      codeQ = '0020';
      break;
    case 'Riesgo A las Personas':
    case 'Riesgo a las personas':
      codeQ = '0030';
      break;
    case 'Riesgo Ambiental':
    case 'Riesgo ambiental':
      codeQ = '0040';
      break;
    case 'Riesgo Inocuidad':
    case 'Riesgo de inocuidad':
      codeQ = '0050';
      break;
    default:
      return 'Error_Cod_Risk';
  }

  sheet.getRange('Q' + row).setValue(codeQ);
}

//Funcion para Sacar codigo o iniciales para columna Cod_Priority 
function assignCodesPriority(sheet, row) {
  const riskType = sheet.getRange('H' + row).getValue();
  let priorityCode = '';

  switch (riskType) {
    case 'A - Alta':
      priorityCode = 'H';
      break;
    case 'B - Media':
      priorityCode = 'L';
      break;
    case 'C - Baja':
      priorityCode = 'N';
      break;
    default:
      return 'Error_Cod_Priority';
      break;
  }

  sheet.getRange('P' + row).setValue(priorityCode);
}
