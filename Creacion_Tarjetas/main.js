//Funcion para Generar Alertas, Menus Personalizados, etc.
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🤖➡ Menú Personalizado')
    .addItem('📄📄- Duplicar Datos', 'duplicateData')
    .addItem('➧📁- Limpiar Datos de Entrada', 'confirmClearDataEntry')
    .addItem('📁➧-Limpiar Datos de Salida', 'confirmClearDataOutput')
    .addItem('📖- Limpiar Todo', 'cleanAll')
    .addItem('💾- Convertir Salida a Excel', 'convertToExcel')
    .addToUi();
}
//Funcion para Conectarse al Sheet
function conectionSheets() {
   //Conectar Sheets a AppScript
  const sheetId = '1IfbxGR6tHOPCHc0r2oVb5R9B598clH6V5Fh5aNiZKqE'; //1pPWE_pS5tNcRabGIIymprfMo2TNWswZKNv5ovOZ95dY Cod Alterno
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
  confirmAndCleanData('CT_Output_Data', '¿Está seguro de que desea limpiar los datos de "Salida"?\n\nEste proceso limpiará cualquier tipo de dato.', 'R');
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
      inputSheet.getRange('A2:AM' + lastRowInput).clearContent();
    }

    // Limpiar hoja de Output
    if (lastRowOutput > 1) { // Verificar que haya datos en la hoja de salida
      outputSheet.getRange('A2:R' + lastRowOutput).clearContent();
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
    const person_name = p_CT_Input_Data.getRange('D' + row).getValue().slice(0, 12);
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
      return 'AI';
    case 'Defecto':
      return 'DC';
    case 'Acto y/o comportamiento':
      return 'AC';
    case 'Condición de operación':
      return 'CO';
    default:
      return 'Error_Cod_Card_Title';
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
    ["Logística - Materias Primas", "DR15-ALMG"/*, mas codigos y especificos en caso de requerir*/],
    ["Logística -  Almacén General", "DR15-ALMG"],
    ["Manufactura - Molino", "DR15-MOL1"],
    ["Manufactura - Pastificio A", "DR15-PAST-FPL1"],
    ["Manufactura - Pastificio B", "DR15-PAST-LINE- B"],
    ["Manufactura - Pastificio C", "DR15-PAST-FPC1"],
    ["Manufactura - Pastificio D", "DR15-EMPA-EFGC"],
    ["Manufactura - Empaque Pasta Larga", "DR15-EMPA-EFPL"],
    ["Manufactura - Empaque Pasta Corta", "DR15-EMPA-EFGC"],
    ["Edificio Información Manufactura", "DR15-PAST-OEPA"],
    ["Ingeniería y Montajes", "DR15-TMTO-INGE"],
    ["Servicios Técnicos", "DR15-TMTO"],
    ["Metrología", "DR15-TMTO"],
    ["SDM", "DR15-PAST-SDME"],
    ["Empaques especiales (CEMPA)", "DR15-EMPA-EESP"],
    ["Logística CEDI A", "DR15-OPER-CEDI"],
    ["Logística CEDI B", "DR15-PAST-CD_B"],
    ["Laboratorio de Calidad", "DR15-LABS-LCAL"],
    ["Laboratorio I+D", "DR15-LABS-INV"],
    ["Edificio Administrativo", "DR15-EADM"],
    ["Exteriores", "DR15-EXTE"],
    ["Plantas de tratamiento de aguas Residuales (PTAR)", "DR15-PTAR"],
    ["Plantas de tratamiento de agua Potable (PTAP)", "DR15-PTAP"],
    ["Bodega de excedentes industriales", "DR15-CRES"],
    ["Zona de contratistas", "DR15-ZCNT"],
    ["Portería", "DR15-PORT"],
    ["Casino", "DR15-EADM-CSNO"],
    ["Cuarto de Baterías", "DR15-OPER-CEDI"],
    ["Cuarto Venta de Empleados", "DR15-OPER-CEDI"],
    ["Taller de Mantenimientos", "DR15-TMTO"],
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
  const valuePlace = sheet.getRange('F' + row).getValue(); //De la hoja de Salida toma la columna F que pertenece al lugar
  const valueResponsible = sheet.getRange('G' + row).getValue();  //
  let codeN = '';
  let codeO = '';

    switch (valuePlace) {
      case 'Logística - Materias Primas':
        codeN = 'M13';
        switch(valueResponsible) {
          case 'Tecnico Eléctrico':
            codeO = 'TECN003';
            break;
          case 'Tecnico Mecánico':
            codeO = 'MECA013';
            break;
          case 'Autónomo':
            codeO = 'JEFE_MP';
            break;
          default:
            break;
        }
      break;

      case 'Logística -  Almacén General':
        codeN = 'M14';
        switch(valueResponsible) {
          case 'Tecnico Eléctrico':
            codeO = 'TECN003';
            break;
          case 'Tecnico Mecánico':
            codeO = 'MECA013';
            break;
          case 'Autónomo':
            codeO = 'JEFEALG';
            break;
          default:
            break;
        }
      break;

      case 'Manufactura - Molino':
        codeN = 'M01';
        switch(valueResponsible) {
          case 'Tecnico Eléctrico':
            codeO = 'TECN003';
            break;
          case 'Tecnico Mecánico':
            codeO = 'TECN001';
            break;
          case 'Autónomo':
            codeO = 'JEFEMOL';
            break;
          default:
            break;
        }
      break;

      case 'Manufactura - Pastificio A':
      case 'Manufactura - Pastificio B':
      case 'Manufactura - Pastificio C':
      case 'Manufactura - Pastificio D':
        codeN = 'M02';
        switch(valueResponsible) {
          case 'Tecnico Eléctrico':
            codeO = 'ELECT008';
            break;
          case 'Tecnico Mecánico':
            codeO = 'TECN008';
            break;
          case 'Autónomo':
            codeO = 'JEFEPAST';
            break;
          default:
            break;
        }
      break;

      case 'Manufactura - Empaque Pasta Larga':
      case 'Manufactura - Empaque Pasta Corta':
        codeN = 'M03';
        switch(valueResponsible) {
          case 'Tecnico Eléctrico':
            codeO = 'ELECT005';
            break;
          case 'Tecnico Mecánico':
            codeO = 'TECN004';
            break;
          case 'Autónomo':
            codeO = 'JEFE_EMP';
            break;
          default:
            break;
        }
      break;

      case 'Edificio Información Manufactura':
        codeN = 'M15';
        switch(valueResponsible) {
          case 'Tecnico Eléctrico':
            codeO = 'TECN003';
            break;
          case 'Tecnico Mecánico':
            codeO = 'MECA013';
            break;
          case 'Autónomo':
            codeO = 'JEFE_EMP';
            break;
          default:
            break;
        }
      break;

      case 'Ingeniería y Montajes':
        codeN = 'M12';
        switch(valueResponsible) {
          case 'Tecnico Eléctrico':
            codeO = 'ELECT004';
            break;
          case 'Tecnico Mecánico':
            codeO = 'MECA013';
            break;
          case 'Autónomo':
            codeO = 'COORING5';
            break;
          default:
            break;
        }
      break;

      case 'Servicios Técnicos':
        codeN = 'M04';
        switch(valueResponsible) {
          case 'Tecnico Eléctrico':
            codeO = 'TECN003';
            break;
          case 'Tecnico Mecánico':
            codeO = 'MECA013';
            break;
          case 'Autónomo':
            codeO = 'JEFIYM03';
            break;
          default:
            break;
        }
      break;

      case 'Metrología':
        codeN = 'M05';
        switch(valueResponsible) {
          case 'Tecnico Eléctrico':
            codeO = 'ELECT009';
            break;
          case 'Tecnico Mecánico':
            codeO = 'MECA013';
            break;
          case 'Autónomo':
            codeO = 'METRO001';
            break;
          default:
            break;
        }
      break;

      case 'Taller de Mantenimientos':
        codeN = 'M16';
        switch(valueResponsible) {
          case 'Tecnico Eléctrico':
            codeO = 'ELECT009';
            break;
          case 'Tecnico Mecánico':
            codeO = 'MECA013';
            break;
          case 'Autónomo':
            codeO = 'COORING5';
            break;
          default:
            break;
        }
      break;

      case 'SDM':
        codeN = 'M11';
        switch(valueResponsible) {
          case 'Tecnico Eléctrico':
            codeO = 'ELECT008';
            break;
          case 'Tecnico Mecánico':
            codeO = 'TECN009';
            break;
          case 'Autónomo':
            codeO = 'JEFIYM03';
            break;
          default:
            break;
        }
      break;

      case 'Empaques especiales (CEMPA)':
        codeN = 'M17';
        switch(valueResponsible) {
          case 'Tecnico Eléctrico':
            codeO = 'ELECT005';
            break;
          case 'Tecnico Mecánico':
            codeO = 'TECN004';
            break;
          case 'Autónomo':
            codeO = 'JEFE_EMP';
            break;
          default:
            break;
        }
      break;

      case 'Logística CEDI A':
        codeN = 'M10';
        switch(valueResponsible) {
          case 'Tecnico Eléctrico':
            codeO = 'TECN003';
            break;
          case 'Tecnico Mecánico':
            codeO = 'MECA013';
            break;
          case 'Autónomo':
            codeO = 'JEFECEDI';
            break;
          default:
            break;
        }
      break;

      case 'Logística CEDI B':
        codeN = 'M10';
        switch(valueResponsible) {
          case 'Tecnico Eléctrico':
            codeO = 'TECN003';
            break;
          case 'Tecnico Mecánico':
            codeO = 'MECA013';
            break;
          case 'Autónomo':
            codeO = 'JEFECEDI';
            break;
          default:
            break;
        }
      break;
      
      case 'Laboratorio de Calidad':
        codeN = 'M07';
        switch(valueResponsible) {
          case 'Tecnico Eléctrico':
            codeO = 'TECN003';
            break;
          case 'Tecnico Mecánico':
            codeO = '';
            break;
          case 'Autónomo':
            codeO = 'JEFECAL';
            break;
          default:
            break;
        }
      break;

      case 'Laboratorio I+D':
        codeN = 'M18';
        switch(valueResponsible) {
          case 'Tecnico Eléctrico':
            codeO = 'TECN003';
            break;
          case 'Tecnico Mecánico':
            codeO = 'MECA013';
            break;
          case 'Autónomo':
            codeO = 'LABI&D';
            break;
          default:
            break;
        }
      break;

      case 'Edificio Administrativo':
        codeN = 'M19';
        switch(valueResponsible) {
          case 'Tecnico Eléctrico':
            codeO = 'TECN003';
            break;
          case 'Tecnico Mecánico':
            codeO = 'MECA013';
            break;
          case 'Autónomo':
            codeO = 'CONTCVIL';
            break;
          default:
            break;
        }
      break;

      case 'Exteriores':
        codeN = 'M20';
        switch(valueResponsible) {
          case 'Tecnico Eléctrico':
            codeO = 'TECN003';
            break;
          case 'Tecnico Mecánico':
            codeO = 'MECA013';
            break;
          case 'Autónomo':
            codeO = 'CONTCVIL';
            break;
          default:
            break;
        }
      break;

      case 'Plantas de tratamiento de aguas Residuales (PTAR)':
        codeN = 'M22';
        switch(valueResponsible) {
          case 'Tecnico Eléctrico':
            codeO = 'TECN003';
            break;
          case 'Tecnico Mecánico':
            codeO = 'MECA013';
            break;
          case 'Autónomo':
            codeO = 'JEFEGAMB';
            break;
          default:
            break;
        }
      break;

      case 'Plantas de tratamiento de agua Potable (PTAP)':
        codeN = 'M21';
        switch(valueResponsible) {
          case 'Tecnico Eléctrico':
            codeO = 'TECN003';
            break;
          case 'Tecnico Mecánico':
            codeO = 'MECA013';
            break;
          case 'Autónomo':
            codeO = 'JEFEGAMB';
            break;
          default:
            break;
        }
      break;

      case 'Bodega de excedentes industriales':
        codeN = 'M23';
        switch(valueResponsible) {
          case 'Tecnico Eléctrico':
            codeO = 'TECN003';
            break;
          case 'Tecnico Mecánico':
            codeO = 'MECA013';
            break;
          case 'Autónomo':
            codeO = 'JEFEGAMB';
            break;
          default:
            break;
        }
      break;

      case 'Zona de contratistas':
        codeN = 'M06';
        switch(valueResponsible) {
          case 'Tecnico Eléctrico':
            codeO = 'TECN003';
            break;
          case 'Tecnico Mecánico':
            codeO = 'MECA013';
            break;
          case 'Autónomo':
            codeO = 'CONTCVIL';
            break;
          default:
            break;
        }
      break;
      case 'Portería':
        codeN = 'M24';
        switch(valueResponsible) {
          case 'Tecnico Eléctrico':
            codeO = 'TECN003';
            break;
          case 'Tecnico Mecánico':
            codeO = 'MECA013';
            break;
          case 'Autónomo':
            codeO = 'CONTCVIL';
            break;
          default:
            break;
        }
      break;

      case 'Casino':
        codeN = 'M25';
        switch(valueResponsible) {
          case 'Tecnico Eléctrico':
            codeO = 'TECN003';
            break;
          case 'Tecnico Mecánico':
            codeO = 'MECA013';
            break;
          case 'Autónomo':
            codeO = 'CONTCVIL';
            break;
          default:
            break;
        }
      break;

      case 'Cuarto de Baterías':
        codeN = 'M26';
        switch(valueResponsible) {
          case 'Tecnico Eléctrico':
            codeO = 'TECN003';
            break;
          case 'Tecnico Mecánico':
            codeO = 'MECA013';
            break;
          case 'Autónomo':
            codeO = 'JEFECEDI';
            break;
          default:
            break;
        }
      break;

      case 'Cuarto Venta de Empleados':
        codeN = 'M27';
        switch(valueResponsible) {
          case 'Tecnico Eléctrico':
            codeO = 'TECN003';
            break;
          case 'Tecnico Mecánico':
            codeO = 'MECA013';
            break;
          case 'Autónomo':
            codeO = 'JEFECEDI';
            break;
          default:
            break;
        }
      break;

      default:
        return 'Error_Cod_PG';
    }
    

    if(valueResponsible === 'Jefe Aseguramiento de Calidad')
      codeO = 'ANLICAL';

    if(valueResponsible === 'Coordinador de Gestión Ambiental')
      codeO = 'JEFEGAMB';

    if(valueResponsible === 'Equipo SST')
      codeO = 'COORSST';

    if(valueResponsible === 'Obras Civiles')
      codeO = 'CONTCVIL';

    if(valueResponsible === 'Reparaciones Metalmecanicas IMB')
      codeO = 'CONTMEC';

    if(valueResponsible === 'Coordinador de Proyectos')
      codeO = 'COORING5';



  sheet.getRange('N' + row).setValue(codeN);  //Cod_Planner_Group
  sheet.getRange('O' + row).setValue(codeO);  //Cod_Planner_Group_Complement
}

//Funcion Para Asociar tipos de riesgo un Codigo "COD_RISK" y imprimirlo en la columna especificada
function assignCodesRisk(sheet, row) {
  const valueI = sheet.getRange('I' + row).getValue();
  let codeQ = '';

  switch (valueI) {
    case 'Riesgo De Falla':
    case 'Riesgo de falla de equipo':
      codeQ = '*0010';
      break;
    case 'Riesgo De Calidad':
    case 'Riesgo de calidad':
      codeQ = '*0020';
      break;
    case 'Riesgo A las Personas':
    case 'Riesgo a las personas':
      codeQ = '*0030';
      break;
    case 'Riesgo Ambiental':
    case 'Riesgo ambiental':
      codeQ = '*0040';
      break;
    case 'Riesgo Inocuidad':
    case 'Riesgo de inocuidad':
      codeQ = '*0050';
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
    case 'A - Alta (3 Días)':
      priorityCode = 'H';
      break;
    case 'B - Media (15 Días)':
      priorityCode = 'L';
      break;
    case 'C - Baja (30 Días)':
      priorityCode = 'N';
      break;
    default:
      return 'Error_Cod_Priority';
  }

  sheet.getRange('P' + row).setValue(priorityCode);
}

function convertToExcel() {
  var hoja = SpreadsheetApp.getActiveSpreadsheet();
  var hojaSeleccionada = hoja.getSheetByName('CT_Output_Data');

  // Verificar si la hoja 'CT_Output_Data' existe
  if (!hojaSeleccionada) {
    SpreadsheetApp.getUi().alert("La hoja 'CT_Output_Data' no existe.");
    return;
  }

  // Obtener los datos de la hoja seleccionada
  var data = hojaSeleccionada.getDataRange().getValues();

  // Crear un nuevo archivo de Excel en Google Drive
  var newSpreadsheet = SpreadsheetApp.create('CT_Output_Data_Excel');
  var newSheet = newSpreadsheet.getActiveSheet();
  newSheet.getRange(1, 1, data.length, data[0].length).setValues(data);

  // Obtener el ID del archivo de Excel recién creado
  var fileId = newSpreadsheet.getId();

  // Obtener la URL de descarga del archivo de Excel
  var url = "https://docs.google.com/spreadsheets/d/" + fileId + "/export?format=xlsx";

  // Abrir la URL en una nueva ventana o pestaña
  var html = "<script>window.open('" + url + "');</script>";
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html), "Descargar archivo");
}