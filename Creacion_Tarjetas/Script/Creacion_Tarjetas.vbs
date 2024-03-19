
If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If

Main
Sub Main()
Set objExcel = CreateObject("Excel.Application")
   objExcel.Visible = True ' Haz que Excel sea visible para verificar

' Abrir el archivo Excel
   Set objWorkbook = objExcel.Workbooks.Open("C:\Users\NCDRPRACPROD\Downloads\CT_Output_Data_Excel.xlsx")
   Set objSheet = objWorkbook.Sheets(1) ' Hoja donde están los datos

Dim lastRow, i
lastRow = objSheet.Cells(objSheet.Rows.Count, 1).End(-4162).Row ' Última fila con datos
For i = 2 To lastRow
Dim colShortDescription, colLongDescription, colPersonName
Dim colCodShortDescTitle, colCodPlace, colCodPlannerGroup, colCodPlannerGroupComplement
Dim colCodPriority, colCodRisk, colCodSpecificPlace
   colShortDescription = objSheet.Cells(i, 2).Value ' Columna B
   colLongDescription = objSheet.Cells(i, 3).Value ' Columna C
   colCodShortDescTitle = objSheet.Cells(i, 12).Value ' Columna L
   colCodPlace = objSheet.Cells(i, 13).Value ' Columna M
   colCodPlannerGroup = objSheet.Cells(i, 14).Value ' Columna N
   colCodPlannerGroupComplement = objSheet.Cells(i, 15).Value ' Columna O
   colPersonName = objSheet.Cells(i, 5).Value ' Columna E
   colCodPriority = objSheet.Cells(i, 16).Value ' Columna P
   colCodRisk = objSheet.Cells(i, 17).Value ' Columna Q

If Left(colCodRisk, 1) = "*" Then
        ' Si es un asterisco, eliminar el primer carácter y convertir a texto
        colCodRisk = Right(colCodRisk, Len(colCodRisk) - 1)
End If

session.findById("wnd[0]/tbar[0]/okcd").text = "/NIW21"
session.findById("wnd[0]").sendVKey 0

session.findById("wnd[0]/usr/ctxtRIWO00-QMART").text = "ZM"
session.findById("wnd[0]/usr/ctxtRIWO00-QMART").caretPosition = 2
session.findById("wnd[0]").sendVKey 0
Wscript.Sleep 2000
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_2:SAPLIQS0:7715/cntlTEXT/shellcont/shell").text = colShortDescription + vbCr + "" + vbCr + colLongDescription + vbCr + ""
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_2:SAPLIQS0:7715/cntlTEXT/shellcont/shell").setSelectionIndexes 61,61
session.findById("wnd[0]/usr/subSCREEN_1:SAPLIQS0:1050/txtVIQMEL-QMTXT").text = colCodShortDescTitle
Wscript.Sleep 1000
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7322/subOBJEKT:SAPLIWO1:0100/ctxtRIWO1-TPLNR").text = colCodPlace
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_3:SAPLIQS0:7326/ctxtVIQMEL-INGRP").text = colCodPlannerGroup
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_3:SAPLIQS0:7326/ctxtRIWO00-GEWRK").text = colCodPlannerGroupComplement
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_3:SAPLIQS0:7326/ctxtVIQMEL-QMNAM").text = colPersonName
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_3:SAPLIQS0:7326/ctxtRIWO00-GEWRK").setFocus
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_3:SAPLIQS0:7326/ctxtVIQMEL-QMNAM").caretPosition = 13
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]").sendVKey 0
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_2:SAPLIQS0:7715/cntlTEXT/shellcont/shell").setUnprotectedTextPart 0, colShortDescription + vbCr + "" + vbCr + colLongDescription + vbCr + ""
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_2:SAPLIQS0:7715/cntlTEXT/shellcont/shell").setSelectionIndexes 97,97
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7330/cmbVIQMEL-PRIOK").key = colCodPriority
Wscript.Sleep 2000
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7330/cmbVIQMEL-PRIOK").setFocus
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12").select
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssubSUB_GROUP_10:SAPLIQS0:7130/tblSAPLIQS0AKTIONEN_VIEWER/ctxtVIQMMA-MNGRP[1,0]").text = "DR15TITA"
Wscript.Sleep 2000
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssubSUB_GROUP_10:SAPLIQS0:7130/tblSAPLIQS0AKTIONEN_VIEWER/ctxtVIQMMA-MNCOD[2,0]").text = colCodRisk
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssubSUB_GROUP_10:SAPLIQS0:7130/tblSAPLIQS0AKTIONEN_VIEWER/ctxtVIQMMA-MNCOD[2,0]").caretPosition = 4
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/btn[11]").press

Wscript.Sleep 2000
session.findById("wnd[0]/tbar[0]/okcd").text = "/NIW22"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtRIWO00-QMNUM").SetFocus
objSheet.Cells(i, 10) = session.findById("wnd[0]/usr/ctxtRIWO00-QMNUM").Text
Wscript.Sleep 2000

Next
objWorkbook.Close True
objExcel.Quit
End Sub

Sub ExecuteAndWaitForSAP()
    ' Run saplogon bin
    WScript.CreateObject("WScript.Shell").Run Chr(34) & SAP_GUI_PATH & Chr(34), 2

    ' Wait to be initialized
    isSapInitialized = False
    Do While Not isSapInitialized
        isSapInitialized = IsProcessRunning(SAP_BIN)
    Loop
    
    WScript.Sleep 3000
End Sub

Function IsProcessRunning(targetProcess)
    Set WMIService = GetObject("winmgmts:\\.\root\cimv2")
    query = "SELECT * FROM Win32_Process"
    Set items = WMIService.ExecQuery(query)

    For Each item In items
        If item.Name = targetProcess Then
            IsProcessRunning = True
            Exit Function
        End If
    Next

    IsProcessRunning = False
End Function

Function FileExists(filePath)
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(filePath) Then 
        FileExists = True 
    Else
        FileExists = False
    End If
End Function

