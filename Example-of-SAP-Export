Sub DatenausSAP()

Dim dateiname As String
Dim int_Jahr As Integer
Dim int_Monat As Integer
Dim ws As Workbook
Dim wsSheet As Worksheet
Dim gesellschaft As Worksheet

Dim export As Workbook
Dim exportSheet As Worksheet

Set ws = ThisWorkbook
Set wsSheet = ws.Sheets("Datengrundlage")


'SAP öffnen
Shell "C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe", 4
Set WshShell = CreateObject("WScript.Shell")

Set SapGui = GetObject("SAPGUI")

Set Appl = SapGui.GetScriptingEngine

'Set Connection = Appl.Openconnection("2. C20 - QA/Test", True)
Set Connection = Appl.Openconnection("1. N20 - Production", True)

'SAP Anmeldedaten eingeben
login = MsgBox("Please log in to SAP, then confirm with 'OK'.", vbOKCancel, "SAP")

Set session = Connection.Children(0)

'Fehler abfangen
temp = 0
While (temp = 0)

    int_Jahr = InputBox("Which year do you want to export?", "Reported Year")
    int_Monat = InputBox("Until which month do you want to report? Please insert the number of the month. E.g. '9' for 'september'.", "Reported Month")
   
    temp = 1
    
    If (int_Jahr = Null Or int_Monat = Null) Then
        fehler = MsgBox("Please insert Data.", vbOKOnly, "Error")
        temp = 0
    End If
    
    If (int_Monat < 1 Or int_Monat > 12) Then
        fehler = MsgBox("Please insert correct moth.", vbOKOnly, "Error")
        temp = 0
    End If
    
    If (int_Jahr < 2000 Or int_Jahr > Year(Date)) Then
        fehler = MsgBox("Please insert correct year.", vbOKOnly, "Error")
        temp = 0
    End If
    
    If Not (IsNumeric(int_Monat) = True) Then
      fehler = MsgBox("The inserted month is not a number. Please try again.", vbOKOnly, "Error")
      temp = 0
    End If
    
    If Not (IsNumeric(int_Jahr) = True) Then
      fehler = MsgBox("The inserted year is not a number. Please try again.", vbOKOnly, "Error")
      temp = 0
    End If
    
Wend

'Daten eingeben
'Code beruht auf SAP-Makro
session.findById("wnd[0]").maximize
session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").doubleClickNode "F00002"
session.findById("wnd[0]/usr/ctxtSD_KTOPL-LOW").Text = "ACOA"
session.findById("wnd[0]/usr/ctxtSD_KTOPL-LOW").caretPosition = 4
session.findById("wnd[0]/usr/btn%_SD_BUKRS_%_APP_%-VALU_PUSH").press

'Buchungskreise auswählen
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "0002"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,7]").SetFocus
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,7]").caretPosition = 4
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/radPRSSALL").Select

'Berichtsjahr und Monate eingeben
session.findById("wnd[0]/usr/ctxtB0JAHR").Text = int_Jahr
session.findById("wnd[0]/usr/txtB0MONA-LOW").Text = "1"
session.findById("wnd[0]/usr/txtB0MONA-HIGH").Text = int_Monat

'Abruf durchführen
session.findById("wnd[0]/usr/txtB0MONA-HIGH").SetFocus
session.findById("wnd[0]/usr/txtB0MONA-HIGH").caretPosition = 2
'MsgBox "SAP is working. This could last a few minutes. Please wait."

session.findById("wnd[0]/tbar[1]/btn[8]").press

'Als Excel XXL exportieren (ohne Zwischensummen)
session.findById("wnd[0]").maximize
session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").Select
session.findById("wnd[1]/usr/cmbG_LISTBOX").SetFocus
session.findById("wnd[1]/usr/cmbG_LISTBOX").Key = "08"
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]").SetFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
