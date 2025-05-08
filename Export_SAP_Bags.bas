Attribute VB_Name = "Export_SAP_Bags"
Sub BagProc()
' This sub will start MD07 in SAP and export the bag information via clip board to the designated workbook.
' Created by Antonio Lassalle - 8/17/2024.

    ' Update Status range for user.
        If Range("Status") = "" Then Range("Status") = "Importing Data.....Please Wait..."
    
    ' Conditional statement to run sub, if canceled exits sub.
        If MsgBox("This macro will import bag inventory from SAP. Would you like to continue?", vbYesNo + vbExclamation, "SAP Bag Procurement") = vbNo Then
            Range("Status") = Null
            Exit Sub
        End If
            
    ' Turn off apps.
        Call TurnOffApps
    
    On Error GoTo ErrHandler:
    
    ' Establish SAP connection.
        Set SapGuiAuto = GetObject("SAPGUI")
        Set SAPApp = SapGuiAuto.GetScriptingEngine
        Set Connection = SAPApp.Children(0)
        Set session = Connection.Children(0)
            
        If IsObject(WScript) Then
            WScript.ConnectObject session, "on"
            WScript.ConnectObject Application, "on"
        End If
    ' Start of SAP transaction.
        session.findById("wnd[0]").resizeWorkingPane 94, 28, False
        session.StartTransaction "MD07"
       
    ' Sap parameter selection for bag data.
        session.findById("wnd[0]/usr/tabsTAB210/tabpF02/ssubINCLUDE210:SAPMM61R:0212/ctxtRM61R-WERKS2").Text = "4014"
        session.findById("wnd[0]/usr/tabsTAB210/tabpF02/ssubINCLUDE210:SAPMM61R:0212/ctxtRM61R-DISPO").Text = "132"
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
        session.findById("wnd[0]/mbar/menu[0]/menu[1]").Select
    ' exports SAP data to the clipboard
        session.findById("wnd[0]").sendVKey 45
        session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").Select
        session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").SetFocus
        session.findById("wnd[1]/tbar[0]/btn[0]").press
        session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
        session.findById("wnd[0]").sendVKey 0
        
    ' Clears old data from columns
        With SheetProc
            .Activate
            .Range("B6:J1000").ClearContents
        End With
        
    ' Clears old data from columns and pastes SAP Data
        With ThisWorkbook.Sheets("SAP Data")
            .Visible = True
            .Activate
            .Cells.ClearContents
            .Range("A1").Select
            .Paste
        End With
        
    ' Format Sap Data with "|" delimiter. This will convert text to columns.
        Columns("A").Select
        Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
            Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
            :="|", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, _
            1), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1), Array(11, 1), Array(12 _
            , 1), Array(13, 1), Array(14, 1), Array(15, 1)), TrailingMinusNumbers:=True
            
    ' Copies data from SAP and installs in desiginated range for Proc Sheet.
        With ThisWorkbook.Sheets("SAP Data")
            .Range("C11", Range("C11").End(xlDown)).Copy SheetProc.Range("B6")
            .Range("D11", Range("D11").End(xlDown)).Copy SheetProc.Range("D6")
            .Range("R11", Range("R11").End(xlDown)).Copy SheetProc.Range("F6")
            .Visible = False
        End With
        
    ' Activate main Worksheet
       With SheetProc
            .Activate
            .Application.Goto Range("A1"), True
       End With
    
    ' Turn on screen updating.
        Call TurnOnApps
    
    ' User notification of success
        MsgBox "Data imported Successfully!", vbInformation, "Success"
        
    ' Update Status range for user.
        Range("Status") = Null
           
    ' Clean exit
        Exit Sub
        
ErrHandler:
        If Err = 614 Then
                MsgBox "Please open a session of SAP and retry angain.", vbCritical
            ElseIf Err < 1 Then MsgBox "Please open a session of SAP and retry angain.", vbCritical
            Else
                MsgBox "Import of Data Failed! Please try again.", vbCritical, "Failed Sub"
        End If
        Range("Status") = Null
        TurnOnApps
End Sub


