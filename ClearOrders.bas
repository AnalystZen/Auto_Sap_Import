Attribute VB_Name = "ClearOrders"
Sub ClearOrderSheet()
' This macro will clear all orders from the order page.
' Created by Antonio Lassalle on 8/30/2024.

    ' Variable declaration.
    Dim MessageResult As VbMsgBoxResult
    
    ' User confirmation.
    MessageResult = MsgBox(Prompt:="Would you like to clear the order data?", Buttons:=vbYesNo + vbExclamation, Title:="Clear Data")
    
    ' User case bases off message selection.
    Select Case MessageResult
    
    Case vbNo
     Exit Sub
    
        Case vbYes
    
    End Select
    
    ' Enable error trapping.
    On Error GoTo ErrHandler:
    
    ' Turn off screen update.
    Call TurnOffApps
    
    ' Clear order data.
    With SheetOrders
        .Activate
        .Range("B6:J1000").ClearContents
        .Application.Goto Range("A1"), True
    End With
    
    ' Turn on screen update.
    Call TurnOnApps
    
    ' User update.
    MsgBox Prompt:="Order Data Cleared!", Buttons:=vbInformation, Title:="Order Data"
    
    ' Clean exit
        Exit Sub
    
ErrHandler:
    ' User update.
        MsgBox Prompt:="Clearing data failed! Please try again.", Buttons:=vbCritical, Title:="Failed"
    ' Turn on screen update.
        Call TurnOnApps
End Sub
