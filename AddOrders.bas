Attribute VB_Name = "AddOrders"
Sub AddOrdersSheet()
' This macro will loop through the procurement data and copy rows with an order date, then paste it to the Order sheet.
' Created 8/30/2024.

    ' Variable declarations.
    Dim MyRange As Range
    Dim Stock As Range
    Dim RowCount As Long
    
    Set MyRange = SheetProc.Range("F6", Range("f6").End(xlDown))
    RowCount = 6
    
    ' User selection, if answer is no the sub will end.
    If Range("B" & RowCount) = "" Then
        MsgBox "Please verify data is imported first and try again!", vbCritical, "Filter Failed"
        Exit Sub
    End If
       
    ' Enable error trapping.
    On Error GoTo ErrHandler:
    
    ' Turn off screen updating
    Call TurnOffApps
   
    ' Start of loop to filter range and hide rows.
    For Each Stock In MyRange
        If Stock.Value <> "" And Range("H" & RowCount) <> "" Then
            Range("B" & RowCount, "J" & RowCount).Copy
            SheetOrders.Cells(Rows.Count, "B").End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
        End If
            RowCount = RowCount + 1
        If Range("B" & RowCount) = "" Then Exit For
    Next

    ' Turn on screen updating
   Call TurnOnApps

    ' Clean Exit
    Exit Sub

ErrHandler:
    MsgBox "Add To Orders Failed!, Please try again.", vbCritical
    Call TurnOnApps
End Sub





