Attribute VB_Name = "FilterListData"
Sub FilterData()
' This sub will filter the data column "Plant Stock", any value with 0 will hide the row, then unhide the row based on user action.


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
        
    On Error GoTo ErrHandler:
        
    ' Turn off screen updating
       Call TurnOffApps
       
    ' Start of loop to filter range and hide rows.
        For Each Stock In MyRange
            If Stock.Value = 0 And Range("H" & RowCount) = "" Then Rows(RowCount).Hidden = True
                RowCount = RowCount + 1
            If Range("B" & RowCount) = "" Then Exit For
        Next
    
    ' Turn on screen updating
       Call TurnOnApps
    
    ' Clean Exit
        Exit Sub
    
ErrHandler:
        MsgBox "Filter Failed!, Please try again.", vbCritical
        Call TurnOnApps
    End Sub
    
Sub UnFilterData()
' This sub will unhide all the filtered data rows on the Bakery Proc Sheet.
' Created by Antonio Lassalle on 08/18/2024.
    
    On Error GoTo ErrHandler:
    
    ' turn of screen updating.
        Call TurnOffApps
    
    ' Unhides all rows hidden by filter sub or user.
        With SheetProc
            .Range("A6:A800").Rows.Hidden = False
            .Range("B6").Select
        End With
    ' turn on screen updating.
        Call TurnOnApps
    
    ' Clean Exit
        Exit Sub
    
ErrHandler:
        MsgBox "Un-Filter Failed!, Please try again.", vbCritical
        Call TurnOnApps
End Sub

