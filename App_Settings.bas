Attribute VB_Name = "App_Settings"
Sub TurnOffApps()
' This sub will turn off excel settings to speed up macro.
    With ThisWorkbook.Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
        .CutCopyMode = False
    End With
    ' Stops excel print preview from calculating. When inserting/hiding rows this will slow down macro if not turned off due to constant re calc.
        ActiveSheet.DisplayPageBreaks = False
End Sub

Sub TurnOnApps()
' This sub will turn on excel settings after macro has been completed.
    With ThisWorkbook.Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
        .CutCopyMode = False
    End With
End Sub
