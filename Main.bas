Sub aMain()
    On Error GoTo ErrorHandler
    ' Turn off screen updating and set calculation to manual
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    WriteOstDataMonthNameToTemplate.WriteOstDataMonthNameToTemplate
    
    MakeDataSheets.MakeDataSheets
    
    MakeOstSheet.MakeOstSheet
    
    Placeholder.placeHolderSubName
    
    ' Turn on screen updating and calculation settings
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    MsgBox "Completed successfully :)", vbInformation
    Exit Sub

ErrorHandler:
    ' Ensure that settings are restored even if there is an error
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
End Sub
