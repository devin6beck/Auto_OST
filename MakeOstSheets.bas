Sub MakeOSTSheets()
    Dim ws As Worksheet
    Dim newWsName As String
    Dim newWs As Worksheet
    Dim templateWs As Worksheet
    
    Application.ScreenUpdating = False ' Optimize performance
    
    ' Set the template worksheet
    On Error Resume Next
    Set templateWs = ThisWorkbook.Sheets("Template")
    On Error GoTo 0
    
    If templateWs Is Nothing Then
        MsgBox "Template sheet 'Template' not found!", vbExclamation
        Exit Sub
    End If
    
    For Each ws In ThisWorkbook.Sheets
        If Right(ws.Name, 5) = " Data" Then
            ' Create the new worksheet name by replacing " Data" with " OST"
            newWsName = Left(ws.Name, Len(ws.Name) - 5) & " OST"
            
            ' Check if the sheet already exists to avoid duplication or error
            On Error Resume Next ' Ignore error if sheet doesn't exist
            Set newWs = Nothing
            Set newWs = ThisWorkbook.Sheets(newWsName)
            On Error GoTo 0 ' Resume normal error handling
            
            If newWs Is Nothing Then
                ' If it doesn't exist, copy the template sheet
                templateWs.Copy after:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
                Set newWs = ActiveSheet
                newWs.Name = newWsName
            End If
        End If
    Next ws
    
    Application.ScreenUpdating = True ' Turn back on screen updating
End Sub


