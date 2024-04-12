
Sub FindAndWriteToOstSheet()
    Dim logSheet As Worksheet, infoSheet As Worksheet
    Dim ws As Worksheet
    Dim lastLogRow As Long, foundOST As Boolean
    Dim dataSheetName As String, ostSheetName As String
    Dim currentDate As String
    Set logSheet = ThisWorkbook.Sheets("Log")
    Set infoSheet = ThisWorkbook.Sheets("Info")
    currentDate = Format(Now(), "m/dd/yyyy")
    
    For Each ws In ThisWorkbook.Sheets
        If Not ws.Name Like "*_Data" And ws.Name Like "* Data" Then
            dataSheetName = ws.Name
            ostSheetName = Replace(dataSheetName, " Data", " OST")
            foundOST = SheetExists(ostSheetName)
            
            If foundOST Then
                ThisWorkbook.Sheets(ostSheetName).Range("E14").value = "found"
            Else
                LogMessage "Missing OST sheet for: " & dataSheetName, lastLogRow, logSheet
            End If
        End If
    Next ws
End Sub
Function SheetExists(sheetName As String) As Boolean
    Dim sht As Worksheet
    On Error Resume Next
    Set sht = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    SheetExists = Not sht Is Nothing
End Function
Sub LogMessage(msg As String, ByRef lastRow As Long, ByRef logSheet As Worksheet)
    lastRow = logSheet.Cells(logSheet.Rows.Count, "A").End(xlUp).Row + 1
    With logSheet
        .Cells(lastRow, 1).value = Format(Now, "m/dd/yyyy hh:mm:ss AM/PM") & " - " & msg
    End With
End Sub

