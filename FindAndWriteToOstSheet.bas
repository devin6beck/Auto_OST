
Sub FindAndWriteToOstSheet()
    On Error GoTo ErrorHandler
    Dim infoSheet As Worksheet, ostWs As Worksheet, codesSheet As Worksheet
    Dim ws As Worksheet
    Dim foundOST As Boolean
    Dim dataSheetName As String, ostSheetName As String, unitNumber As String
    
    Set infoSheet = ThisWorkbook.Sheets("Info")
    Set codesSheet = ThisWorkbook.Sheets("Codes")
    
    For Each ws In ThisWorkbook.Sheets
        If ws.Name Like "* Data" And Not ws.Name Like "*_Data" Then
            dataSheetName = ws.Name
            unitNumber = Left(dataSheetName, InStr(dataSheetName, " Data") - 1)
            ostSheetName = Replace(dataSheetName, " Data", " OST")
            foundOST = SheetExists(ostSheetName)
            MsgBox "unitNumber is: " & unitNumber
            If foundOST Then
                Set ostWs = ThisWorkbook.Sheets(ostSheetName)
                If Not ostWs Is Nothing Then
                    ostWs.Range("A1").value = "Found " & ostSheetName
                    InfoValuesToOstSheet ostWS, unitNumber
                End If
            Else
                MsgBox "Sheet not found: " & ostSheetName
            End If
        End If
    Next ws
    Exit Sub
ErrorHandler:
    MsgBox "An error occurred: " & Err.Description
End Sub

Function SheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function

Sub InfoValuesToOstSheet(ostWS As Worksheet, unitNum As String)
    Dim infoWS As Worksheet
    Set infoWS = ThisWorkbook.Sheets("Info")
    
    Dim lastRow As Long
    lastRow = infoWS.Cells(infoWS.Rows.Count, 1).End(xlUp).Row
    
    Dim matchRow As Long
    matchRow = 0 ' Initialize with 0 to denote no match found initially
    
    ' Loop through column A to find the matching unitNum
    Dim i As Long
    For i = 1 To lastRow
        If infoWS.Cells(i, 1).Value = unitNum Then
            matchRow = i
            Exit For
        End If
    Next i
    
    ' Check if a matching row was found
    If matchRow = 0 Then
        MsgBox "No unitNum for " & unitNum & " was found in the Unit column of the Info Sheet", vbInformation
    Else
        ' Move values from Info sheet to ostWS as specified
        ostWS.Cells(5, 1).Value = infoWS.Cells(matchRow, 2).Value ' Column B to Cell A5
        ostWS.Cells(6, 1).Value = infoWS.Cells(matchRow, 3).Value ' Column C to Cell A6
        ostWS.Cells(7, 1).Value = infoWS.Cells(matchRow, 4).Value ' Column D to Cell A7
        ostWS.Cells(8, 1).Value = infoWS.Cells(matchRow, 5).Value ' Column E to Cell A8
        ostWS.Cells(9, 1).Value = infoWS.Cells(matchRow, 6).Value ' Column F to Cell A9
        ostWS.Cells(8, 11).Value = infoWS.Cells(matchRow, 4).Value ' Column D to Cell K8
        ostWS.Cells(5, 11).Value = infoWS.Cells(matchRow, 1).Value ' Column A to Cell K5
    End If
End Sub
