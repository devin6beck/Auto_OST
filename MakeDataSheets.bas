Function SheetExists(sheetName As String, wb As Workbook) As Boolean
    Dim tmpSheet As Worksheet
    On Error Resume Next
    Set tmpSheet = wb.Sheets(sheetName)
    On Error GoTo 0
    SheetExists = Not tmpSheet Is Nothing
End Function

Sub MakeDataSheets()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet, newWs As Worksheet
    Dim lastRow As Long, lastColumn As Long, i As Long, j As Long
    Dim headerText As String, otownValue As String
    Dim uniqueData() As Variant
    Dim dataCount As Long: dataCount = 0
    Dim validSheetName As String

    ' Locate the specific pattern-matched worksheet
    For Each ws In ThisWorkbook.Worksheets
        If LCase(ws.Name) Like "*_ost_data" Then
            Exit For
        End If
    Next ws

    ' Check if the specific worksheet was found
    If ws Is Nothing Then
        MsgBox "Pattern-matched sheet not found.", vbCritical
        Exit Sub
    End If

    ' Determine the size of the data
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    lastColumn = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column

    ' Find the column with the "OTOWN" header
    Dim otownCol As Integer: otownCol = 0
    For i = 1 To lastColumn
        headerText = Trim(ws.Cells(1, i).value)
        If LCase(headerText) = "otown" Then
            otownCol = i
            Exit For
        End If
    Next i

    ' Ensure the "OTOWN" column exists
    If otownCol = 0 Then
        MsgBox "Necessary header 'OTOWN' not found.", vbExclamation
        Exit Sub
    End If

    ' Collect unique OTOWN values
    ReDim uniqueData(1 To 1, 1 To 100)
    For i = 2 To lastRow
        otownValue = Trim(ws.Cells(i, otownCol).value)
        If Not IsInArray(otownValue, uniqueData, dataCount) Then
            If dataCount >= UBound(uniqueData, 2) Then
                ReDim Preserve uniqueData(1 To 1, 1 To dataCount + 100)
            End If
            dataCount = dataCount + 1
            uniqueData(1, dataCount) = otownValue
        End If
    Next i

    ' Create or get existing sheets based on unique OTOWN and copy relevant data
    For i = 1 To dataCount
        otownValue = uniqueData(1, i)
        validSheetName = Left(otownValue, 31) ' Excel sheet name limitation
        If Not SheetExists(validSheetName & " Data", ThisWorkbook) Then
            Set newWs = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
            newWs.Name = validSheetName & " Data"
            ws.Rows(1).Copy Destination:=newWs.Rows(1)  ' Copy header row

            ' Copy rows corresponding to the current OTOWN value
            For j = 2 To lastRow
                If Trim(ws.Cells(j, otownCol).value) = otownValue Then
                    ws.Rows(j).Copy Destination:=newWs.Cells(newWs.Rows.count, 1).End(xlUp).Offset(1, 0)
                End If
            Next j
        Else
            MsgBox "Sheet with name '" & validSheetName & " Data' already exists.", vbExclamation
        End If

        Set newWs = Nothing
    Next i

    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
    Resume Next
End Sub

Function IsInArray(valToBeFound As Variant, arr As Variant, ByVal uboundVal As Long) As Boolean
    Dim element As Long
    For element = 1 To uboundVal
        If arr(1, element) = valToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next element
    IsInArray = False
End Function


