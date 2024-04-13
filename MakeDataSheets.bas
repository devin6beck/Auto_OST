Sub MakeDataSheets()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet, foundWs As Worksheet
    Dim wsNamePattern As String
    Dim patternMatched As Boolean: patternMatched = False
    Dim headerRow As Range, dataRow As Range, pasteRange As Range
    Dim lastRow As Long, lastColumn As Long, i As Long
    Dim otunitVal As String
    Dim uniqueOtunits As Collection: Set uniqueOtunits = New Collection
    
    ' Loop through each worksheet to find the correct sheet
    For Each ws In ThisWorkbook.Worksheets
        wsNamePattern = "*" & "_OST_Data"
        If LCase(ws.Name) Like LCase(wsNamePattern) Then
            If Not patternMatched Then
                Set foundWs = ws
                patternMatched = True
            Else
                MsgBox "Error - Duplicate OST_Data sheets.", vbCritical
                Exit Sub
            End If
        End If
    Next ws
    
    If Not patternMatched Then
        MsgBox "Error - (Month)_(Year)_OST_Data sheet not found", vbCritical
        Exit Sub
    End If
    
    Set ws = foundWs

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    lastColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Set headerRow = ws.Rows(1)
    
    ' Build a collection of unique OTUNIT values
    Dim otunitColumn As Integer
    On Error Resume Next ' Ignore error if Find returns Nothing
    otunitColumn = headerRow.Find("OTUNIT", LookIn:=xlValues, LookAt:=xlWhole).Column
    On Error GoTo ErrorHandler ' Resume normal error handling
    If otunitColumn = 0 Then
        MsgBox "Header 'OTUNIT' not found in sheet " & ws.Name, vbExclamation
        Exit Sub
    End If
    
    For i = 2 To lastRow
        otunitVal = Trim(ws.Cells(i, otunitColumn).value)
        If otunitVal = "" Then otunitVal = "No OTUNIT" ' Default value for empty cells
        On Error Resume Next ' Ignore errors when adding duplicate keys
        uniqueOtunits.Add otunitVal, CStr(otunitVal)
        On Error GoTo ErrorHandler ' Resume normal error handling
    Next i
    
    ' Process each unique OTUNIT
    Dim uniqueIndex As Variant
    For Each uniqueIndex In uniqueOtunits
        ' Create a new sheet for this OTUNIT or get the existing one
        Set newWs = Nothing
        On Error Resume Next ' Ignore error if sheet doesn't exist
        Set newWs = ThisWorkbook.Sheets(uniqueIndex + " Data")
        On Error GoTo ErrorHandler ' Resume normal error handling
        

        If newWs Is Nothing Then ' If sheet doesn't exist, create it
            Set newWs = ThisWorkbook.Worksheets.Add(after:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
            newWs.Name = uniqueIndex & " Data" ' Updated to add " Data" at the end of the sheet name
            headerRow.Copy Destination:=newWs.Rows(1) ' Copy headers to the new sheet
        End If

        ' Copy data for this OTUNIT
        For i = 2 To lastRow
            If Trim(ws.Cells(i, otunitColumn).value) = uniqueIndex Then
                Set dataRow = ws.Range(ws.Cells(i, 1), ws.Cells(i, lastColumn))
                Set pasteRange = newWs.Cells(newWs.Rows.Count, "A").End(xlUp).Offset(1, 0)
                dataRow.Copy Destination:=pasteRange
            End If
        Next i
        
    Next uniqueIndex

    Exit Sub

ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
End Sub




