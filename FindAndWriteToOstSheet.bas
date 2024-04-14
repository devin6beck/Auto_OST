Sub FindAndWriteToOstSheet()
    On Error GoTo ErrorHandler
    Dim infoSheet As Worksheet, ostWS As Worksheet, codesSheet As Worksheet
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
                Set ostWS = ThisWorkbook.Sheets(ostSheetName)
                If Not ostWS Is Nothing Then
                    ostWS.Range("A1").value = "Found " & ostSheetName
                    InfoValuesToOstSheet ostWS, unitNumber
                    dataValuesToOstSheet ws, ostWS
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
    lastRow = infoWS.Cells(infoWS.Rows.count, 1).End(xlUp).Row
    
    Dim matchRow As Long
    matchRow = 0 ' Initialize with 0 to denote no match found initially
    
    ' Loop through column A to find the matching unitNum
    Dim i As Long
    For i = 1 To lastRow
        If infoWS.Cells(i, 1).value = unitNum Then
            matchRow = i
            Exit For
        End If
    Next i
    
    ' Check if a matching row was found
    If matchRow = 0 Then
        MsgBox "No unitNum for " & unitNum & " was found in the Unit column of the Info Sheet", vbInformation
    Else
        ' Move values from Info sheet to ostWS as specified
        ostWS.Cells(1, 1).value = infoWS.Cells(matchRow, 2).value ' Column B to Cell A1
        ostWS.Cells(2, 1).value = infoWS.Cells(matchRow, 3).value ' Column C to Cell A2
        ostWS.Cells(3, 1).value = infoWS.Cells(matchRow, 4).value ' Column D to Cell A3
        ostWS.Cells(4, 1).value = infoWS.Cells(matchRow, 5).value ' Column E to Cell A4
        ostWS.Cells(5, 1).value = infoWS.Cells(matchRow, 6).value ' Column F to Cell A5
        ostWS.Cells(4, 12).value = infoWS.Cells(matchRow, 4).value ' Column D to Cell K4
        ostWS.Cells(1, 12).value = infoWS.Cells(matchRow, 1).value ' Column A to Cell K1
    End If
End Sub

Sub dataValuesToOstSheet(dataWs As Worksheet, ostWS As Worksheet)
    Dim lastRow As Long
    Dim i As Long
    Dim otcCodeCol As Integer, otcDescCol As Integer, otcDebitCol As Integer, otcCreditCol As Integer
    Dim otcCode As String, otcDesc As String
    Dim sumTotal As Double, debitValue As Double, creditValue As Double
    Dim count As Long

    ' Find column indices
    For i = 1 To dataWs.Cells(1, dataWs.Columns.count).End(xlToLeft).Column
        If Left(dataWs.Cells(1, i).value, 6) = "OTCODE" Then otcCodeCol = i
        If Left(dataWs.Cells(1, i).value, 9) = "OTDESCRIP" Then otcDescCol = i
        If Left(dataWs.Cells(1, i).value, 7) = "OTDEBIT" Then otcDebitCol = i
        If Left(dataWs.Cells(1, i).value, 8) = "OTCREDIT" Then otcCreditCol = i
    Next i

    ' Initialize sums and counts
    Dim ohmcrfSum As Double, taxgrtSum As Double, incomeSum As Double, depclnSum As Double, ownlsbSum As Double
    Dim ohmcrfCount As Long, depclnCount As Long

    ' Process rows
    lastRow = dataWs.Cells(dataWs.Rows.count, otcCodeCol).End(xlUp).Row
    For i = 2 To lastRow
        otcCode = dataWs.Cells(i, otcCodeCol).value
        otcDesc = dataWs.Cells(i, otcDescCol).value
        debitValue = dataWs.Cells(i, otcDebitCol).value
        creditValue = dataWs.Cells(i, otcCreditCol).value
        
        Select Case otcCode
            Case "OHMCRF"
                ohmcrfSum = ohmcrfSum + debitValue
                ohmcrfCount = ohmcrfCount + 1 ' Adjust based on refunds if required
            Case "TAXGRT"
                taxgrtSum = taxgrtSum + debitValue
            Case "INCOME"
                incomeSum = incomeSum + debitValue
            Case "DEPCLN"
                depclnSum = depclnSum + debitValue
                depclnCount = depclnCount + 1 ' Adjust based on refunds if required
            Case "OWNLSB"
                ownlsbSum = ownlsbSum + creditValue
            Case Else
                If otcCode <> "CLEAN" And otcCode <> "TNTCLN" And otcCode <> "STYCLN" And otcCode <> "DPPCLN" Then
                    MsgBox "Unrecognized OTCODE found: " & otcCode
                End If
        End Select
    Next i

    ' Output results to ostWS
    ostWS.Cells(15, 12).value = ohmcrfSum
    ostWS.Cells(15, 4).value = ohmcrfCount
    ostWS.Cells(40, 12).value = taxgrtSum
    ostWS.Cells(10, 12).value = incomeSum
    ostWS.Cells(33, 12).value = depclnSum
    ostWS.Cells(33, 4).value = depclnCount
    ostWS.Cells(10, 12).value = ownlsbSum

End Sub

Function IsInArray(valToBeFound As Variant, arr As Variant) As Boolean
    IsInArray = Not IsError(Application.Match(valToBeFound, arr, 0))
End Function


