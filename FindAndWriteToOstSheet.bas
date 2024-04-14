
Sub FindAndWriteToOstSheet()
    On Error GoTo ErrorHandler
    Dim infoSheet As Worksheet, ostWS As Worksheet, codesSheet As Worksheet
    Dim ws As Worksheet
    Dim foundOST As Boolean
    Dim dataSheetName As String, ostSheetName As String, otOwn As String
    
    Set infoSheet = ThisWorkbook.Sheets("Info")
    Set codesSheet = ThisWorkbook.Sheets("Codes")
    
    For Each ws In ThisWorkbook.Sheets
        If ws.Name Like "* Data" And Not ws.Name Like "*_Data" Then
            dataSheetName = ws.Name
            otOwn = Left(dataSheetName, InStr(dataSheetName, " Data") - 1)  ' Changed from unitNumber to otOwn
            ostSheetName = Replace(dataSheetName, " Data", " OST")
            foundOST = SheetExists(ostSheetName)
            MsgBox "Owner Contract is: " & otOwn  ' Updated message text
            If foundOST Then
                Set ostWS = ThisWorkbook.Sheets(ostSheetName)
                If Not ostWS Is Nothing Then
                    ostWS.Range("A1").value = "Found " & ostSheetName
                    InfoValuesToOstSheet ostWS, otOwn  ' Changed argument from unitNumber to otOwn
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
Sub InfoValuesToOstSheet(ostWS As Worksheet, otOwn As String)
    Dim infoWS As Worksheet
    Set infoWS = ThisWorkbook.Sheets("Info")
    
    Dim lastRow As Long
    lastRow = infoWS.Cells(infoWS.Rows.count, 1).End(xlUp).Row
    
    Dim matchRow As Long
    matchRow = 0 ' Initialize with 0 to denote no match found initially
    
    ' Loop through column A to find the matching otOwn
    Dim i As Long
    For i = 1 To lastRow
        If infoWS.Cells(i, 8).value = otOwn Then
            matchRow = i
            Exit For
        End If
    Next i
    
    ' Check if a matching row was found
    If matchRow = 0 Then
        MsgBox "No otOwn for " & otOwn & " was found in the Contract column of the Info Sheet", vbInformation
    Else
        ' Move values from Info sheet to ostWS as specified
        ostWS.Cells(1, 1).value = infoWS.Cells(matchRow, 2).value ' Column B to Cell A1
        ostWS.Cells(2, 1).value = infoWS.Cells(matchRow, 3).value ' Column C to Cell A2
        ostWS.Cells(3, 1).value = infoWS.Cells(matchRow, 4).value ' Column D to Cell A3
        ostWS.Cells(4, 1).value = infoWS.Cells(matchRow, 5).value ' Column E to Cell A4
        ostWS.Cells(5, 1).value = infoWS.Cells(matchRow, 6).value ' Column F to Cell A5
        ostWS.Cells(4, 12).value = infoWS.Cells(matchRow, 7).value ' Column G to Cell K4
        ostWS.Cells(1, 12).value = infoWS.Cells(matchRow, 1).value ' Column A to Cell K1
    End If
End Sub
Sub dataValuesToOstSheet(dataWs As Worksheet, ostWS As Worksheet)
    Dim lastRow As Long, i As Long, j As Long
    Dim otcCodeCol As Integer, otcDescCol As Integer, otcDebitCol As Integer, otcCreditCol As Integer
    Dim otcCode As String, otcDesc As String, matched As Boolean
    Dim debitValue As Double, creditValue As Double
    Dim credits() As Variant ' Declare the credits array
    ' Additional sums and counts declaration as before
    ' Find column indices
    For i = 1 To dataWs.Cells(1, dataWs.Columns.count).End(xlToLeft).Column
        Select Case True
            Case Left(dataWs.Cells(1, i).value, 6) = "OTCODE": otcCodeCol = i
            Case Left(dataWs.Cells(1, i).value, 9) = "OTDESCRIP": otcDescCol = i
            Case Left(dataWs.Cells(1, i).value, 7) = "OTDEBIT": otcDebitCol = i
            Case Left(dataWs.Cells(1, i).value, 8) = "OTCREDIT": otcCreditCol = i
        End Select
    Next i
    lastRow = dataWs.Cells(dataWs.Rows.count, otcCodeCol).End(xlUp).Row
    ReDim debits(1 To lastRow), descriptions(1 To lastRow), codes(1 To lastRow), credits(1 To lastRow) ' Include credits array
    ' Load data into arrays
    For i = 2 To lastRow
        debits(i) = dataWs.Cells(i, otcDebitCol).value
        credits(i) = dataWs.Cells(i, otcCreditCol).value ' Load credit values
        descriptions(i) = LCase(dataWs.Cells(i, otcDescCol).value)
        codes(i) = dataWs.Cells(i, otcCodeCol).value
    Next i
    ' Process rows for sums and counts
    For i = 2 To lastRow
        debitValue = debits(i)
        creditValue = credits(i) ' Assign the credit value from the array
        otcDesc = descriptions(i)
        otcCode = codes(i)
        matched = False
        
        ' Check for matching negative debit within the same group
        If debitValue > 0 Then
            For j = 2 To lastRow
                If debits(j) = -debitValue And descriptions(j) = otcDesc And codes(j) = otcCode Then
                    matched = True
                    debits(j) = 0 ' Mark as matched to avoid recounting
                    Exit For
                End If
            Next j
        End If
        ' Process based on OTCODE and OTDESCRIP
        If Not matched Or debitValue < 0 Then ' Include unmatched debits and all credits
            Select Case otcCode
                Case "CLEAN", "TNTCLN", "STYCLN", "DPPCLN"
                    If InStr(otcDesc, "stayover") > 0 Then
                        stayOverSum = stayOverSum + debitValue
                        If debitValue > 0 Then stayOverCount = stayOverCount + 1
                    ElseIf InStr(otcDesc, "departure") > 0 Then
                        departureSum = departureSum + debitValue
                        If debitValue > 0 Then departureCount = departureCount + 1
                    ElseIf InStr(otcDesc, "trash") > 0 Then
                        trashSum = trashSum + debitValue
                        If debitValue > 0 Then trashCount = trashCount + 1
                    End If
                Case Else
                    ' Generic grouping by OTCODE
                    Select Case otcCode
                        Case "OHMCRF"
                            ohmcrfSum = ohmcrfSum + debitValue
                            If debitValue > 0 Then ohmcrfCount = ohmcrfCount + 1
                        Case "TAXGRT"
                            taxgrtSum = taxgrtSum + debitValue
                        Case "INCOME"
                            incomeSum = incomeSum + creditValue
                        Case "DEPCLN"
                            depclnSum = depclnSum + debitValue
                            If debitValue > 0 Then depclnCount = depclnCount + 1
                        Case "OWNLSB"
                            ownlsbSum = ownlsbSum + creditValue
                        Case Else
                            MsgBox "Unrecognized OTCODE found: " & otcCode
                    End Select
            End Select
        End If
    Next i
    ' Output results to ostWS
    ostWS.Cells(32, 12).value = stayOverSum
    ostWS.Cells(32, 4).value = stayOverCount
    ostWS.Cells(34, 12).value = departureSum
    ostWS.Cells(34, 4).value = departureCount
    ostWS.Cells(31, 12).value = trashSum
    ostWS.Cells(31, 4).value = trashCount
    ' Additional outputs for other codes
    ostWS.Cells(16, 12).value = ohmcrfSum
    ostWS.Cells(41, 12).value = taxgrtSum
    ostWS.Cells(10, 12).value = incomeSum
    ostWS.Cells(33, 12).value = depclnSum
    ostWS.Cells(33, 4).value = depclnCount
    ostWS.Cells(11, 12).value = ownlsbSum
End Sub


Function IsInArray(valToBeFound As Variant, arr As Variant) As Boolean
    IsInArray = Not IsError(Application.Match(valToBeFound, arr, 0))
End Function

