
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
    Dim otCodeCol As Integer, otDescCol As Integer, otDebitCol As Integer, otCreditCol As Integer, otDateCol As Integer
    Dim otCode As String, otDesc As String, otDate As String, matched As Boolean
    Dim debitValue As Double, creditValue As Double
    Dim credits() As Double, debits() As Double, descriptions() As String, codes() As String, dates() As String
    ' Variables for sums and counts
    Dim stayOverDebitSum As Double, stayOverCount As Long
    Dim departureDebitSum As Double, departureCount As Long
    Dim trashDebitSum As Double, trashCount As Long
    Dim ohmcrfSum As Double, ohmcrfCount As Long
    Dim checkSum As Double, commisSum As Double
    Dim wochrgSum As Double, ownstaySum As Double
    Dim taxgrtSum As Double, incomeSum As Double
    Dim invpurSum As Double, omntlbSum As Double
    Dim omntptSum As Double, omntptCount As Long
    Dim omrkupSum As Double, preventmaintSum As Double
    Dim depclnSum As Double, depclnCount As Long
    Dim ownlsbSum As Double, trdownSum As Double
    Dim ownffdSum As Double, ownffcSum As Double, ownffcDate As String
    Dim pgassnSum As Double, pycashSum As Double
    Dim pycheckSum As Double, pycheckDate As String
    Dim reimboCreditSum As Double, reimboDebitSum As Double
    Dim prclnSum As Double, prcclnCount As Long
    For i = 1 To dataWs.Cells(1, dataWs.Columns.count).End(xlToLeft).Column
        Select Case True
            Case Left(dataWs.Cells(1, i).value, 6) = "OTCODE": otCodeCol = i
            Case Left(dataWs.Cells(1, i).value, 9) = "OTDESCRIP": otDescCol = i
            Case Left(dataWs.Cells(1, i).value, 7) = "OTDEBIT": otDebitCol = i
            Case Left(dataWs.Cells(1, i).value, 8) = "OTCREDIT": otCreditCol = i
            Case Left(dataWs.Cells(1, i).value, 6) = "OTDATE": otDateCol = i
        End Select
    Next i
    lastRow = dataWs.Cells(dataWs.Rows.count, otCodeCol).End(xlUp).Row
    ReDim debits(1 To lastRow), descriptions(1 To lastRow), codes(1 To lastRow), credits(1 To lastRow), dates(1 To lastRow) ' Include credits array
    ' Load data into arrays
    For i = 2 To lastRow
        debits(i) = dataWs.Cells(i, otDebitCol).value
        credits(i) = dataWs.Cells(i, otCreditCol).value ' Load credit values
        descriptions(i) = LCase(dataWs.Cells(i, otDescCol).value)
        codes(i) = dataWs.Cells(i, otCodeCol).value
        dates(i) = dataWs.Cells(i, otDateCol).value
        
    Next i
    ' Process rows for sums and counts
    For i = 2 To lastRow
        debitValue = debits(i)
        creditValue = credits(i) ' Assign the credit value from the array
        otDesc = descriptions(i)
        otCode = codes(i)
        otDate = dates(i)
        
    ' Process based on OTCODE and OTDESCRIP
        If otCode = "CLEAN" Or otCode = "TNTCLN" Or otCode = "STYCLN" Or otCode = "DPPCLN" Then
            If InStr(otDesc, "stayover") > 0 Then
                MsgBox "You made iT! the otCode is: " & otCode & "otDesc is: " & otDesc
                stayOverDebitSum = stayOverDebitSum + debitValue
                If debitValue > 0 Then stayOverCount = stayOverCount + 1
                If debitValue < 0 Then stayOverCount = stayOverCount - 1
            ElseIf InStr(otDesc, "departure") > 0 Then
                MsgBox "You made iT! the otCode is: " & otCode & "otDesc is: " & otDesc
                departureDebitSum = departureDebitSum + debitValue
                If debitValue > 0 Then departureCount = departureCount + 1
                If debitValue < 0 Then departureCount = departureCount - 1
            ElseIf InStr(otDesc, "trash") > 0 Then
                MsgBox "You made iT! the otCode is: " & otCode & "otDesc is: " & otDesc
                trashDebitSum = trashDebitSum + debitValue
                If debitValue > 0 Then trashCount = trashCount + 1
                If debitValue < 0 Then trashCount = trashCount - 1
            Else
                MsgBox "Unknown description found for CLEAN code. Description is '" & otDesc & "'. This transaction will NOT be moved to contract: " & ostWS.Name & " from contract." & dataWs.Name & "Contact Devin to add description."
            End If
        Else
            Select Case otCode
                Case "OHMCRF"
                    ohmcrfDebitSum = ohmcrfDebitSum + debitValue
                    If debitValue > 0 Then ohmcrfCount = ohmcrfCount + 1
                    If debitValue < 0 Then ohmcrfCount = ohmcrfCount - 1
                Case "CHECK"
                    checkDebitSum = checkDebitSum + debitValue
                Case "COMMIS"
                    commisDebitSum = commisDebitSum + debitValue
                Case "WOCHRG"
                    wochrgDebitSum = wochrgDebitSum + debitValue
                Case "OAXFD"
                    ownstayDebitSum = ownstayDebitSum + debitValue
                Case "TAXGRT"
                    taxgrtDebitSum = taxgrtDebitSum + debitValue
                Case "INCOME"
                    incomeCreditSum = incomeCreditSum + creditValue
                Case "INVPUR"
                    invpurDebitSum = invpurDebitSum + debitValue
                Case "OMNTLB"
                    omntlbDebitSum = omntlbDebitSum + debitValue
                Case "OMNTPT"
                    omntptDebitSum = omntptDebitSum + debitValue
                    If debitValue > 0 Then omntptCount = omntptCount + 1
                    If debitValue < 0 Then omntptCount = omntptCount - 1
                Case "OMRKUP"
                    omrkupDebitSum = omrkupDebitSum + debitValue
                Case "PMFEE"
                    preventmaintDebitSum = preventmaintDebitSum + debitValue
                Case "DEPCLN"
                    depclnDebitSum = depclnDebitSum + debitValue
                    If debitValue > 0 Then depclnCount = depclnCount + 1
                    If debitValue < 0 Then depclnCount = depclnCount - 1
                Case "OWNLSB"
                    ownlsbCreditSum = ownlsbCreditSum + creditValue
                Case "TRDOWN"
                    trdownDebitSum = trdownDebitSum + debitValue
                Case "OWNFFD"
                    ownffdDebitSum = ownffdDebitSum + debitValue
                Case "OWNFFC"
                    ownffcCreditSum = ownffcCreditSum + creditValue
                    ownffcDate = otDate
                Case "PGASSN"
                    pgassnDebitSum = pgassnDebitSum + debitValue
                Case "PYCASH"
                    pycashCreditSum = pycashCreditSum + creditValue
                Case "PYCHCK"
                    pycheckCreditSum = pycheckCreditSum + creditValue
                    pycheckDate = otDate
                Case "REIMBO"
                    reimboCreditSum = reimboCreditSum + creditValue
                    reimboDebitSum = reimboDebitSum + debitValue
                Case "PRCLN"
                    prclnDebitSum = prclnDebitSum + debitValue
                    If debitValue > 0 Then prcclnCount = prcclnCount + 1
                    If debitValue < 0 Then prcclnCount = prcclnCount - 1
                Case Else
                    MsgBox "Unrecognized OTCODE found: " & otCode & " with description: " & otDesc
            End Select
        End If

        ' Output results to ostWS
        ostWS.Cells(32, 12).value = stayOverDebitSum ' Change as required, no definition in snippet
        ostWS.Cells(32, 4).value = stayOverCount ' Change as required, no definition in snippet
        ostWS.Cells(34, 12).value = departureDebitSum ' Change as required, no definition in snippet
        ostWS.Cells(34, 4).value = departureCount ' Change as required, no definition in snippet
        ostWS.Cells(31, 12).value = trashDebitSum ' Change as required, no definition in snippet
        ostWS.Cells(31, 4).value = trashCount ' Change as required, no definition in snippet
        ostWS.Cells(16, 12).value = ohmcrfDebitSum
        ostWS.Cells(41, 12).value = taxgrtDebitSum
        ostWS.Cells(10, 12).value = incomeCreditSum
        ostWS.Cells(33, 12).value = depclnDebitSum
        ostWS.Cells(33, 4).value = depclnCount
        ostWS.Cells(11, 12).value = ownlsbCreditSum
        ' Placeholder for checkSum but might not ever use
        ' Placeholder for invpurSum but might not ever use
        ostWS.Cells(12, 12).value = -commisDebitSum ' Convert to negative
        ostWS.Cells(45, 12).value = ownstayDebitSum
        ostWS.Cells(19, 12).value = wochrgDebitSum
        ostWS.Cells(20, 12).value = omntlbDebitSum
        ostWS.Cells(21, 12).value = omntptDebitSum
        ostWS.Cells(21, 4).value = omntptCount
        ostWS.Cells(22, 12).value = omrkupDebitSum
        ostWS.Cells(23, 12).value = preventmaintDebitSum
        ostWS.Cells(49, 12).value = trdownDebitSum
        ostWS.Cells(55, 12).value = ownffdDebitSum
        ostWS.Cells(61, 12).value = -ownffcCreditSum
        ostWS.Cells(61, 1).value = "Fee Reserve Payment"
        ostWS.Cells(61, 9).value = "Transaction Date: " & ownffcDate
        ostWS.Cells(53, 12).value = pgassnDebitSum
        ostWS.Cells(30, 12).value = prclnDebitSum
        ostWS.Cells(30, 4).value = prcclnCount
        ' Placeholder for pycashSum but might not ever use
        ostWS.Cells(60, 12).value = -pycheckCreditSum
        ostWS.Cells(60, 1).value = "EFT/Check Payment"
        ostWS.Cells(60, 9).value = "Transaction Date: " & pycheckDate
        ' Placeholder for reimboDebitSum or reimboCreditSum but might not ever use
    Next i
End Sub

Function IsInArray(valToBeFound As Variant, arr As Variant) As Boolean
    IsInArray = Not IsError(Application.Match(valToBeFound, arr, 0))
End Function

