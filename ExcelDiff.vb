Sub ExcelDiff()
    Dim i1 As Long
    Dim i2 As Long
    Dim start1 As Long
    Dim start2 As Long
    Dim length1 As Long
    Dim length2 As Long
    Dim tmp As Long
    Call askFirstRow(start1, start2)
    Call askLastRow(length1, length2)
    If MsgBox("Are these values right?" & Chr(10) & Chr(10) & Sheet1.Name + ":" & Chr(10) & "First row: " + CStr(start1) & Chr(10) & "Last row: " + CStr(length1) & Chr(10) & Chr(10) & Sheet2.Name + ":" & Chr(10) & "First row: " + CStr(start2) & Chr(10) & "Last row: " + CStr(length2), vbYesNo, "Confirm") = vbNo Then
        Exit Sub
    End If
    'Disable'
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    i1 = start1
    For i2 = start2 To length2
        Application.StatusBar = "Stage 1 of 4 | " + Sheet2.Name + ": " + CStr(i2) + "/" + CStr(length2)
        i1 = findRow(i1, i2, Sheet1, Sheet2, start1, length1)
        DoEvents
    Next i2
    i2 = start2
    For i1 = start1 To length1
        Application.StatusBar = "Stage 2 of 4 | " + Sheet1.Name + ": " + CStr(i1) + "/" + CStr(length1)
        i2 = findRow(i2, i1, Sheet2, Sheet1, start2, length2)
        DoEvents
    Next i1
    i1 = start1
    For i2 = start2 To length2
        If StrComp(Sheet2.Cells(i2, 23).Value, "<- Change", vbTextCompare) = 0 Then
            Application.StatusBar = "Stage 3 of 4 | " + Sheet2.Name + ": " + CStr(i2) + "/" + CStr(length2)
            tmp = findKeyFiltered(i1, i2, Sheet1, Sheet2, start1, length1)
            If tmp = 0 Then
                Call fillRowNew(i2, Sheet2)
            Else
                i1 = tmp
            End If
            DoEvents
        End If
    Next i2
    i2 = start2
    For i1 = start1 To length1
        If StrComp(Sheet1.Cells(i1, 23).Value, "<- Change", vbTextCompare) = 0 Then
            Application.StatusBar = "Stage 4 of 4 | " + Sheet1.Name + ": " + CStr(i1) + "/" + CStr(length1)
            tmp = findKeyFiltered(i2, i1, Sheet2, Sheet1, start2, length2)
            If tmp = 0 Then
                Call fillRowNew(i1, Sheet1)
            Else
                i2 = tmp
            End If
            DoEvents
        End If
    Next i1
    Application.StatusBar = "Run Complete!"
    DoEvents
    'Enable'
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub

Function findRow(lastProgress As Long, sourceIndex As Long, target As Worksheet, source As Worksheet, targetStart As Long, targetLength As Long)
    Dim found As Boolean
    Dim keyFound As Boolean
    Dim tmp As Long
    Dim targetIndex As Long
    Dim count As Integer
    Dim key As String
    targetIndex = lastProgress
    Do While targetIndex <= targetLength And found = False
        tmp = findKey(targetIndex, sourceIndex, lastProgress, target, source, targetStart, targetLength) 'Etsitään avainta
        If tmp <> 0 Then 'Avain löytyi
            If count = 0 Then lastProgress = tmp
            count = count + 1
            keyFound = True
            targetIndex = tmp
            found = checkRow(targetIndex, sourceIndex, count, target, source)
             If found = False Then
                If targetIndex < lastProgress Then
                    targetIndex = targetIndex - 1
                Else
                    targetIndex = targetIndex + 1 'Avain lÃ¶ytyi, mutta ei riviÃ¤ -> jatketaan etsimistÃ¤
                End If
            End If
        End If
        If tmp = 0 Then Exit Do 'Avainta ei lÃ¶ydy -> poistutaan
    Loop
    If found = True Then
        findRow = targetIndex + 1
    Else
        If keyFound = False Then
            Call fillRow(sourceIndex, source)
            findRow = lastProgress
        Else
            findRow = lastProgress + 1
        End If
    End If
End Function

Function findKey(targetIndex As Long, sourceIndex As Long, lastProgress As Long, target As Worksheet, source As Worksheet, targetStart As Long, targetLength As Long) As Long
    Dim ind1 As Long
    ind1 = targetIndex
    Dim key1 As String
    Dim key2 As String
    key2 = source.Cells(sourceIndex, 1).Value
    If ind1 >= lastProgress Then
        Do While StrComp(target.Cells(ind1, 2).Value, source.Cells(sourceIndex, 2).Value) = -1
            ind1 = ind1 + 1
        Loop
        Do While ind1 <= targetLength
            key1 = target.Cells(ind1, 1).Value
            If key1 = key2 Then
                findKey = ind1
                Exit Function
            End If
            If StrComp(key1, key2, vbTextCompare) = 1 Or StrComp(target.Cells(ind1, 2).Value, source.Cells(sourceIndex, 2).Value) = 1 Then
                Exit Do
            End If
            ind1 = ind1 + 1
        Loop
        ind1 = lastProgress - 1
    End If
    Do While ind1 >= targetStart
        key1 = target.Cells(ind1, 1).Value
        If key1 = key2 Then
            findKey = ind1
            Exit Function
        End If
        If StrComp(key1, key2, vbTextCompare) = -1 Or StrComp(target.Cells(ind1, 2).Value, source.Cells(sourceIndex, 2).Value) = -1 Then
            Exit Do
        End If
        ind1 = ind1 - 1
    Loop
End Function

Function findKeyFiltered(targetIndex As Long, sourceIndex As Long, target As Worksheet, source As Worksheet, targetStart As Long, targetLength As Long) As Long
    Dim ind1 As Long
    ind1 = targetIndex
    Dim key1 As String
    Dim key2 As String
    key2 = source.Cells(sourceIndex, 1).Value
    Do While StrComp(target.Cells(ind1, 23).Value, "<- Change", vbTextCompare) <> 0
        ind1 = ind1 + 1
    Loop
    Do While StrComp(target.Cells(ind1, 2).Value, source.Cells(sourceIndex, 2).Value) = -1
        ind1 = ind1 + 1
        Do While StrComp(target.Cells(ind1, 23).Value, "<- Change", vbTextCompare) <> 0
        ind1 = ind1 + 1
        Loop
    Loop
    Do While ind1 <= targetLength
        If StrComp(target.Cells(ind1, 23).Value, "<- Change", vbTextCompare) = 0 Then
            key1 = target.Cells(ind1, 1).Value
            If key1 = key2 Then
                findKeyFiltered = ind1
                Exit Function
            End If
            If StrComp(key1, key2, vbTextCompare) = 1 Or StrComp(target.Cells(ind1, 2).Value, source.Cells(sourceIndex, 2).Value) = 1 Then
                Exit Do
            End If
        End If
        ind1 = ind1 + 1
    Loop
End Function

Function checkRow(i1 As Long, i2 As Long, count As Integer, target As Worksheet, source As Worksheet) As Boolean
    If (target.Cells(i1, 2).Value = source.Cells(i2, 2).Value And target.Cells(i1, 4).Value = source.Cells(i2, 4).Value And target.Cells(i1, 7).Value = source.Cells(i2, 7).Value And target.Cells(i1, 13).Value = source.Cells(i2, 13).Value And target.Cells(i1, 14).Value = source.Cells(i2, 14).Value) Then
        checkRow = True
        If count > 1 Then Call clearRow(i2, source)
    ElseIf count = 1 Then
        Call checkCells(i1, i2, target, source)
    End If
End Function

Sub clearRow(sourceIndex As Long, source As Worksheet)
    Dim range As String
    range = "A" + CStr(sourceIndex) + ":V" + CStr(sourceIndex)
    source.range(range).Cells.Interior.ColorIndex = xlNone
    source.Cells(sourceIndex, 23).Value = ""
End Sub

Sub fillRow(sourceIndex As Long, source As Worksheet)
    Dim range As String
    range = "A" + CStr(sourceIndex) + ":V" + CStr(sourceIndex)
            source.range(range).Cells.Interior.ColorIndex = 44
            source.Cells(sourceIndex, 23).Value = "<- Missing"
End Sub

Sub fillRowNew(sourceIndex As Long, source As Worksheet)
    Dim range As String
    range = "A" + CStr(sourceIndex)
    If source.range(range).Interior.ColorIndex = xlNone Then
        source.range(range).Interior.ColorIndex = 4
    End If
    range = "B" + CStr(sourceIndex)
    If source.range(range).Interior.ColorIndex = xlNone Then
        source.range(range).Interior.ColorIndex = 4
    End If
    range = "C" + CStr(sourceIndex)
    If source.range(range).Interior.ColorIndex = xlNone Then
        source.range(range).Interior.ColorIndex = 4
    End If
    range = "D" + CStr(sourceIndex)
    If source.range(range).Interior.ColorIndex = xlNone Then
        source.range(range).Interior.ColorIndex = 4
    End If
    range = "E" + CStr(sourceIndex)
    If source.range(range).Interior.ColorIndex = xlNone Then
        source.range(range).Interior.ColorIndex = 4
    End If
    range = "F" + CStr(sourceIndex)
    If source.range(range).Interior.ColorIndex = xlNone Then
        source.range(range).Interior.ColorIndex = 4
    End If
    range = "G" + CStr(sourceIndex)
    If source.range(range).Interior.ColorIndex = xlNone Then
        source.range(range).Interior.ColorIndex = 4
    End If
    range = "H" + CStr(sourceIndex)
    If source.range(range).Interior.ColorIndex = xlNone Then
        source.range(range).Interior.ColorIndex = 4
    End If
    range = "I" + CStr(sourceIndex)
    If source.range(range).Interior.ColorIndex = xlNone Then
        source.range(range).Interior.ColorIndex = 4
    End If
    range = "J" + CStr(sourceIndex)
    If source.range(range).Interior.ColorIndex = xlNone Then
        source.range(range).Interior.ColorIndex = 4
    End If
    range = "K" + CStr(sourceIndex)
    If source.range(range).Interior.ColorIndex = xlNone Then
        source.range(range).Interior.ColorIndex = 4
    End If
    range = "L" + CStr(sourceIndex)
    If source.range(range).Interior.ColorIndex = xlNone Then
        source.range(range).Interior.ColorIndex = 4
    End If
    range = "M" + CStr(sourceIndex)
    If source.range(range).Interior.ColorIndex = xlNone Then
        source.range(range).Interior.ColorIndex = 4
    End If
    range = "N" + CStr(sourceIndex)
    If source.range(range).Interior.ColorIndex = xlNone Then
        source.range(range).Interior.ColorIndex = 4
    End If
    range = "O" + CStr(sourceIndex)
    If source.range(range).Interior.ColorIndex = xlNone Then
        source.range(range).Interior.ColorIndex = 4
    End If
    range = "P" + CStr(sourceIndex)
    If source.range(range).Interior.ColorIndex = xlNone Then
        source.range(range).Interior.ColorIndex = 4
    End If
    range = "Q" + CStr(sourceIndex)
    If source.range(range).Interior.ColorIndex = xlNone Then
        source.range(range).Interior.ColorIndex = 4
    End If
    range = "R" + CStr(sourceIndex)
    If source.range(range).Interior.ColorIndex = xlNone Then
        source.range(range).Interior.ColorIndex = 4
    End If
    range = "S" + CStr(sourceIndex)
    If source.range(range).Interior.ColorIndex = xlNone Then
        source.range(range).Interior.ColorIndex = 4
    End If
    range = "T" + CStr(sourceIndex)
    If source.range(range).Interior.ColorIndex = xlNone Then
        source.range(range).Interior.ColorIndex = 4
    End If
    range = "U" + CStr(sourceIndex)
    If source.range(range).Interior.ColorIndex = xlNone Then
        source.range(range).Interior.ColorIndex = 4
    End If
    range = "V" + CStr(sourceIndex)
    If source.range(range).Interior.ColorIndex = xlNone Then
        source.range(range).Interior.ColorIndex = 4
    End If
    source.Cells(sourceIndex, 23).Value = "<- New"
End Sub

Sub checkCells(targetIndex As Long, sourceIndex As Long, target As Worksheet, source As Worksheet)
    Dim range As String
    range = "A" + CStr(sourceIndex)
    If target.Cells(targetIndex, 1).Value <> source.Cells(sourceIndex, 1).Value Then
        source.range(range).Cells.Interior.ColorIndex = 3
'    Else
'        source.range(range).Cells.Interior.ColorIndex = xlNone
    End If
    range = "B" + CStr(sourceIndex)
    If target.Cells(targetIndex, 2).Value <> source.Cells(sourceIndex, 2).Value Then
        source.range(range).Cells.Interior.ColorIndex = 3
'    Else
'        source.range(range).Cells.Interior.ColorIndex = xlNone
    End If
'    range = "C" + CStr(sourceIndex)
'    If target.Cells(targetIndex, 3).Value <> source.Cells(sourceIndex, 3).Value Then
'        source.range(range).Cells.Interior.ColorIndex = 3
'    Else
'        source.range(range).Cells.Interior.ColorIndex = xlNone
'    End If
    range = "D" + CStr(sourceIndex)
    If target.Cells(targetIndex, 4).Value <> source.Cells(sourceIndex, 4).Value Then
        source.range(range).Cells.Interior.ColorIndex = 3
'    Else
'        source.range(range).Cells.Interior.ColorIndex = xlNone
    End If
'    range = "E" + CStr(sourceIndex)
'    If target.Cells(targetIndex, 5).Value <> source.Cells(sourceIndex, 5).Value Then
'        source.range(range).Cells.Interior.ColorIndex = 3
'    Else
'        source.range(range).Cells.Interior.ColorIndex = xlNone
'    End If
'    range = "F" + CStr(sourceIndex)
'    If target.Cells(targetIndex, 6).Value <> source.Cells(sourceIndex, 6).Value Then
'        source.range(range).Cells.Interior.ColorIndex = 3
'    Else
'        source.range(range).Cells.Interior.ColorIndex = xlNone
'    End If
    range = "G" + CStr(sourceIndex)
    If target.Cells(targetIndex, 7).Value <> source.Cells(sourceIndex, 7).Value Then
        source.range(range).Cells.Interior.ColorIndex = 3
'    Else
'       source.range(range).Cells.Interior.ColorIndex = xlNone
    End If
'    range = "H" + CStr(sourceIndex)
'    If target.Cells(targetIndex, 8).Value <> source.Cells(sourceIndex, 8).Value Then
'        source.range(range).Cells.Interior.ColorIndex = 3
'    Else
'        source.range(range).Cells.Interior.ColorIndex = xlNone
'    End If
'    range = "I" + CStr(sourceIndex)
'    If target.Cells(targetIndex, 9).Value <> source.Cells(sourceIndex, 9).Value Then
'        source.range(range).Cells.Interior.ColorIndex = 3
'    Else
'        source.range(range).Cells.Interior.ColorIndex = xlNone
'    End If
'    range = "J" + CStr(sourceIndex)
'    If target.Cells(targetIndex, 10).Value <> source.Cells(sourceIndex, 10).Value Then
'        source.range(range).Cells.Interior.ColorIndex = 3
'    Else
'        source.range(range).Cells.Interior.ColorIndex = xlNone
'    End If
'    range = "K" + CStr(sourceIndex)
'    If target.Cells(targetIndex, 11).Value <> source.Cells(sourceIndex, 11).Value Then
'        source.range(range).Cells.Interior.ColorIndex = 3
'    Else
'        source.range(range).Cells.Interior.ColorIndex = xlNone
'    End If
'    range = "L" + CStr(sourceIndex)
'    If target.Cells(targetIndex, 12).Value <> source.Cells(sourceIndex, 12).Value Then
'        source.range(range).Cells.Interior.ColorIndex = 3
'    Else
'        source.range(range).Cells.Interior.ColorIndex = xlNone
'    End If
    range = "M" + CStr(sourceIndex)
    If target.Cells(targetIndex, 13).Value <> source.Cells(sourceIndex, 13).Value Then
        source.range(range).Cells.Interior.ColorIndex = 3
'    Else
'        source.range(range).Cells.Interior.ColorIndex = xlNone
    End If
    range = "N" + CStr(sourceIndex)
    If target.Cells(targetIndex, 14).Value <> source.Cells(sourceIndex, 14).Value Then
        source.range(range).Cells.Interior.ColorIndex = 3
'    Else
'        source.range(range).Cells.Interior.ColorIndex = xlNone
    End If
'    range = "O" + CStr(sourceIndex)
'    If target.Cells(targetIndex, 15).Value <> source.Cells(sourceIndex, 15).Value Then
'        source.range(range).Cells.Interior.ColorIndex = 3
'    Else
'        source.range(range).Cells.Interior.ColorIndex = xlNone
'    End If
'    range = "P" + CStr(sourceIndex)
'    If target.Cells(targetIndex, 16).Value <> source.Cells(sourceIndex, 16).Value Then
'        source.range(range).Cells.Interior.ColorIndex = 3
'    Else
'        source.range(range).Cells.Interior.ColorIndex = xlNone
'    End If
'    range = "Q" + CStr(sourceIndex)
'    If target.Cells(targetIndex, 17).Value <> source.Cells(sourceIndex, 17).Value Then
'        source.range(range).Cells.Interior.ColorIndex = 3
'    Else
'        source.range(range).Cells.Interior.ColorIndex = xlNone
'    End If
'    range = "R" + CStr(sourceIndex)
'    If target.Cells(targetIndex, 18).Value <> source.Cells(sourceIndex, 18).Value Then
'        source.range(range).Cells.Interior.ColorIndex = 3
'    Else
'        source.range(range).Cells.Interior.ColorIndex = xlNone
'    End If
'    range = "S" + CStr(sourceIndex)
'    If target.Cells(targetIndex, 19).Value <> source.Cells(sourceIndex, 19).Value Then
'        source.range(range).Cells.Interior.ColorIndex = 3
'    Else
'        source.range(range).Cells.Interior.ColorIndex = xlNone
'    End If
'    range = "T" + CStr(sourceIndex)
'    If target.Cells(targetIndex, 20).Value <> source.Cells(sourceIndex, 20).Value Then
'        source.range(range).Cells.Interior.ColorIndex = 3
'    Else
'        source.range(range).Cells.Interior.ColorIndex = xlNone
'    End If
'    range = "U" + CStr(sourceIndex)
'    If target.Cells(targetIndex, 21).Value <> source.Cells(sourceIndex, 21).Value Then
'        source.range(range).Cells.Interior.ColorIndex = 3
'    Else
'        source.range(range).Cells.Interior.ColorIndex = xlNone
'    End If
'    range = "V" + CStr(sourceIndex)
'    If target.Cells(targetIndex, 22).Value <> source.Cells(sourceIndex, 22).Value Then
'        source.range(range).Cells.Interior.ColorIndex = 3
'    Else
'        source.range(range).Cells.Interior.ColorIndex = xlNone
'    End If
    source.Cells(sourceIndex, 23).Value = "<- Change"
End Sub

Sub askFirstRow(start1 As Long, start2 As Long)
    'Dim result1 As Long
    'Dim result2 As Long
    Dim iput As Long
    Do While start1 = 0
        iput = InputBox("First row of data in " + Sheet1.Name, "First, I need to know first rows") 'The variable is assigned the value entered in the InputBox
        If iput > 0 Then 'If the value > 0 the result is valid
            MsgBox "OK, the first row in " + Sheet1.Name + " is " + CStr(iput)
            start1 = iput
        End If
    Loop
    iput = 0
    Do While start2 = 0
        iput = InputBox("First row of data in " + Sheet2.Name, "First, I need to know first rows") 'The variable is assigned the value entered in the InputBox
        If iput > 0 Then 'If the value > 0 the result is valid
            MsgBox "OK, the first row in " + Sheet2.Name + " is " + CStr(iput)
            start2 = iput
        End If
    Loop
End Sub

Sub askLastRow(length1 As Long, length2 As Long)
    Dim iput As Long
    If MsgBox("Do you want to manually set last rows? If NO, worksheets' last rows are defined automatically.", vbYesNo, "Manually set last rows?") = vbYes Then
        Do While length1 = 0
            iput = InputBox("Last row of data in " + Sheet1.Name, "Please, tell me last rows") 'The variable is assigned the value entered in the InputBox
            If iput > 0 Then 'If the value > 0 the result is valid
                MsgBox "OK, the last row in " + Sheet1.Name + " is " + CStr(iput)
                length1 = iput
            End If
        Loop
        iput = 0
        Do While length2 = 0
            iput = InputBox("Last row of data in " + Sheet2.Name, "Please, tell me last rows") 'The variable is assigned the value entered in the InputBox
            If iput > 0 Then 'If the value > 0 the result is valid
                MsgBox "OK, the last row in " + Sheet2.Name + " is " + CStr(iput)
                length2 = iput
            End If
        Loop
    Else
        length1 = Sheet1.UsedRange.Row - 1 + Sheet1.UsedRange.Rows.count
        length2 = Sheet2.UsedRange.Row - 1 + Sheet2.UsedRange.Rows.count
    End If
End Sub
