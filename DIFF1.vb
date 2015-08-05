Sub DIFF1()
    'Disable'
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Dim i2 As Long
    Dim i1 As Long
    i1 = 5
    For i2 = 5 To 60083
        Application.StatusBar = i2
        i1 = findRow(i1, i2)
        DoEvents
    Next i2
    'Enable'
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub

Function findRow(lastProgress As Long, i2 As Long)
    Dim found As Boolean
    Dim keyFound As Boolean
    Dim tmp As Long
    Dim i1 As Long
    Dim count As Integer
    Dim key As String
    i1 = lastProgress
    Do While i1 < 68198 And found = False
        tmp = findKey(i2, i1, lastProgress) 'Etsitään avainta
        If tmp <> 0 Then 'Avain löytyi
            If count = 0 Then lastProgress = tmp
            count = count + 1
            keyFound = True
            i1 = tmp
            found = checkRow(i1, i2, count)
             If found = False Then
                If i1 < lastProgress Then
                    i1 = i1 - 1
                Else
                    i1 = i1 + 1 'Avain lÃ¶ytyi, mutta ei riviÃ¤ -> jatketaan etsimistÃ¤
                End If
            End If
        End If
        If tmp = 0 Then Exit Do 'Avainta ei lÃ¶ydy -> poistutaan
    Loop
    If found = True Then
        findRow = i1 + 1
    Else
        If keyFound = False Then
            Call fillRow(i2)
            findRow = lastProgress
        Else
            findRow = lastProgress + 1
        End If
    End If
End Function

Function findKey(i2 As Long, i1 As Long, lastProgress As Long) As Long
    Dim ind1 As Long
    ind1 = i1
    Dim key1 As String
    Dim key2 As String
    key2 = Sheet2.Cells(i2, 1).Value
    If ind1 >= lastProgress Then
        Do While StrComp(Sheet1.Cells(ind1, 2).Value, Sheet2.Cells(i2, 2).Value) = -1
            ind1 = ind1 + 1
        Loop
        Do While ind1 < 68198
            key1 = Sheet1.Cells(ind1, 1).Value
            If key1 = key2 Then
                findKey = ind1
                Exit Function
            End If
            If StrComp(key1, key2, vbTextCompare) = 1 Or StrComp(Sheet1.Cells(ind1, 2).Value, Sheet2.Cells(i2, 2).Value) = 1 Then
                Exit Do
            End If
            ind1 = ind1 + 1
        Loop
        ind1 = lastProgress - 1
    End If
    Do While ind1 > 4
        key1 = Sheet1.Cells(ind1, 1).Value
        If key1 = key2 Then
            findKey = ind1
            Exit Function
        End If
        If StrComp(key1, key2, vbTextCompare) = -1 Or StrComp(Sheet1.Cells(ind1, 2).Value, Sheet2.Cells(i2, 2).Value) = -1 Then
            Exit Do
        End If
        ind1 = ind1 - 1
    Loop
End Function

Function checkRow(i1 As Long, i2 As Long, count As Integer) As Boolean
    If (Sheet1.Cells(i1, 2).Value = Sheet2.Cells(i2, 2).Value And Sheet1.Cells(i1, 4).Value = Sheet2.Cells(i2, 4).Value And Sheet1.Cells(i1, 7).Value = Sheet2.Cells(i2, 7).Value And Sheet1.Cells(i1, 13).Value = Sheet2.Cells(i2, 13).Value And Sheet1.Cells(i1, 14).Value = Sheet2.Cells(i2, 14).Value) Then
        checkRow = True
        If count > 1 Then Call clearRow(i2)
    ElseIf count = 1 Then
        Call checkCells(i1, i2)
    End If
End Function

Sub clearRow(i2 As Long)
    Dim range2 As String
    range2 = "A" + CStr(i2) + ":V" + CStr(i2)
    Sheet2.Range(range2).Cells.Interior.ColorIndex = xlNone
    Sheet2.Cells(i2, 23).Value = ""
End Sub

Sub fillRow(i2 As Long)
    range2 = "A" + CStr(i2) + ":V" + CStr(i2)
            Sheet2.Range(range2).Cells.Interior.ColorIndex = 44
            Sheet2.Cells(i2, 23).Value = "<- Puuttuu"
End Sub

Sub checkCells(i1 As Long, i2 As Long)
    Dim range2 As String
    range2 = "A" + CStr(i2)
    If Sheet1.Cells(i1, 1).Value <> Sheet2.Cells(i2, 1).Value Then
        Sheet2.Range(range2).Cells.Interior.ColorIndex = 3
'    Else
'        Sheet2.range(range2).Cells.Interior.ColorIndex = xlNone
    End If
    range2 = "B" + CStr(i2)
    If Sheet1.Cells(i1, 2).Value <> Sheet2.Cells(i2, 2).Value Then
        Sheet2.Range(range2).Cells.Interior.ColorIndex = 3
'    Else
'        Sheet2.range(range2).Cells.Interior.ColorIndex = xlNone
    End If
'    range2 = "C" + CStr(i2)
'    If Sheet1.Cells(i1, 3).Value <> Sheet2.Cells(i2, 3).Value Then
'        Sheet2.range(range2).Cells.Interior.ColorIndex = 3
'    Else
'        Sheet2.range(range2).Cells.Interior.ColorIndex = xlNone
'    End If
    range2 = "D" + CStr(i2)
    If Sheet1.Cells(i1, 4).Value <> Sheet2.Cells(i2, 4).Value Then
        Sheet2.Range(range2).Cells.Interior.ColorIndex = 3
'    Else
'        Sheet2.range(range2).Cells.Interior.ColorIndex = xlNone
    End If
'    range2 = "E" + CStr(i2)
'    If Sheet1.Cells(i1, 5).Value <> Sheet2.Cells(i2, 5).Value Then
'        Sheet2.range(range2).Cells.Interior.ColorIndex = 3
'    Else
'        Sheet2.range(range2).Cells.Interior.ColorIndex = xlNone
'    End If
'    range2 = "F" + CStr(i2)
'    If Sheet1.Cells(i1, 6).Value <> Sheet2.Cells(i2, 6).Value Then
'        Sheet2.range(range2).Cells.Interior.ColorIndex = 3
'    Else
'        Sheet2.range(range2).Cells.Interior.ColorIndex = xlNone
'    End If
    range2 = "G" + CStr(i2)
    If Sheet1.Cells(i1, 7).Value <> Sheet2.Cells(i2, 7).Value Then
        Sheet2.Range(range2).Cells.Interior.ColorIndex = 3
'    Else
'       Sheet2.range(range2).Cells.Interior.ColorIndex = xlNone
    End If
'    range2 = "H" + CStr(i2)
'    If Sheet1.Cells(i1, 8).Value <> Sheet2.Cells(i2, 8).Value Then
'        Sheet2.range(range2).Cells.Interior.ColorIndex = 3
'    Else
'        Sheet2.range(range2).Cells.Interior.ColorIndex = xlNone
'    End If
'    range2 = "I" + CStr(i2)
'    If Sheet1.Cells(i1, 9).Value <> Sheet2.Cells(i2, 9).Value Then
'        Sheet2.range(range2).Cells.Interior.ColorIndex = 3
'    Else
'        Sheet2.range(range2).Cells.Interior.ColorIndex = xlNone
'    End If
'    range2 = "J" + CStr(i2)
'    If Sheet1.Cells(i1, 10).Value <> Sheet2.Cells(i2, 10).Value Then
'        Sheet2.range(range2).Cells.Interior.ColorIndex = 3
'    Else
'        Sheet2.range(range2).Cells.Interior.ColorIndex = xlNone
'    End If
'    range2 = "K" + CStr(i2)
'    If Sheet1.Cells(i1, 11).Value <> Sheet2.Cells(i2, 11).Value Then
'        Sheet2.range(range2).Cells.Interior.ColorIndex = 3
'    Else
'        Sheet2.range(range2).Cells.Interior.ColorIndex = xlNone
'    End If
'    range2 = "L" + CStr(i2)
'    If Sheet1.Cells(i1, 12).Value <> Sheet2.Cells(i2, 12).Value Then
'        Sheet2.range(range2).Cells.Interior.ColorIndex = 3
'    Else
'        Sheet2.range(range2).Cells.Interior.ColorIndex = xlNone
'    End If
    range2 = "M" + CStr(i2)
    If Sheet1.Cells(i1, 13).Value <> Sheet2.Cells(i2, 13).Value Then
        Sheet2.Range(range2).Cells.Interior.ColorIndex = 3
'    Else
'        Sheet2.range(range2).Cells.Interior.ColorIndex = xlNone
    End If
    range2 = "N" + CStr(i2)
    If Sheet1.Cells(i1, 14).Value <> Sheet2.Cells(i2, 14).Value Then
        Sheet2.Range(range2).Cells.Interior.ColorIndex = 3
'    Else
'        Sheet2.range(range2).Cells.Interior.ColorIndex = xlNone
    End If
'    range2 = "O" + CStr(i2)
'    If Sheet1.Cells(i1, 15).Value <> Sheet2.Cells(i2, 15).Value Then
'        Sheet2.range(range2).Cells.Interior.ColorIndex = 3
'    Else
'        Sheet2.range(range2).Cells.Interior.ColorIndex = xlNone
'    End If
'    range2 = "P" + CStr(i2)
'    If Sheet1.Cells(i1, 16).Value <> Sheet2.Cells(i2, 16).Value Then
'        Sheet2.range(range2).Cells.Interior.ColorIndex = 3
'    Else
'        Sheet2.range(range2).Cells.Interior.ColorIndex = xlNone
'    End If
'    range2 = "Q" + CStr(i2)
'    If Sheet1.Cells(i1, 17).Value <> Sheet2.Cells(i2, 17).Value Then
'        Sheet2.range(range2).Cells.Interior.ColorIndex = 3
'    Else
'        Sheet2.range(range2).Cells.Interior.ColorIndex = xlNone
'    End If
'    range2 = "R" + CStr(i2)
'    If Sheet1.Cells(i1, 18).Value <> Sheet2.Cells(i2, 18).Value Then
'        Sheet2.range(range2).Cells.Interior.ColorIndex = 3
'    Else
'        Sheet2.range(range2).Cells.Interior.ColorIndex = xlNone
'    End If
'    range2 = "S" + CStr(i2)
'    If Sheet1.Cells(i1, 19).Value <> Sheet2.Cells(i2, 19).Value Then
'        Sheet2.range(range2).Cells.Interior.ColorIndex = 3
'    Else
'        Sheet2.range(range2).Cells.Interior.ColorIndex = xlNone
'    End If
'    range2 = "T" + CStr(i2)
'    If Sheet1.Cells(i1, 20).Value <> Sheet2.Cells(i2, 20).Value Then
'        Sheet2.range(range2).Cells.Interior.ColorIndex = 3
'    Else
'        Sheet2.range(range2).Cells.Interior.ColorIndex = xlNone
'    End If
'    range2 = "U" + CStr(i2)
'    If Sheet1.Cells(i1, 21).Value <> Sheet2.Cells(i2, 21).Value Then
'        Sheet2.range(range2).Cells.Interior.ColorIndex = 3
'    Else
'        Sheet2.range(range2).Cells.Interior.ColorIndex = xlNone
'    End If
'    range2 = "V" + CStr(i2)
'    If Sheet1.Cells(i1, 22).Value <> Sheet2.Cells(i2, 22).Value Then
'        Sheet2.range(range2).Cells.Interior.ColorIndex = 3
'    Else
'        Sheet2.range(range2).Cells.Interior.ColorIndex = xlNone
'    End If
    Sheet2.Cells(i2, 23).Value = "<- Muutoksia"
End Sub
