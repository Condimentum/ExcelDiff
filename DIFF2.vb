Sub DIFF2()
    'Disable'
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Dim i1 As Long
    Dim i2 As Long
    i2 = 5
    For i1 = 5 To 68197
        Application.StatusBar = i1
        i2 = findRow(i2, i1)
        DoEvents
    Next i1
    'Enable'
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub

Function findRow(lastProgress As Long, i1 As Long)
    Dim found As Boolean
    Dim keyFound As Boolean
    Dim tmp As Long
    Dim i2 As Long
    Dim count As Integer
    Dim key As String
    i2 = lastProgress
    key = Sheet1.Cells(i1, 1)
    Do While i2 < 69448 And found = False
        tmp = findKey(key, i2, lastProgress) 'EtsitÃ¤Ã¤n avainta
        If tmp <> 0 Then 'Avain lÃ¶ytyi
            If count = 0 Then lastProgress = tmp
            count = count + 1
            keyFound = True
            i2 = tmp
            found = checkRow(i1, i2, count)
            If found = False Then
                If i2 < lastProgress Then
                    i2 = i2 - 1
                Else
                    i2 = i2 + 1 'Avain lÃ¶ytyi, mutta ei riviÃ¤ -> jatketaan etsimistÃ¤
                End If
            End If
        End If
        If tmp = 0 Then Exit Do 'Avainta ei lÃ¶ydy -> poistutaan
    Loop
    If found = True Then
        findRow = i2 + 1
    Else
        If keyFound = False Then
            Call fillRow(i1)
            findRow = lastProgress
        Else
            findRow = lastProgress + 1
        End If
    End If
End Function
    
Function findKey(key1 As String, i2 As Long, lastProgress As Long) As Long
    Dim ind2 As Long
    ind2 = i2
    Dim key2 As String
    If ind2 >= lastProgress Then
        Do While ind2 < 69448
            key2 = Sheet2.Cells(ind2, 1).Value
            If key2 = key1 Then
                findKey = ind2
                Exit Function
            End If
            If StrComp(key2, key1, vbTextCompare) = 1 Then Exit Do
            ind2 = ind2 + 1
        Loop
        ind2 = lastProgress - 1
    End If
    Do While ind2 > 4
        key2 = Sheet2.Cells(ind2, 1).Value
        If key2 = key1 Then
            findKey = ind2
            Exit Function
        End If
        If StrComp(key2, key1, vbTextCompare) = -1 Then Exit Do
        ind2 = ind2 - 1
    Loop
End Function

Function checkRow(i1 As Long, i2 As Long, count As Integer) As Boolean
    If (Sheet1.Cells(i1, 2).Value = Sheet2.Cells(i2, 2).Value And Sheet1.Cells(i1, 4).Value = Sheet2.Cells(i2, 4).Value And Sheet1.Cells(i1, 7).Value = Sheet2.Cells(i2, 7).Value And Sheet1.Cells(i1, 13).Value = Sheet2.Cells(i2, 13).Value And Sheet1.Cells(i1, 14).Value = Sheet2.Cells(i2, 14).Value) Then
        checkRow = True
        If count > 1 Then Call clearRow(i1)
    ElseIf count = 1 Then
        Call checkCells(i1, i2)
    End If
End Function

Sub clearRow(i1 As Long)
    Dim range1 As String
    range1 = "A" + CStr(i1) + ":V" + CStr(i1)
    Sheet1.range(range1).Cells.Interior.ColorIndex = xlNone
    Sheet1.Cells(i1, 23).Value = ""
End Sub

Sub fillRow(i1 As Long)
    range1 = "A" + CStr(i1) + ":V" + CStr(i1)
            Sheet1.range(range1).Cells.Interior.ColorIndex = 44
            Sheet1.Cells(i1, 23).Value = "<-"
End Sub

Sub checkCells(i1 As Long, i2 As Long)
    Dim range1 As String
    range1 = "A" + CStr(i1)
    If Sheet1.Cells(i1, 1).Value <> Sheet2.Cells(i2, 1).Value Then
        Sheet1.range(range1).Cells.Interior.ColorIndex = 3
    Else
        Sheet1.range(range1).Cells.Interior.ColorIndex = xlNone
    End If
    range1 = "B" + CStr(i1)
    If Sheet1.Cells(i1, 2).Value <> Sheet2.Cells(i2, 2).Value Then
        Sheet1.range(range1).Cells.Interior.ColorIndex = 3
    Else
        Sheet1.range(range1).Cells.Interior.ColorIndex = xlNone
    End If
'    range1 = "C" + CStr(i1)
'    If Sheet1.Cells(i1, 3).Value <> Sheet2.Cells(i2, 3).Value Then
'        Sheet1.range(range1).Cells.Interior.ColorIndex = 3
'    Else
'        Sheet1.range(range1).Cells.Interior.ColorIndex = xlNone
'    End If
    range1 = "D" + CStr(i1)
    If Sheet1.Cells(i1, 4).Value <> Sheet2.Cells(i2, 4).Value Then
        Sheet1.range(range1).Cells.Interior.ColorIndex = 3
    Else
        Sheet1.range(range1).Cells.Interior.ColorIndex = xlNone
    End If
'    range1 = "E" + CStr(i1)
'    If Sheet1.Cells(i1, 5).Value <> Sheet2.Cells(i2, 5).Value Then
'        Sheet1.range(range1).Cells.Interior.ColorIndex = 3
'    Else
'        Sheet1.range(range1).Cells.Interior.ColorIndex = xlNone
'    End If
'    range1 = "F" + CStr(i1)
'    If Sheet1.Cells(i1, 6).Value <> Sheet2.Cells(i2, 6).Value Then
'        Sheet1.range(range1).Cells.Interior.ColorIndex = 3
'    Else
'        Sheet1.range(range1).Cells.Interior.ColorIndex = xlNone
'    End If
    range1 = "G" + CStr(i1)
    If Sheet1.Cells(i1, 7).Value <> Sheet2.Cells(i2, 7).Value Then
        Sheet1.range(range1).Cells.Interior.ColorIndex = 3
    Else
        Sheet1.range(range1).Cells.Interior.ColorIndex = xlNone
    End If
'    range1 = "H" + CStr(i1)
'    If Sheet1.Cells(i1, 8).Value <> Sheet2.Cells(i2, 8).Value Then
'        Sheet1.range(range1).Cells.Interior.ColorIndex = 3
'    Else
'        Sheet1.range(range1).Cells.Interior.ColorIndex = xlNone
'    End If
'    range1 = "I" + CStr(i1)
'    If Sheet1.Cells(i1, 9).Value <> Sheet2.Cells(i2, 9).Value Then
'        Sheet1.range(range1).Cells.Interior.ColorIndex = 3
'    Else
'        Sheet1.range(range1).Cells.Interior.ColorIndex = xlNone
'    End If
'    range1 = "J" + CStr(i1)
'    If Sheet1.Cells(i1, 10).Value <> Sheet2.Cells(i2, 10).Value Then
'        Sheet1.range(range1).Cells.Interior.ColorIndex = 3
'    Else
'        Sheet1.range(range1).Cells.Interior.ColorIndex = xlNone
'    End If
'    range1 = "K" + CStr(i1)
'    If Sheet1.Cells(i1, 11).Value <> Sheet2.Cells(i2, 11).Value Then
'        Sheet1.range(range1).Cells.Interior.ColorIndex = 3
'    Else
'        Sheet1.range(range1).Cells.Interior.ColorIndex = xlNone
'    End If
'    range1 = "L" + CStr(i1)
'    If Sheet1.Cells(i1, 12).Value <> Sheet2.Cells(i2, 12).Value Then
'        Sheet1.range(range1).Cells.Interior.ColorIndex = 3
'    Else
'        Sheet1.range(range1).Cells.Interior.ColorIndex = xlNone
'    End If
    range1 = "M" + CStr(i1)
    If Sheet1.Cells(i1, 13).Value <> Sheet2.Cells(i2, 13).Value Then
        Sheet1.range(range1).Cells.Interior.ColorIndex = 3
    Else
        Sheet1.range(range1).Cells.Interior.ColorIndex = xlNone
    End If
    range1 = "N" + CStr(i1)
    If Sheet1.Cells(i1, 14).Value <> Sheet2.Cells(i2, 14).Value Then
        Sheet1.range(range1).Cells.Interior.ColorIndex = 3
    Else
        Sheet1.range(range1).Cells.Interior.ColorIndex = xlNone
    End If
'    range1 = "O" + CStr(i1)
'    If Sheet1.Cells(i1, 15).Value <> Sheet2.Cells(i2, 15).Value Then
'        Sheet1.range(range1).Cells.Interior.ColorIndex = 3
'    Else
'        Sheet1.range(range1).Cells.Interior.ColorIndex = xlNone
'    End If
'    range1 = "P" + CStr(i1)
'    If Sheet1.Cells(i1, 16).Value <> Sheet2.Cells(i2, 16).Value Then
'        Sheet1.range(range1).Cells.Interior.ColorIndex = 3
'    Else
'        Sheet1.range(range1).Cells.Interior.ColorIndex = xlNone
'    End If
'    range1 = "Q" + CStr(i1)
'    If Sheet1.Cells(i1, 17).Value <> Sheet2.Cells(i2, 17).Value Then
'        Sheet1.range(range1).Cells.Interior.ColorIndex = 3
'    Else
'        Sheet1.range(range1).Cells.Interior.ColorIndex = xlNone
'    End If
'    range1 = "R" + CStr(i1)
'    If Sheet1.Cells(i1, 18).Value <> Sheet2.Cells(i2, 18).Value Then
'        Sheet1.range(range1).Cells.Interior.ColorIndex = 3
'    Else
'        Sheet1.range(range1).Cells.Interior.ColorIndex = xlNone
'    End If
'    range1 = "S" + CStr(i1)
'    If Sheet1.Cells(i1, 19).Value <> Sheet2.Cells(i2, 19).Value Then
'        Sheet1.range(range1).Cells.Interior.ColorIndex = 3
'    Else
'        Sheet1.range(range1).Cells.Interior.ColorIndex = xlNone
'    End If
'    range1 = "T" + CStr(i1)
'    If Sheet1.Cells(i1, 20).Value <> Sheet2.Cells(i2, 20).Value Then
'        Sheet1.range(range1).Cells.Interior.ColorIndex = 3
'    Else
'        Sheet1.range(range1).Cells.Interior.ColorIndex = xlNone
'    End If
'    range1 = "U" + CStr(i1)
'    If Sheet1.Cells(i1, 21).Value <> Sheet2.Cells(i2, 21).Value Then
'        Sheet1.range(range1).Cells.Interior.ColorIndex = 3
'    Else
'        Sheet1.range(range1).Cells.Interior.ColorIndex = xlNone
'    End If
'    range1 = "V" + CStr(i1)
'    If Sheet1.Cells(i1, 22).Value <> Sheet2.Cells(i2, 22).Value Then
'        Sheet1.range(range1).Cells.Interior.ColorIndex = 3
'    Else
'        Sheet1.range(range1).Cells.Interior.ColorIndex = xlNone
'    End If
    Sheet1.Cells(i1, 23).Value = "<-"
End Sub



