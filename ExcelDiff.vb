Sub DIFF()
    Dim i As Long
    For i = 5 To 69448
        findRow (i)
    Next i
    
End Sub

Function findRow(i2 As Long)
    Dim key As String
    Dim found As Boolean
    Dim i1 As Long
    Dim tmp As Long
    i1 = 5
    key = Sheet2.Cells(i2, 1)
    Do While i1 < 10 And found = False
        tmp = findKey(key, i1)
        If tmp <> 0 Then
            i1 = tmp
            found = checkRow(i1, i2)
        End If
        If found = False Then i1 = i1 + 1
    Loop
    If found = False Then Sheet2.Cells(i, 1).EntireRow.Interior.ColorIndex = 44
End Function
    
Function findKey(key2 As String, i As Long) As Long
    Dim key1 As String
    Do While i < 10
        key1 = Sheet1.Cells(i, 1).Value
        If key1 = key2 Then
            findKey = i
            Exit Do
        End If
        i = i + 1
    Loop
End Function

Function checkRow(i1 As Long, i2 As Long) As Boolean
    If (Sheet1.Cells(i1, 2).Value = Sheet2.Cells(i2, 2).Value And Sheet1.Cells(i1, 4).Value = Sheet2.Cells(i2, 4).Value And Sheet1.Cells(i1, 7).Value = Sheet2.Cells(i2, 7).Value And Sheet1.Cells(i1, 13).Value = Sheet2.Cells(i2, 13).Value And Sheet1.Cells(i1, 14).Value = Sheet2.Cells(i2, 14).Value) Then checkRow = True
    
End Function
