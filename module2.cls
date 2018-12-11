Private Sub CommandButton1_Click()
Dim total
Dim dateSelected
Dim eatTime1, eatTime2, eatTime3
Dim k1, k2, k3, r
Dim valueSheet5

k1 = 0
k2 = 0
k3 = 0
total = 30000
dateSelected = Format(DateSerial(Left(TextBox1.Value, 4), Mid(TextBox1.Value, 5, 2), Right(TextBox1.Value, 2)), "YYYY-MM-DD")
'dateSelected = DateSerial(Left(TextBox1.Value, 4), Mid(TextBox1.Value, 5, 2), Right(TextBox1.Value, 2))

    For i = 2 To total
        eatTime1 = Format(Sheet2.Cells(i, 5).Value, "YYYY-MM-DD")
        If dateSelected <> "" Then
            If eatTime1 <> "" Then
                If eatTime1 = dateSelected Then
                 k1 = k1 + 1
                End If
            End If
        End If
    Next
    
    For j = 2 To total
        eatTime2 = Format(Sheet3.Cells(j, 5).Value, "YYYY-MM-DD")
        If dateSelected <> "" Then
            If eatTime2 <> "" Then
                If eatTime2 = dateSelected Then
                 k2 = k2 + 1
                End If
            End If
        End If
    Next
    
    For l = 2 To total
        eatTime3 = Format(Sheet4.Cells(l, 5).Value, "YYYY-MM-DD")
        If dateSelected <> "" Then
            If eatTime3 <> "" Then
                If eatTime3 = dateSelected Then
                 k3 = k3 + 1
                End If
            End If
        End If
    Next
    
    r = 2
    For m = 2 To total
        valueSheet5 = Sheet5.Cells(m, 2).Value
        If valueSheet5 <> "" Then
            r = r + 1
        Else
            Exit For
        End If
    Next
    
    Sheet5.Cells(r, 1).Value = dateSelected
    Sheet5.Cells(r + 1, 1).Value = dateSelected
    Sheet5.Cells(r + 2, 1).Value = dateSelected
    Sheet5.Cells(r, 2).Value = "早"
    Sheet5.Cells(r + 1, 2).Value = "中"
    Sheet5.Cells(r + 2, 2).Value = "晚"
    Sheet5.Cells(r, 3).Value = k1
    Sheet5.Cells(r + 1, 3).Value = k2
    Sheet5.Cells(r + 2, 3).Value = k3
    
    If Sheet5.Cells(r, 1).Value = Sheet5.Cells(r + 1, 1).Value And Sheet5.Cells(r, 1).Value = Sheet5.Cells(r + 2, 1).Value Then
        Application.DisplayAlerts = False
                Range(Cells(r, 1), Cells(r + 1, 1)).Merge
                Range(Cells(r + 1, 1), Cells(r + 2, 1)).Merge
        Application.DisplayAlerts = True
    End If

End Sub

'Private Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
'End Sub

Private Sub TextBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii < Asc(0) Or KeyAscii > Asc(9) Then
        MsgBox "请输入正确日期!"
    End If
End Sub

