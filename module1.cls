Sub 确认提交()
Dim total
Dim eatTime
Dim diffTime1, diffTime2
Dim k1, k2, k3

k1 = 1
k2 = 1
k3 = 1
total = 30000
diffTime1 = "10:00"
diffTime2 = "16:00"

    For i = 2 To total
        eatTime = Format(Sheet1.Cells(i, 5).Value, "HH:MM")
        If eatTime <> "" Then
            If DateDiff("s", eatTime, diffTime1) >= 0 Then
                k1 = k1 + 1
                Sheet1.Range(Sheet1.Cells(i, 1), Sheet1.Cells(i, 6)).Copy Sheet2.Range(Sheet2.Cells(k1, 1), Sheet2.Cells(k1, 6))
                
            ElseIf DateDiff("s", eatTime, diffTime1) < 0 And DateDiff("s", eatTime, diffTime2) >= 0 Then
                k2 = k2 + 1
                Sheet1.Range(Sheet1.Cells(i, 1), Sheet1.Cells(i, 6)).Copy Sheet3.Range(Sheet3.Cells(k2, 1), Sheet3.Cells(k2, 6))
                
            ElseIf DateDiff("s", eatTime, diffTime2) < 0 Then
                k3 = k3 + 1
                Sheet1.Range(Sheet1.Cells(i, 1), Sheet1.Cells(i, 6)).Copy Sheet4.Range(Sheet4.Cells(k3, 1), Sheet4.Cells(k3, 6))
                
            End If
        End If
    Next
    
End Sub

