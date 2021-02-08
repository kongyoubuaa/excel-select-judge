Sub 选择判题()
    Rem Sheet1 is answer sheet
    Rem Sheet2 is correct answer
    
    Rem 变量声明
    Dim myRow As Integer
    Dim myCol As Integer
    Dim totalRow As Integer
    
    Rem 判题
    totalRow = Sheet1.UsedRange.Rows.Count
    For myRow = 2 To totalRow
        For myCol = 2 To 21
            If Sheet1.Cells(myRow, myCol).Value <> 1 And Sheet1.Cells(myRow, myCol).Value <> 0 Then
                If Sheet1.Cells(myRow, myCol).Value = Sheet3.Cells(2, myCol).Value Then
                    Sheet1.Cells(myRow, myCol).Value = 1
                Else
                    Sheet1.Cells(myRow, myCol).Value = 0
                End If
            End If
        Next myCol
    Next myRow

End Sub
