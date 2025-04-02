```
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim cell As Range
    If Not Intersect(Target, Me.Range("C9:C62")) Is Nothing Then
        Application.EnableEvents = False
        For Each cell In Target
            If cell.Value = "저녁식대" Or cell.Value = "점심식대" Then
                cell.Offset(0, 4).Select
            ElseIf cell.Value = "택배비용" Or cell.Value = "퀵배송" Or cell.Value = "음료대" Or cell.Value = "소모품비" Then
                cell.Offset(0, 8).Select
            ElseIf cell.Value = "주차비" Or cell.Value = "이동비" Then
                cell.Offset(0, 7).Select
            ElseIf cell.Value = "주유비" Then
                cell.Offset(0, 6).Select
            End If
        Next cell
        Application.EnableEvents = True
    End If
    
    ' G9:M62 범위에 대한 처리
    If Not Intersect(Target, Me.Range("G9:M62")) Is Nothing Then
        Application.EnableEvents = False
        For Each cell In Target
            If cell.Value <> "" Then
                Dim nextRow As Long
                Dim nextCol As Long
                nextRow = cell.Row + 1
                nextCol = 2 ' B열로 시작

                ' 다음 위치에 값이 있는지 확인
                Do While nextRow <= 62
                    If Cells(nextRow, nextCol).Value = "" Then
                        Cells(nextRow, nextCol).Select
                        Exit Do
                    Else
                        ' 값이 있으면 열을 +1, 또는 행을 +1
                        If nextCol = 2 Then
                            nextCol = 3 ' C열로 이동
                        Else
                            nextCol = 2 ' B열로 돌아가서 다음 행으로
                            nextRow = nextRow + 1 ' 다음 행으로 이동
                        End If
                    End If
                Loop
            End If
        Next cell
        Application.EnableEvents = True
    End If
    
End Sub
```
