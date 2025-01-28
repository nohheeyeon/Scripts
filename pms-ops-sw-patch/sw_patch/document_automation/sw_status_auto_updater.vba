Sub CompareAndCopy()
    Dim ws1 As Worksheet
    Dim ws4 As Worksheet
    Dim dValues As Object
    Dim cell As Range
    Dim foundValue As Variant

    ' 첫번째 시트와 네번째 시트 설정
    Set ws1 = ThisWorkbook.Sheets(1)
    Set ws4 = ThisWorkbook.Sheets(4)

    ' 첫번째 시트의 D열 값을 컬렉션에 저장
    Set dValues = CreateObject("Scripting.Dictionary")
    For Each cell In ws1.Range("D1:D" & ws1.Cells(ws1.Rows.Count, "D").End(xlUp).Row)
        If Not IsEmpty(cell.Value) Then
            dValues(cell.Value) = cell.Row
        End If
    Next cell

    ' 네번째 시트의 E열 값ㅇ르 순회하며 비교
    For Each cell In ws4.Range("E1:E" & ws4.Cells(ws4.Rows.Count, "E").End(xlUp).Row)
        If Not IsEmpty(cell.Value) Then
            If dValues.Exists(cell.Value) Then
                ' 동일한 값이 있을 때 네번째 시트의 D열 값을 첫번째 시트의 F열로 복사
                ws1.Cells(dValues(cell.Value), "F").Value = ws4.Cells(cell.Row, "D").Value
                ' 동일한 값이 있을 때 네번째 시트의 H열 값을 첫번째 시트의 G열로 복사
                ws1.Cells(dValues(cell.Value), "G").Value = ws4.Cells(cell.Row, "D").Value
                ' 동일한 값이 있을 때 네번째 시트의 I열 값을 첫번째 시트의 H열로 복사
                ws1.Cells(dValues(cell.Value), "H").Value = ws4.Cells(cell.Row, "I").Value
                Debug.Print "동일한 값 발견: " & cell.Value & " (첫번째 시트의 " & dValues(cell.Value) & "행, 네번째 시트의 " & cell.Row & "행)"
                Debug.Print "네번째 시트의 D열 값 '" & ws4.Cells(cell.Row, "D").Value & "'을 첫번째 시트의 F열로 복사"
                Debug.Print "네번째 시트의 H열 값 '" & ws4.Cells(cell.Row, "H").Value & "'을 첫번째 시트의 G열로 복사"
                Debug.Print "네번째 시트의 I열 값 '" & ws4.Cells(cell.Row, "I").Value & "'을 첫번째 시트의 H열로 복사"
            End If
        End If
    Next cell
End Sub