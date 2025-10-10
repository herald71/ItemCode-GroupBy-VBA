Attribute VB_Name = "URL이하삭세"

Sub URL에서물음표이하_텍스트삭제()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim questionMarkPos As Long
    Dim targetColumn As String
    Dim lastRow As Long
    
    ' 현재 활성화된 워크시트 설정
    Set ws = ActiveSheet
    
    ' 사용자로부터 열 입력받기 (예: "A"와 같이 입력)
    targetColumn = InputBox("URL이 포함된 열을 입력하세요 (예: A)")
    
    ' 입력이 유효하지 않으면 종료
    If targetColumn = "" Then
        MsgBox "유효한 열을 입력하세요."
        Exit Sub
    End If
    
    ' 마지막 행 번호 찾기 (해당 열에서)
    lastRow = ws.Cells(ws.Rows.count, targetColumn).End(xlUp).Row
    
    ' 지정한 열의 모든 셀 범위 설정
    Set rng = ws.Range(targetColumn & "1:" & targetColumn & lastRow)
    
    ' 각 셀을 순회하며 '?' 이후 텍스트 삭제
    For Each cell In rng
        If InStr(cell.Value, "?") > 0 Then
            questionMarkPos = InStr(cell.Value, "?")
            cell.Value = Left(cell.Value, questionMarkPos - 1)
        End If
    Next cell
End Sub


