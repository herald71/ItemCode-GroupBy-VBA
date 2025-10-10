Attribute VB_Name = "Module4"
Sub 텍스트숫자를숫자로변환_단건()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim columnLetter As String
    
    ' 현재 활성 시트 설정
    Set ws = ActiveSheet
    
    ' 사용자로부터 열 문자를 입력받음
    columnLetter = InputBox("변환할 열의 문자를 입력하세요 (예: A, B, C, ...)", "열 선택")
    
    ' 입력받은 열의 데이터 범위 설정 (제목 행 제외)
    Set rng = ws.Range(columnLetter & "2:" & columnLetter & ws.Cells(ws.Rows.count, columnLetter).End(xlUp).Row)
    
    ' 범위 내의 각 셀에 대해 텍스트를 숫자로 변환
    For Each cell In rng
        ' 셀 값이 숫자로 변환 가능한지 확인 후 변환
        If IsNumeric(cell.Value) Then
            cell.Value = val(cell.Value)
        End If
    Next cell
    
    MsgBox "텍스트 숫자가 숫자로 변환되었습니다.", vbInformation
End Sub
