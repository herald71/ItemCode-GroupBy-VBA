Attribute VB_Name = "Module5"
Sub ConvertTextPercentToNumbers()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim columnLetters As String
    Dim arrColumns As Variant
    Dim column As Variant
    Dim lastRow As Long
    
    ' 현재 활성 시트 설정
    Set ws = ActiveSheet
    
    ' 사용자로부터 쉼표로 구분된 열 문자를 입력받음
    columnLetters = InputBox("변환할 열의 문자를 쉼표로 구분하여 입력하세요 (예: A,B,C)", "열 선택")
    
    ' 입력받은 열 문자를 쉼표로 분리하여 배열에 저장
    arrColumns = Split(columnLetters, ",")
    
    ' 배열의 각 열에 대한 작업 수행
    For Each column In arrColumns
        column = Trim(column) ' 앞뒤 공백 제거
        If ws.Columns(column).column <= ws.Columns.count Then
            lastRow = ws.Cells(ws.Rows.count, column).End(xlUp).Row
            
            ' 마지막 행이 제목 행(1행)보다 클 경우에만 작업 수행
            If lastRow > 1 Then
                Set rng = ws.Range(column & "2:" & column & lastRow)
                
                ' 범위 내의 각 셀에 대한 작업 수행
                For Each cell In rng
                    If InStr(cell.Value, "%") > 0 Then
                        ' 텍스트 형식의 백분율 값을 실제 숫자로 변환
                        cell.Value = CDbl(Replace(cell.Value, "%", "")) / 100
                        ' 셀 형식을 백분율로 설정
                        cell.NumberFormat = "0.00%"
                    ElseIf IsNumeric(cell.Value) Then
                        ' 셀 값이 숫자로 변환 가능한 경우 변환
                        cell.Value = val(cell.Value)
                    End If
                Next cell
            End If
        Else
            MsgBox column & "는 유효하지 않은 열 문자입니다.", vbCritical
            Exit Sub ' 유효하지 않은 입력에 대해 코드 실행 종료
        End If
    Next column
    
    MsgBox "선택한 열의 텍스트 형식의 백분율 및 숫자가 숫자 형식으로 변환되었습니다.", vbInformation
End Sub

