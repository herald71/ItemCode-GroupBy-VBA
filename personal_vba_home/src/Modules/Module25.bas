Attribute VB_Name = "Module25"
Sub 단어포함행삭제()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim colNum As Integer
    Dim searchWords As String
    Dim keywords() As String
    Dim i As Long, j As Integer
    Dim cellValue As String
    Dim deleteCount As Integer
    
    ' 현재 활성화된 시트를 대상으로 함
    Set ws = ActiveSheet
    
    ' 사용자에게 열 번호 입력받기
    colNum = Application.InputBox("검색할 열 번호를 입력하세요 (예: 2는 B열)", Type:=1)
    If colNum < 1 Then Exit Sub ' 입력이 잘못되면 종료
    
    ' 사용자에게 삭제할 단어 입력받기 (쉼표로 구분)
    searchWords = Application.InputBox("삭제할 단어들을 입력하세요 (쉼표로 구분)", Type:=2)
    If searchWords = "" Then Exit Sub ' 입력이 없으면 종료
    
    ' 단어 배열로 변환
    keywords = Split(searchWords, ",")
    
    ' 마지막 행 찾기
    lastRow = ws.Cells(ws.Rows.count, colNum).End(xlUp).Row
    
    ' 삭제된 행 개수 초기화
    deleteCount = 0

    ' 뒤에서 앞으로 삭제해야 문제 없음
    For i = lastRow To 1 Step -1
        cellValue = ws.Cells(i, colNum).Value
        
        ' 입력된 단어 중 하나라도 포함되면 행 삭제
        For j = LBound(keywords) To UBound(keywords)
            If InStr(1, cellValue, Trim(keywords(j)), vbTextCompare) > 0 Then
                ws.Rows(i).Delete
                deleteCount = deleteCount + 1 ' 삭제된 행 개수 증가
                Exit For ' 한 번 삭제되면 해당 행은 더 이상 검사할 필요 없음
            End If
        Next j
    Next i
    
    ' 삭제된 행 수 출력
    MsgBox deleteCount & "개의 행이 삭제되었습니다.", vbInformation, "삭제 완료"
End Sub

