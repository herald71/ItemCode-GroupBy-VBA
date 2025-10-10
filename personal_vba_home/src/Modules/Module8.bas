Attribute VB_Name = "Module8"
Sub 하이퍼링크만들기_열전체()
    Dim sourceCol As String
    Dim targetCol As String
    Dim lastRow As Long
    Dim ws As Worksheet
    Dim i As Long
    Dim hyperlinkAddress As String
    
    ' 활성 시트 설정
    Set ws = ActiveSheet
    
    ' 하이퍼링크 주소가 있는 열 번호 입력 받기
    sourceCol = InputBox("하이퍼링크 주소가 있는 열 번호를 입력하세요 (예: A, B, C 등):", "하이퍼링크 소스 열 번호 입력")
    If sourceCol = "" Then
        MsgBox "소스 열 입력이 취소되었습니다."
        Exit Sub
    End If
    
    ' "바로가기"가 생성될 열 번호 입력 받기
    targetCol = InputBox("하이퍼링크를 생성할 열 번호를 입력하세요 (예: A, B, C 등):", "하이퍼링크 타겟 열 번호 입력")
    If targetCol = "" Then
        MsgBox "타겟 열 입력이 취소되었습니다."
        Exit Sub
    End If
    
    ' 입력된 소스 열에서 마지막 행 찾기
    lastRow = ws.Cells(ws.Rows.count, sourceCol).End(xlUp).Row
    
    ' 소스 열의 하이퍼링크 주소를 읽고 타겟 열에 "바로가기" 생성
    For i = 1 To lastRow
        hyperlinkAddress = ws.Cells(i, sourceCol).Value
        If hyperlinkAddress Like "http*://*" Then
            ' 타겟 열의 같은 행에 "바로가기" 텍스트로 하이퍼링크 추가
            With ws.Cells(i, targetCol)
                .Value = "바로가기"
                ws.Hyperlinks.Add Anchor:=.Cells, Address:=hyperlinkAddress, TextToDisplay:="바로가기"
            End With
        End If
    Next i
    
    MsgBox "하이퍼링크 생성 작업이 완료되었습니다."
End Sub


