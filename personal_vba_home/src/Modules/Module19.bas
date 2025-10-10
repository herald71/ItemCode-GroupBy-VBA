Attribute VB_Name = "Module19"
Public Sub 도매꾹_베스트상품정리()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim Confirm As VbMsgBoxResult
    Dim cell As Range
    Dim img As Picture
    Dim imageUrl As String
    Dim TotalSheets As Long
    Dim SheetCount As Long
    Dim Progress As Double
    Dim wb As Workbook
    
    ' 현재 활성화된 통합 문서 참조
    Set wb = ActiveWorkbook
    
    ' 실행 여부 확인
    Confirm = MsgBox("도매꾹 베스트 상품 편집 작업을 진행 할까요?", vbYesNo + vbQuestion, "작업 확인")
    If Confirm = vbNo Then Exit Sub
    
    ' 화면 업데이트 및 이벤트 비활성화
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayStatusBar = True
    
    ' 총 시트 수 확인
    TotalSheets = wb.Worksheets.count
    SheetCount = 0
    
    ' 현재 통합 문서의 모든 시트에서 작업 실행
    For Each ws In wb.Worksheets
        SheetCount = SheetCount + 1
        
        ' 마지막 행 찾기
        lastRow = ws.Cells(ws.Rows.count, "G").End(xlUp).Row
        
        ' 진행률 표시 업데이트
        Progress = (SheetCount / TotalSheets) * 100
        Application.StatusBar = "작업 진행률: " & Format(Progress, "0.00") & "% 완료 중..."
        
        ' 1. G열의 URL을 I열에 "바로가기" 하이퍼링크로 생성
        For Each cell In ws.Range("G2:G" & lastRow)
            If cell.Value Like "http*://*" Then
                ws.Cells(cell.Row, "I").Value = "바로가기"
                ws.Hyperlinks.Add Anchor:=ws.Cells(cell.Row, "I"), Address:=cell.Value, TextToDisplay:="바로가기"
            End If
        Next cell
        
        ' 2. H열의 이미지 URL을 참고하여 J열에 이미지 삽입
        For Each cell In ws.Range("H2:H" & lastRow)
            imageUrl = cell.Value
            If imageUrl <> "" Then
                On Error Resume Next
                Set img = ws.Pictures.Insert(imageUrl)
                If Not img Is Nothing Then
                    With img
                        .Top = ws.Cells(cell.Row, "J").Top
                        .Left = ws.Cells(cell.Row, "J").Left
                        ws.Cells(cell.Row, "J").RowHeight = 50
                        
                        ' 비율에 맞게 크기 조정
                        Dim origRatio As Double
                        origRatio = .Width / .Height
                        Dim cellRatio As Double
                        cellRatio = ws.Cells(cell.Row, "J").Width / ws.Cells(cell.Row, "J").Height
                        
                        If origRatio > cellRatio Then
                            .Width = ws.Cells(cell.Row, "J").Width
                            .Height = .Width / origRatio
                        Else
                            .Height = ws.Cells(cell.Row, "J").Height
                            .Width = .Height * origRatio
                        End If
                        
                        ' 이미지 중앙 배치
                        .Top = ws.Cells(cell.Row, "J").Top + (ws.Cells(cell.Row, "J").Height - .Height) / 2
                        .Left = ws.Cells(cell.Row, "J").Left + (ws.Cells(cell.Row, "J").Width - .Width) / 2
                    End With
                End If
                On Error GoTo 0
            End If
        Next cell
        
        ' 3. G열과 H열 숨김 처리
        ws.Columns("G").EntireColumn.Hidden = True
        ws.Columns("H").EntireColumn.Hidden = True
        
        ' 4. I1, J1에 텍스트 추가
        ws.Range("I1").Value = "바로가기"
        ws.Range("J1").Value = "이미지"
        
        ' 5. 테두리 정리 및 정렬
        With ws.Range("A1:J" & lastRow)
            .Borders.LineStyle = xlContinuous
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        
        ' 6. D열의 본문 내용 왼쪽 정렬 및 D1 중앙 정렬
        ws.Range("D2:D" & lastRow).HorizontalAlignment = xlLeft
        ws.Range("D1").HorizontalAlignment = xlCenter
    Next ws
    
    ' 작업 종료
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "도매꾹 베스트 상품 편집 작업이 완료되었습니다.", vbInformation, "작업 완료"
End Sub
