Attribute VB_Name = "Module23"
Sub 쿠팡광고_키워드별_분석()
    Dim ws As Worksheet, analysisWs As Worksheet
    Dim wb As Workbook
    Dim lastRow As Long, analysisRow As Long
    Dim dict As Object
    Dim key As String
    Dim rng As Range, cell As Range
    
    ' 작업할 워크북 선택
    On Error Resume Next
    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        MsgBox "활성화된 워크북이 없습니다. 먼저 분석할 파일을 열어 주세요.", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    
    ' "Sheet1"을 분석 대상으로 자동 설정
    On Error Resume Next
    Set ws = wb.Sheets("Sheet1")
    On Error GoTo 0
    
    ' 시트 존재 여부 확인
    If ws Is Nothing Then
        MsgBox """Sheet1""이 존재하지 않습니다. 올바른 파일을 열어 주세요.", vbCritical
        Exit Sub
    End If
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' 데이터 범위 설정
    lastRow = ws.Cells(ws.Rows.count, "M").End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "키워드 데이터가 없습니다.", vbExclamation
        Exit Sub
    End If
    Set rng = ws.Range("M2:M" & lastRow) ' "키워드"가 M열에 있다고 가정
    
    ' 데이터 수집 및 집계
    Dim dataArr As Variant
    For Each cell In rng
        key = Trim(cell.Value) ' 공백 제거한 키워드
        
        If key <> "" Then ' 빈 값 제외
            If Not dict.exists(key) Then
                dict.Add key, Array(0, 0, 0, 0, 0) ' 노출수, 클릭수, 광고비, 주문수, 전환매출액 초기화
            End If
            
            dataArr = dict(key)
            
            ' 데이터 값 집계 (N열: 노출수, O열: 클릭수, P열: 광고비, R열: 주문수, X열: 전환매출액)
            dataArr(0) = dataArr(0) + CDbl(ws.Cells(cell.Row, 14).Value) ' 노출수 (N열)
            dataArr(1) = dataArr(1) + CDbl(ws.Cells(cell.Row, 15).Value) ' 클릭수 (O열)
            dataArr(2) = dataArr(2) + CDbl(ws.Cells(cell.Row, 16).Value) ' 광고비 (P열)
            dataArr(3) = dataArr(3) + CDbl(ws.Cells(cell.Row, 18).Value) ' 주문수 (R열)
            dataArr(4) = dataArr(4) + CDbl(ws.Cells(cell.Row, 24).Value) ' 전환매출액 (X열)
            
            dict(key) = dataArr
        End If
    Next cell
    
    ' 기존 "키워드 분석" 시트 삭제 후 새로 생성
    On Error Resume Next
    Application.DisplayAlerts = False
    If Not wb.Sheets("키워드 분석") Is Nothing Then
        wb.Sheets("키워드 분석").Delete
    End If
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Set analysisWs = wb.Sheets.Add
    analysisWs.Name = "키워드 분석"
    
    ' 헤더 추가 및 서식 적용
    With analysisWs.Range("A1:J1")
        .Value = Array("키워드", "노출수", "클릭수", "클릭률(%)", "주문수", "전환율(%)", "CPC", "광고비", "광고매출", "ROAS(%)")
        .Font.Bold = True
        .Borders.LineStyle = xlContinuous
        .Interior.Color = RGB(220, 220, 220) ' 연한 회색 배경
    End With
    
    ' 데이터 출력
    analysisRow = 2
    Dim item As Variant
    For Each item In dict.Keys
        Dim impressions As Double, clicks As Double, cost As Double, orders As Double, revenue As Double
        
        dataArr = dict(item)
        
        impressions = dataArr(0)
        clicks = dataArr(1)
        cost = dataArr(2) * 1.1 ' VAT 포함 광고비
        orders = dataArr(3)
        revenue = dataArr(4)
        
        analysisWs.Cells(analysisRow, 1).Value = item
        analysisWs.Cells(analysisRow, 2).Value = impressions
        analysisWs.Cells(analysisRow, 3).Value = clicks
        
        ' 클릭률 계산 (0으로 나누기 방지)
        If impressions > 0 Then
            analysisWs.Cells(analysisRow, 4).Value = (clicks / impressions) * 100 ' 클릭률(%)
        Else
            analysisWs.Cells(analysisRow, 4).Value = 0
        End If

        analysisWs.Cells(analysisRow, 5).Value = orders
        
        ' 전환율 계산
        If clicks > 0 Then
            analysisWs.Cells(analysisRow, 6).Value = (orders / clicks) * 100 ' 전환율(%)
        Else
            analysisWs.Cells(analysisRow, 6).Value = 0
        End If
        
        ' CPC 계산
        If clicks > 0 Then
            analysisWs.Cells(analysisRow, 7).Value = cost / clicks ' CPC
        Else
            analysisWs.Cells(analysisRow, 7).Value = 0
        End If

        analysisWs.Cells(analysisRow, 8).Value = cost ' 광고비(VAT 포함)
        analysisWs.Cells(analysisRow, 9).Value = revenue ' 광고매출
        
        ' ROAS 계산
        If cost > 0 Then
            analysisWs.Cells(analysisRow, 10).Value = (revenue / cost) * 100 ' ROAS(%)
        Else
            analysisWs.Cells(analysisRow, 10).Value = 0
        End If
        
        analysisRow = analysisRow + 1
    Next item
    
    ' 테두리 및 열 너비 조정
    With analysisWs.Range("A1:J" & analysisRow - 1)
        .Borders.LineStyle = xlContinuous
        .Columns.AutoFit
    End With
    
    ' 정렬 (노출수 기준 내림차순)
    analysisWs.Range("A1:J" & analysisRow - 1).Sort Key1:=analysisWs.Range("B1"), Order1:=xlDescending, Header:=xlYes
    
    MsgBox "키워드 분석 완료!", vbInformation
End Sub


