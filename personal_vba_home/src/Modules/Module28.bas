Attribute VB_Name = "Module28"
Sub 쿠팡광고노출지면분석()
    Dim ws As Worksheet, analysisWs As Worksheet
    Dim lastRow As Long, analysisRow As Long
    Dim dict As Object
    Dim key As String
    Dim rng As Range, cell As Range

    ' 현재 시트 설정 (활성화된 워크북의 Sheet1)
    Set ws = ActiveWorkbook.Sheets("Sheet1")
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' 데이터 범위 설정
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    Set rng = ws.Range("L2:L" & lastRow) ' "광고 노출 지면"이 L열에 있다고 가정
    
    ' 데이터 수집 및 집계
    For Each cell In rng
        key = ws.Cells(cell.Row, 6).Value & "|" & cell.Value ' 캠페인명(F열)과 광고 노출 지면(L열) 조합
        
        If Not dict.exists(key) Then
            dict.Add key, Array(0, 0, 0, 0, 0, 0) ' 노출수, 클릭수, 광고비, 주문수, 전환매출액 초기화
        End If
        
        Dim dataArr As Variant
        dataArr = dict(key)
        
        ' 데이터 값 집계 (N열: 노출수, O열: 클릭수, P열: 광고비, R열: 주문수, X열: 전환매출액)
        dataArr(0) = dataArr(0) + val(ws.Cells(cell.Row, 14).Value) ' 노출수
        dataArr(1) = dataArr(1) + val(ws.Cells(cell.Row, 15).Value) ' 클릭수
        dataArr(2) = dataArr(2) + val(ws.Cells(cell.Row, 16).Value) ' 광고비
        dataArr(3) = dataArr(3) + val(ws.Cells(cell.Row, 18).Value) ' 주문수
        dataArr(4) = dataArr(4) + val(ws.Cells(cell.Row, 24).Value) ' 전환매출액
        
        dict(key) = dataArr
    Next cell
    
    ' 기존 시트 삭제 후 새로 생성
    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets("광고노출지면 분석").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Set analysisWs = ActiveWorkbook.Sheets.Add
    analysisWs.Name = "광고노출지면 분석"
    
    ' 헤더 추가 및 서식 적용
    With analysisWs.Range("A1:N1")
        .Value = Array("캠페인명", "광고 노출 지면", "노출수", "클릭수", "주문수", "클릭률(%)", "전환율(%)", "CPM", "CPC", "광고비", "광고매출", "ROAS(%)", "전환당비용", "객단가")
        .Font.Bold = True
        .Borders.LineStyle = xlContinuous
        .Interior.Color = RGB(220, 220, 220) ' 연한 회색 배경
    End With
    
    ' 데이터 출력
    analysisRow = 2
    Dim item As Variant
    For Each item In dict.Keys
        Dim splitKey As Variant
        Dim campaign As String, exposure As String
        Dim impressions As Double, clicks As Double, cost As Double, orders As Double, revenue As Double
        
        splitKey = Split(item, "|")
        campaign = splitKey(0)
        exposure = splitKey(1)
        
        impressions = dict(item)(0)
        clicks = dict(item)(1)
        cost = dict(item)(2) * 1.1 ' VAT 포함 광고비
        orders = dict(item)(3)
        revenue = dict(item)(4)
        
        analysisWs.Cells(analysisRow, 1).Value = campaign
        analysisWs.Cells(analysisRow, 2).Value = exposure
        analysisWs.Cells(analysisRow, 3).Value = impressions
        analysisWs.Cells(analysisRow, 4).Value = clicks
        analysisWs.Cells(analysisRow, 5).Value = orders
        analysisWs.Cells(analysisRow, 6).Value = IIf(impressions > 0, (clicks / impressions) * 100, 0) ' 클릭률
        analysisWs.Cells(analysisRow, 7).Value = IIf(clicks > 0, (orders / clicks) * 100, 0) ' 전환율
        analysisWs.Cells(analysisRow, 8).Value = IIf(impressions > 0, (cost / impressions) * 1000, 0) ' CPM
        analysisWs.Cells(analysisRow, 9).Value = IIf(clicks > 0, cost / clicks, 0) ' CPC
        analysisWs.Cells(analysisRow, 10).Value = cost ' 광고비(VAT 포함)
        analysisWs.Cells(analysisRow, 11).Value = revenue ' 광고매출
        analysisWs.Cells(analysisRow, 12).Value = IIf(cost > 0, (revenue / cost) * 100, 0) ' ROAS
        analysisWs.Cells(analysisRow, 13).Value = IIf(orders > 0, cost / orders, 0) ' 전환당비용
        analysisWs.Cells(analysisRow, 14).Value = IIf(orders > 0, revenue / orders, 0) ' 객단가
        
        analysisRow = analysisRow + 1
    Next item
    
    ' 테두리 및 열 너비 조정
    With analysisWs.Range("A1:N" & analysisRow - 1)
        .Borders.LineStyle = xlContinuous
        .Columns.AutoFit
    End With
    
    ' 정렬 (노출수 기준 내림차순)
    analysisWs.Range("A1:N" & analysisRow - 1).Sort Key1:=analysisWs.Range("C1"), Order1:=xlDescending, Header:=xlYes
    
    MsgBox "광고 노출 지면 분석 완료!", vbInformation
End Sub


