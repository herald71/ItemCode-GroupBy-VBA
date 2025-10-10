Attribute VB_Name = "Module24"
Sub 쿠팡광고집행상품분석()
    Dim ws As Worksheet, analysisWs As Worksheet
    Dim lastRow As Long, analysisRow As Long
    Dim dict As Object
    Dim key As String
    Dim rng As Range, cell As Range
    Dim wb As Workbook

    ' 현재 활성화된 워크북을 참조
    Set wb = ActiveWorkbook
    Set ws = wb.Sheets("Sheet1") ' 원본 데이터가 들어 있는 시트

    Set dict = CreateObject("Scripting.Dictionary")

    ' 데이터 범위 설정
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    Set rng = ws.Range("H2:H" & lastRow) ' "광고집행 상품명"이 H열에 있다고 가정

    ' 데이터 수집 및 집계
    For Each cell In rng
        key = ws.Cells(cell.Row, 8).Value & "|" & ws.Cells(cell.Row, 9).Value

        If Not dict.exists(key) Then
            dict.Add key, Array(0#, 0#, 0#, 0#, 0#) ' 주문수, 광고비, 광고매출, 노출수, 클릭수 초기화
        End If

        Dim dataArr As Variant
        dataArr = dict(key)

        ' 데이터 값 집계 (Nz 함수 적용)
        dataArr(0) = dataArr(0) + Nz(ws.Cells(cell.Row, 18).Value, 0) ' 주문수 (R열)
        dataArr(1) = dataArr(1) + Nz(ws.Cells(cell.Row, 16).Value, 0) ' 광고비 (P열)
        dataArr(2) = dataArr(2) + Nz(ws.Cells(cell.Row, 24).Value, 0) ' 광고매출 (X열)
        dataArr(3) = dataArr(3) + Nz(ws.Cells(cell.Row, 14).Value, 0) ' 노출수 (N열)
        dataArr(4) = dataArr(4) + Nz(ws.Cells(cell.Row, 15).Value, 0) ' 클릭수 (O열)

        dict(key) = dataArr
    Next cell

    ' 기존 분석 시트 삭제 후 새로 생성
    On Error Resume Next
    Application.DisplayAlerts = False
    wb.Sheets("광고집행 상품분석").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set analysisWs = wb.Sheets.Add
    analysisWs.Name = "광고집행 상품분석"

    ' 헤더 추가
    With analysisWs.Range("A1:J1")
        .Value = Array("광고집행 상품명", "광고집행 옵션ID", "주문수", "광고비", "광고매출", "ROAS(%)", "노출수", "클릭수", "클릭률(%)", "전환율(%)")
        .Font.Bold = True
        .Borders.LineStyle = xlContinuous
        .Interior.Color = RGB(220, 220, 220)
    End With

    ' 데이터 출력
    analysisRow = 2
    Dim item As Variant
    For Each item In dict.Keys
        Dim splitKey As Variant
        Dim productName As String, optionID As String
        Dim orders As Double, cost As Double, revenue As Double, impressions As Double, clicks As Double

        splitKey = Split(item, "|")
        productName = splitKey(0)
        optionID = splitKey(1)

        orders = dict(item)(0)
        cost = dict(item)(1) * 1.1 ' VAT 포함 광고비
        revenue = dict(item)(2)
        impressions = dict(item)(3)
        clicks = dict(item)(4)

        analysisWs.Cells(analysisRow, 1).Value = productName
        analysisWs.Cells(analysisRow, 2).Value = optionID
        analysisWs.Cells(analysisRow, 3).Value = orders
        analysisWs.Cells(analysisRow, 4).Value = cost
        analysisWs.Cells(analysisRow, 5).Value = revenue

        ' ROAS 계산
        If cost > 0 Then
            analysisWs.Cells(analysisRow, 6).Value = Round((revenue / cost) * 100, 2)
        Else
            analysisWs.Cells(analysisRow, 6).Value = 0
        End If

        analysisWs.Cells(analysisRow, 7).Value = impressions
        analysisWs.Cells(analysisRow, 8).Value = clicks

        ' 클릭률 계산
        If impressions > 0 Then
            analysisWs.Cells(analysisRow, 9).Value = Round((clicks / impressions) * 100, 2)
        Else
            analysisWs.Cells(analysisRow, 9).Value = 0
        End If

        ' 전환율 계산
        If clicks > 0 Then
            analysisWs.Cells(analysisRow, 10).Value = Round((orders / clicks) * 100, 2)
        Else
            analysisWs.Cells(analysisRow, 10).Value = 0
        End If

        analysisRow = analysisRow + 1
    Next item

    MsgBox "광고집행 상품분석 완료!", vbInformation
End Sub


Function Nz(Value, Default As Double) As Double
    If Not IsNumeric(Value) Or IsError(Value) Or IsEmpty(Value) Then
        Nz = Default
    Else
        Nz = CDbl(Value)
    End If
End Function


