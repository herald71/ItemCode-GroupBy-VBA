Attribute VB_Name = "Module29"

Sub 쿠팡광고전환상품분석()
    Dim ws As Worksheet, analysisWs As Worksheet
    Dim lastRow As Long, analysisRow As Long
    Dim dict As Object
    Dim key As String
    Dim rng As Range, cell As Range
    Dim wb As Workbook
    
    ' 현재 활성화된 파일을 기준으로 설정 (ThisWorkbook → ActiveWorkbook 변경)
    Set wb = ActiveWorkbook
    
    ' Sheet1이 아닌 사용자가 선택한 시트를 기준으로 실행하도록 변경
    On Error Resume Next
    Set ws = wb.ActiveSheet
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "활성 시트가 없습니다. 파일을 연 후 다시 시도하세요.", vbExclamation
        Exit Sub
    End If
    
    ' Dictionary 생성 (키: 상품명 + 옵션ID, 값: 키워드, 광고매출, 주문수)
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' 데이터 범위 설정
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    Set rng = ws.Range("J2:J" & lastRow) ' "광고전환매출발생 상품명"이 J열에 있다고 가정
    
    ' 데이터 수집 및 집계
    For Each cell In rng
        key = ws.Cells(cell.Row, 10).Value & "|" & ws.Cells(cell.Row, 11).Value ' 상품명(J열) + 옵션ID(K열)
        
        If Not dict.exists(key) Then
            dict.Add key, Array("", 0#, 0#) ' 키워드 초기화 (문자열), 광고매출 (Double), 주문수 (Double)
        End If
        
        Dim dataArr As Variant
        dataArr = dict(key)
        
        ' 키워드 (첫 번째 값만 저장)
        If dataArr(0) = "" Then
            dataArr(0) = ws.Cells(cell.Row, 13).Value ' 키워드 (M열)
        End If
        
        ' 광고매출 집계 (X열)
        dataArr(1) = dataArr(1) + CDbl(Nz(ws.Cells(cell.Row, 24).Value, 0)) ' 광고매출 (X열)
        
        ' 주문수 집계 (R열)
        dataArr(2) = dataArr(2) + CDbl(Nz(ws.Cells(cell.Row, 18).Value, 0)) ' 주문수 (R열)
        
        dict(key) = dataArr
    Next cell
    
    ' 기존 분석 시트 삭제 후 새로 생성
    On Error Resume Next
    Application.DisplayAlerts = False
    If Not wb.Sheets("광고전환 분석") Is Nothing Then wb.Sheets("광고전환 분석").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    ' 새 시트 생성
    Set analysisWs = wb.Sheets.Add
    analysisWs.Name = "광고전환 분석"
    
    ' 헤더 추가 및 서식 적용
    With analysisWs.Range("A1:E1")
        .Value = Array("광고전환매출발생 상품명", "광고전환매출발생 옵션ID", "키워드", "광고매출", "주문수")
        .Font.Bold = True
        .Borders.LineStyle = xlContinuous
        .Interior.Color = RGB(220, 220, 220) ' 연한 회색 배경
    End With
    
    ' 데이터 출력
    analysisRow = 2
    Dim item As Variant
    For Each item In dict.Keys
        Dim splitKey As Variant
        Dim productName As String, optionID As String, keyword As String
        Dim revenue As Double, orders As Double
        
        splitKey = Split(item, "|")
        productName = splitKey(0)
        optionID = splitKey(1)
        keyword = dict(item)(0)
        revenue = dict(item)(1)
        orders = dict(item)(2)
        
        analysisWs.Cells(analysisRow, 1).Value = productName
        analysisWs.Cells(analysisRow, 2).Value = optionID
        analysisWs.Cells(analysisRow, 3).Value = keyword
        analysisWs.Cells(analysisRow, 4).Value = revenue
        analysisWs.Cells(analysisRow, 5).Value = orders
        
        analysisRow = analysisRow + 1
    Next item
    
    ' 테두리 및 열 너비 조정
    With analysisWs.Range("A1:E" & analysisRow - 1)
        .Borders.LineStyle = xlContinuous
        .Columns.AutoFit
    End With
    
    ' 정렬 (광고매출 기준 내림차순)
    analysisWs.Range("A1:E" & analysisRow - 1).Sort _
        Key1:=analysisWs.Range("D1"), Order1:=xlDescending, Header:=xlYes
    
    MsgBox "광고전환 분석 완료!", vbInformation
End Sub

Function Nz(Value As Variant, Default As Double) As Double
    If IsNumeric(Value) Then
        Nz = CDbl(Value)
    Else
        Nz = Default
    End If
End Function




