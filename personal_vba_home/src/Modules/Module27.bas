Attribute VB_Name = "Module27"
Sub 쿠팡광고_날짜별_캠페인별_분석()
    Dim ws As Worksheet, wsNew As Worksheet
    Dim lastRow As Long
    Dim dict As Object
    Dim i As Long
    Dim key As Variant

    ' 원본 데이터 시트 설정 (ActiveWorkbook 사용)
    Set ws = ActiveWorkbook.Sheets("Sheet1") ' 기존 데이터 시트 변경 가능
    
    ' 마지막 행 찾기
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    
    ' 데이터 저장할 Dictionary 생성
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' 기존 데이터에서 날짜 + 캠페인명별 그룹화
    For i = 2 To lastRow
        Dim 날짜 As String, 캠페인명 As String
        Dim 노출수 As Double, 클릭수 As Double, 광고비 As Double
        Dim 주문수 As Double, 전환매출액 As Double
        
        날짜 = ws.Cells(i, 1).Value ' 날짜 (A열)
        캠페인명 = ws.Cells(i, 6).Value ' 캠페인명 (F열)
        노출수 = ws.Cells(i, 14).Value ' 노출수 (N열)
        클릭수 = ws.Cells(i, 15).Value ' 클릭수 (O열)
        광고비 = ws.Cells(i, 16).Value ' 광고비 (P열)
        주문수 = ws.Cells(i, 18).Value ' 총 주문수(1일) (R열)
        전환매출액 = ws.Cells(i, 24).Value ' 총 전환매출액(1일) (X열)
        
        ' Key: "날짜|캠페인명" 조합으로 설정
        Dim dictKey As String
        dictKey = 날짜 & "|" & 캠페인명
        
        ' 기존 데이터가 없으면 초기화
        If Not dict.exists(dictKey) Then
            dict.Add dictKey, Array(0, 0, 0, 0, 0) ' [노출수, 클릭수, 광고비, 주문수, 전환매출액]
        End If
        
        ' 기존 값 누적
        Dim values As Variant
        values = dict(dictKey)
        
        values(0) = values(0) + 노출수
        values(1) = values(1) + 클릭수
        values(2) = values(2) + 광고비
        values(3) = values(3) + 주문수
        values(4) = values(4) + 전환매출액
        
        dict(dictKey) = values
    Next i
    
    ' 기존 시트 삭제 후 새로운 시트 생성 (ActiveWorkbook 기준)
    On Error Resume Next
    Application.DisplayAlerts = False
    ActiveWorkbook.Sheets("날짜별 캠페인 분석").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    ' 새로운 시트 추가
    Set wsNew = ActiveWorkbook.Sheets.Add
    wsNew.Name = "날짜별 캠페인 분석"
    
    ' 헤더 작성
    Dim headers As Variant
    headers = Array("날짜", "캠페인명", "ROAS(%)", "CPC", "클릭률(%)", "전환율(%)", "주문수", "광고비(VAT 포함)", "광고매출", "전환당비용", "객단가")
    
    For i = 0 To UBound(headers)
        wsNew.Cells(1, i + 1).Value = headers(i)
        wsNew.Cells(1, i + 1).Font.Bold = True ' 헤더 글씨 굵게
    Next i
    
    ' 데이터 입력
    Dim rowIndex As Integer
    rowIndex = 2
    
    For Each key In dict.Keys
        values = dict(key)
        
        ' Key에서 날짜와 캠페인명 분리
        Dim keyParts() As String
        keyParts = Split(key, "|")
        
        Dim 노출합 As Double, 클릭합 As Double, 광고비합 As Double
        Dim 주문합 As Double, 전환매출합 As Double
        Dim VAT포함광고비 As Double, ROAS As Double
        Dim CPC As Double, 클릭률 As Double, 전환율 As Double
        Dim 전환당비용 As Double, 객단가 As Double
        
        노출합 = values(0)
        클릭합 = values(1)
        광고비합 = values(2)
        주문합 = values(3)
        전환매출합 = values(4)
        
        ' VAT 포함 광고비 계산
        VAT포함광고비 = 광고비합 * 1.1
        
        ' 지표 계산 (0으로 나누는 경우 방지 + 반올림)
        If VAT포함광고비 <> 0 Then
            ROAS = Round((전환매출합 / VAT포함광고비) * 100, 2)
            CPC = Round(VAT포함광고비 / IIf(클릭합 = 0, 1, 클릭합), 2)
        Else
            ROAS = 0
            CPC = 0
        End If
        
        If 클릭합 <> 0 Then
            클릭률 = Round((클릭합 / 노출합) * 100, 2)
            전환율 = Round((주문합 / 클릭합) * 100, 2)
        Else
            클릭률 = 0
            전환율 = 0
        End If
        
        If 주문합 <> 0 Then
            전환당비용 = Round(VAT포함광고비 / 주문합, 2)
            객단가 = Round(전환매출합 / 주문합, 2)
        Else
            전환당비용 = 0
            객단가 = 0
        End If
        
        ' 시트에 값 입력
        wsNew.Cells(rowIndex, 1).Value = keyParts(0) ' 날짜
        wsNew.Cells(rowIndex, 2).Value = keyParts(1) ' 캠페인명
        wsNew.Cells(rowIndex, 3).Value = ROAS
        wsNew.Cells(rowIndex, 4).Value = CPC
        wsNew.Cells(rowIndex, 5).Value = 클릭률
        wsNew.Cells(rowIndex, 6).Value = 전환율
        wsNew.Cells(rowIndex, 7).Value = 주문합
        wsNew.Cells(rowIndex, 8).Value = VAT포함광고비
        wsNew.Cells(rowIndex, 9).Value = 전환매출합
        wsNew.Cells(rowIndex, 10).Value = 전환당비용
        wsNew.Cells(rowIndex, 11).Value = 객단가
        
        rowIndex = rowIndex + 1
    Next key

    ' 서식 적용
    With wsNew.Columns("C:K")
        .NumberFormat = "#,##0.00" ' 숫자 포맷 적용 (소수점 2자리)
        .AutoFit ' 열 너비 자동 조정
    End With
    
    wsNew.Columns("A:B").AutoFit ' 날짜, 캠페인명 열 너비 자동 조정
    
    ' 완료 메시지
    MsgBox "날짜별 캠페인 분석이 완료되었습니다!", vbInformation, "완료"
End Sub


