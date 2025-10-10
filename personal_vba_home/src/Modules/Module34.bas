Attribute VB_Name = "Module34"
Sub RunKeywordAnalysis_MultiYear()
    ' ----------------------------------------------------------
    ' 프로그램명 : RunKeywordAnalysis_MultiYear
    ' 설명       : 연도 수 제한 없이 순위 상승/하락, 신규, 사라진, 최신상승 키워드 분석
    ' 작성일자   : 2025-07-02
    ' 수정내용   : 분석 시트에 "일련번호" 열 추가 및 자동 채번, 최신상승키워드 시트 추가
    ' ----------------------------------------------------------

    Dim wsData As Worksheet: Set wsData = ActiveWorkbook.Sheets("데이터")
    Dim lastRow As Long: lastRow = wsData.Cells(wsData.Rows.count, "B").End(xlUp).Row

    Dim rankData As Object: Set rankData = CreateObject("Scripting.Dictionary")
    Dim allYears As Object: Set allYears = CreateObject("Scripting.Dictionary")

    Dim r As Long, keyword As Variant, yearStr As String, year As Integer, rank As Long
    Dim data As Object    ' ★★★ data 변수는 여기서 한 번만 선언! ★★★

    ' ------------------------------
    ' 1. 데이터 수집
    ' ------------------------------
    For r = 2 To lastRow
        keyword = Trim(wsData.Cells(r, "B").Value)
        yearStr = wsData.Cells(r, "C").Value
        If Len(yearStr) >= 4 Then year = val(Left(yearStr, 4)) Else year = 0
        rank = wsData.Cells(r, "A").Value

        If year > 0 Then
            If Not rankData.exists(keyword) Then
                Set rankData(keyword) = CreateObject("Scripting.Dictionary")
            End If
            rankData(keyword)(year) = rank
            allYears(year) = True
        End If
    Next r

    ' 연도 정렬
    Dim sortedYears() As Long
    ReDim sortedYears(0 To allYears.count - 1)
    Dim i As Integer: i = 0
    Dim y As Variant
    For Each y In allYears.Keys
        sortedYears(i) = y
        i = i + 1
    Next y
    Call QuickSortLong(sortedYears, LBound(sortedYears), UBound(sortedYears))
    Dim latestYear As Long: latestYear = sortedYears(UBound(sortedYears))

    ' ------------------------------
    ' 2. 기존 시트 제거
    ' ------------------------------
    Dim shtNames: shtNames = Array("순위상승키워드", "순위하락키워드", "사라진키워드", "신규키워드", "키워드분석_요약보고서", "최근상승키워드")
    Application.DisplayAlerts = False
    For Each y In shtNames
        On Error Resume Next: Sheets(y).Delete: On Error GoTo 0
    Next y
    Application.DisplayAlerts = True

    ' ------------------------------
    ' 3. 결과 시트 생성
    ' ------------------------------
    Dim wsUp As Worksheet: Set wsUp = Sheets.Add: wsUp.Name = "순위상승키워드"
    Dim wsDown As Worksheet: Set wsDown = Sheets.Add: wsDown.Name = "순위하락키워드"
    Dim wsGone As Worksheet: Set wsGone = Sheets.Add: wsGone.Name = "사라진키워드"
    Dim wsNew As Worksheet: Set wsNew = Sheets.Add: wsNew.Name = "신규키워드"

    ' ------------------------------
    ' 3-1. 일련번호 열 추가 및 제목 작성
    ' ------------------------------
    Dim wsList As Variant: wsList = Array(wsUp, wsDown, wsGone, wsNew)
    For i = 0 To 3
        With wsList(i)
            .Columns("A:A").Insert Shift:=xlToRight
            .Cells(1, 1).Value = "일련번호"
        End With
    Next i

    wsUp.Cells(1, 2).Value = "인기검색어"
    wsDown.Cells(1, 2).Value = "인기검색어"
    For i = 0 To UBound(sortedYears)
        wsUp.Cells(1, i + 3).Value = sortedYears(i) & " 순위"
        wsDown.Cells(1, i + 3).Value = sortedYears(i) & " 순위"
    Next i
    wsUp.Cells(1, i + 3).Value = "순위 개선폭"
    wsDown.Cells(1, i + 3).Value = "순위 하락폭"

    wsGone.Range("B1:C1").Value = Array("인기검색어", "마지막 등장년도")
    wsNew.Range("B1:C1").Value = Array("신규 키워드", latestYear & " 순위")

    ' ------------------------------
    ' 4. 분석
    ' ------------------------------
    Dim iUp As Long: iUp = 2
    Dim iDown As Long: iDown = 2
    Dim iGone As Long: iGone = 2
    Dim iNew As Long: iNew = 2

    Dim goneDict As Object: Set goneDict = CreateObject("Scripting.Dictionary")

    For Each keyword In rankData.Keys
        Set data = rankData(keyword)
        Dim available() As Long: ReDim available(0 To data.count - 1)
        i = 0
        For Each y In sortedYears
            If data.exists(y) Then
                available(i) = y
                i = i + 1
            End If
        Next y

        If i >= 3 Then
            Dim upFlag As Boolean: upFlag = True
            Dim downFlag As Boolean: downFlag = True
            Dim j As Integer
            For j = 1 To i - 1
                If data(available(j - 1)) <= data(available(j)) Then upFlag = False
                If data(available(j - 1)) >= data(available(j)) Then downFlag = False
            Next j
            If upFlag Then
                wsUp.Cells(iUp, 1).Value = iUp - 1
                wsUp.Cells(iUp, 2).Value = keyword
                For j = 0 To UBound(sortedYears)
                    If data.exists(sortedYears(j)) Then
                        wsUp.Cells(iUp, j + 3).Value = data(sortedYears(j))
                    End If
                Next j
                wsUp.Cells(iUp, j + 3).Value = data(available(0)) - data(available(i - 1))
                iUp = iUp + 1
            ElseIf downFlag Then
                wsDown.Cells(iDown, 1).Value = iDown - 1
                wsDown.Cells(iDown, 2).Value = keyword
                For j = 0 To UBound(sortedYears)
                    If data.exists(sortedYears(j)) Then
                        wsDown.Cells(iDown, j + 3).Value = data(sortedYears(j))
                    End If
                Next j
                wsDown.Cells(iDown, j + 3).Value = data(available(i - 1)) - data(available(0))
                iDown = iDown + 1
            End If
        End If

        If Not data.exists(latestYear) Then
            Dim maxY As Long: maxY = 0
            For Each y In data.Keys
                If y > maxY Then maxY = y
            Next y
            goneDict(keyword) = maxY
        End If

        If data.count = 1 And data.exists(latestYear) Then
            wsNew.Cells(iNew, 1).Value = iNew - 1
            wsNew.Cells(iNew, 2).Value = keyword
            wsNew.Cells(iNew, 3).Value = data(latestYear)
            iNew = iNew + 1
        End If
    Next keyword

    For Each keyword In goneDict.Keys
        wsGone.Cells(iGone, 1).Value = iGone - 1
        wsGone.Cells(iGone, 2).Value = keyword
        wsGone.Cells(iGone, 3).Value = goneDict(keyword)
        iGone = iGone + 1
    Next keyword

    ' ------------------------------
    ' 4-1. 최신년도에 단 한 번이라도 순위가 오른 키워드 시트
    ' ------------------------------
    Dim wsRecentUp As Worksheet: Set wsRecentUp = Sheets.Add: wsRecentUp.Name = "최근상승키워드"
    wsRecentUp.Cells(1, 1).Value = "일련번호"
    wsRecentUp.Cells(1, 2).Value = "인기검색어"
    wsRecentUp.Cells(1, 3).Value = "역대 최고 순위(과거)"
    wsRecentUp.Cells(1, 4).Value = latestYear & " 순위"
    wsRecentUp.Cells(1, 5).Value = "순위 개선폭"

    Dim iRecentUp As Long: iRecentUp = 2

    For Each keyword In rankData.Keys
        Set data = rankData(keyword)
        If data.exists(latestYear) And data.count > 1 Then
            Dim bestOldRank As Variant: bestOldRank = 1000000
            For Each y In data.Keys
                If y <> latestYear Then
                    If data(y) < bestOldRank Then bestOldRank = data(y)
                End If
            Next y
            If data(latestYear) < bestOldRank Then
                wsRecentUp.Cells(iRecentUp, 1).Value = iRecentUp - 1
                wsRecentUp.Cells(iRecentUp, 2).Value = keyword
                wsRecentUp.Cells(iRecentUp, 3).Value = bestOldRank
                wsRecentUp.Cells(iRecentUp, 4).Value = data(latestYear)
                wsRecentUp.Cells(iRecentUp, 5).Value = bestOldRank - data(latestYear)
                iRecentUp = iRecentUp + 1
            End If
        End If
    Next keyword

    ' 제목 서식 적용
    With wsRecentUp.Range("A1:E1")
        .Font.Bold = True
        .Interior.Color = RGB(204, 255, 229)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    wsRecentUp.Columns.AutoFit
    wsRecentUp.Range("A1:E" & iRecentUp - 1).Borders.LineStyle = xlContinuous

    ' ------------------------------
    ' 5. 요약 보고서
    ' ------------------------------
    Dim wsReport As Worksheet: Set wsReport = Sheets.Add: wsReport.Name = "키워드분석_요약보고서"
    wsReport.Range("A1").Value = "★ 다년간 네이버 데이타랩 키워드 분석 요약 보고서"
    wsReport.Range("A1").Font.Bold = True
    wsReport.Range("A1").Font.Size = 14
    wsReport.Range("A3:E3").Value = Array("항목", "키워드 수", "대표 키워드", "분석 시트로 이동", "비고")
    wsReport.Range("A3:E3").Font.Bold = True

    Dim counts(1 To 5) As Long, examples(1 To 5) As String
    Dim snames: snames = Array("순위상승키워드", "순위하락키워드", "사라진키워드", "신규키워드", "최근상승키워드")
    For i = 0 To 4
        With Worksheets(snames(i))
            counts(i + 1) = .Cells(.Rows.count, "B").End(xlUp).Row - 1
            If counts(i + 1) > 0 Then
                examples(i + 1) = .Cells(2, 2).Value
            Else
                examples(i + 1) = "(데이터 없음)"
            End If
        End With
    Next i

    For i = 0 To 4
        wsReport.Cells(i + 4, 1).Value = Choose(i + 1, "순위 상승 키워드", "순위 하락 키워드", "사라진 키워드", "신규 키워드", "최근 상승 키워드")
        wsReport.Cells(i + 4, 2).Value = counts(i + 1)
        wsReport.Cells(i + 4, 3).Value = examples(i + 1)
        wsReport.Hyperlinks.Add Anchor:=wsReport.Cells(i + 4, 4), _
            Address:="", SubAddress:="'" & snames(i) & "'!A1", _
            TextToDisplay:="이동"
        wsReport.Cells(i + 4, 5).Value = ""
    Next i

    With wsReport.Range("A3:E8")
        .Columns.AutoFit
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' ------------------------------
    ' 6. 시트 서식 적용 (제목 + 전체 테두리 + 자동 너비)
    ' ------------------------------
    Dim pastelColors As Variant
    pastelColors = Array(RGB(204, 255, 229), RGB(255, 230, 255), RGB(255, 255, 204), RGB(221, 235, 247), RGB(204, 255, 229)) ' 5개 색

    Dim wsTitles As Variant
    wsTitles = Array(wsUp, wsDown, wsGone, wsNew, wsRecentUp)

    Dim colCount As Long, rowCount As Long
    For i = 0 To 4
        With wsTitles(i)
            colCount = .Cells(1, .Columns.count).End(xlToLeft).column
            rowCount = .Cells(.Rows.count, "A").End(xlUp).Row

            ' 제목 행 서식
            With .Range(.Cells(1, 1), .Cells(1, colCount))
                .Font.Bold = True
                .Interior.Color = pastelColors(i)
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With

            ' 전체 셀 테두리
            With .Range(.Cells(1, 1), .Cells(rowCount, colCount))
                .Borders.LineStyle = xlContinuous
            End With

            .Columns.AutoFit
        End With
    Next i

    MsgBox "★ 연도 제한 없이 분석 및 보고서 생성 완료!", vbInformation
End Sub

' 연도 정렬용 퀵소트 함수
Sub QuickSortLong(arr() As Long, ByVal first As Long, ByVal last As Long)
    Dim low As Long, high As Long, pivot As Long, temp As Long
    low = first: high = last
    pivot = arr((first + last) \ 2)
    Do While low <= high
        Do While arr(low) < pivot: low = low + 1: Loop
        Do While arr(high) > pivot: high = high - 1: Loop
        If low <= high Then
            temp = arr(low): arr(low) = arr(high): arr(high) = temp
            low = low + 1: high = high - 1
        End If
    Loop
    If first < high Then QuickSortLong arr, first, high
    If low < last Then QuickSortLong arr, low, last
End Sub


