Attribute VB_Name = "Module36"
' ===============================================
' 프로그램명 : 키워드 변화 자동 분석 도구
' 작성일자   : 2025-07-19
' 버전       : v2.1 (결과 시트 서식 + 완전신규어 색상 강조)
' 설명       : '데이터' 시트의 C열(기간)을 자동 인식하여
'              기준/비교기간을 정해 신규/상승 키워드를 분석하고
'              결과 시트를 생성 및 요약 보고서 시트를 만듦
' ===============================================

Option Explicit

' ▶ 메인 실행 프로시저: C열에서 3개 최신 기간을 인식해 분석 시작
Sub AnalyzeKeywords_AutoPeriod()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim wb As Workbook
    Dim dictPeriods As Object
    Dim periodList() As Variant
    Dim i As Long
    Dim cellValue As String
    Dim basePeriod As String, comparePeriod1 As String, comparePeriod2 As String

    Set wb = ActiveWorkbook
    Set ws = wb.Sheets("데이터")
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row

    ' --- C열의 기간 정보 수집 ---
    Set dictPeriods = CreateObject("Scripting.Dictionary")
    For i = 2 To lastRow
        cellValue = Trim(ws.Cells(i, 3).Value)
        If Len(cellValue) > 0 Then dictPeriods(cellValue) = True
    Next i

    ' --- 기간이 3개 이상 있어야 분석 가능 ---
    If dictPeriods.count < 3 Then
        MsgBox "기간 정보가 3개 이상 존재하지 않습니다.", vbExclamation
        Exit Sub
    End If

    ' --- 내림차순 정렬: 최신 순으로 정렬하여 기준·비교시점 추출 ---
    periodList = dictPeriods.Keys
    SortDescending periodList
    basePeriod = periodList(0)
    comparePeriod1 = periodList(1)
    comparePeriod2 = periodList(2)

    MsgBox "기준 기간: " & basePeriod & vbCrLf & _
           "비교 기간 1: " & comparePeriod1 & vbCrLf & _
           "비교 기간 2: " & comparePeriod2, vbInformation

    ' --- 5가지 분석 수행 ---
    ExtractNewKeywordsBase ws, lastRow, basePeriod, comparePeriod1, comparePeriod2
    ExtractRisingKeywordsBase ws, lastRow, basePeriod, comparePeriod1
    ExtractNewKeywordsVsPeriod ws, lastRow, basePeriod, comparePeriod2
    ExtractNewKeywordsVsPeriod ws, lastRow, basePeriod, comparePeriod1
    ExtractRisingKeywordsVsPeriod ws, lastRow, basePeriod, comparePeriod2
    ExtractRisingFromPastToCurrent ws, lastRow, basePeriod, comparePeriod1, comparePeriod2

    ' --- 요약 보고서 생성 ---
    Call CreateFormattedSummaryReport
    MsgBox "모든 분석이 완료되었습니다.", vbInformation
End Sub

' ▶ 문자열 배열 내림차순 정렬 함수
Sub SortDescending(arr() As Variant)
    Dim i As Long, j As Long, temp As String
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i) < arr(j) Then
                temp = arr(i): arr(i) = arr(j): arr(j) = temp
            End If
        Next j
    Next i
End Sub

' ▶ 기준 기간에만 있는 키워드 추출 ("완전신규어")
Sub ExtractNewKeywordsBase(ws As Worksheet, lastRow As Long, basePeriod As String, comparePeriod1 As String, comparePeriod2 As String)
    Dim newSheet As Worksheet, i As Long, keyword As Variant, outputRow As Long
    Dim keywordsPrev As Object, keywordsBase As Object
    Dim sheetName As String

    Set keywordsPrev = CreateObject("Scripting.Dictionary")
    Set keywordsBase = CreateObject("Scripting.Dictionary")

    ' --- 비교 기간 1, 2에서 등장한 키워드 수집 ---
    For i = 2 To lastRow
        If ws.Cells(i, 3).Value = comparePeriod1 Or ws.Cells(i, 3).Value = comparePeriod2 Then
            keyword = ws.Cells(i, 2).Value
            keywordsPrev(keyword) = True
        End If
    Next i

    ' --- 기준 기간 키워드 수집 ---
    For i = 2 To lastRow
        If ws.Cells(i, 3).Value = basePeriod Then
            keyword = ws.Cells(i, 2).Value
            keywordsBase(keyword) = ws.Cells(i, 1).Value
        End If
    Next i

    ' --- 결과 시트 생성 및 출력 ---
    sheetName = basePeriod & " 완전신규어"
    Set newSheet = CreateResultSheet(sheetName)
    newSheet.Cells(1, 1).Value = "순위"
    newSheet.Cells(1, 2).Value = "인기검색어"
    outputRow = 2

    For Each keyword In keywordsBase.Keys
        If Not keywordsPrev.exists(keyword) Then
            newSheet.Cells(outputRow, 1).Value = keywordsBase(keyword)
            newSheet.Cells(outputRow, 2).Value = keyword
            outputRow = outputRow + 1
        End If
    Next keyword

    Call ApplySheetFormatting(newSheet, True)
End Sub

' ▶ 기준 기간에서 순위가 상승한 키워드 추출
Sub ExtractRisingKeywordsBase(ws As Worksheet, lastRow As Long, basePeriod As String, comparePeriod As String)
    Dim newSheet As Worksheet, i As Long, keyword As Variant, outputRow As Long
    Dim keywordsCompare As Object, keywordsBase As Object
    Dim rankCompare As Long, rankBase As Long
    Dim sheetName As String

    Set keywordsCompare = CreateObject("Scripting.Dictionary")
    Set keywordsBase = CreateObject("Scripting.Dictionary")

    For i = 2 To lastRow
        If ws.Cells(i, 3).Value = comparePeriod Then
            keywordsCompare(ws.Cells(i, 2).Value) = ws.Cells(i, 1).Value
        End If
    Next i

    For i = 2 To lastRow
        If ws.Cells(i, 3).Value = basePeriod Then
            keywordsBase(ws.Cells(i, 2).Value) = ws.Cells(i, 1).Value
        End If
    Next i

    sheetName = comparePeriod & " 대비순위상승검색어"
    Set newSheet = CreateResultSheet(sheetName)
    newSheet.Cells(1, 1).Value = basePeriod & "_순위"
    newSheet.Cells(1, 2).Value = "인기검색어"
    newSheet.Cells(1, 3).Value = "순위변동"
    newSheet.Cells(1, 4).Value = comparePeriod & "_순위"
    outputRow = 2

    For Each keyword In keywordsBase.Keys
        If keywordsCompare.exists(keyword) Then
            rankBase = keywordsBase(keyword)
            rankCompare = keywordsCompare(keyword)
            If IsNumeric(rankBase) And IsNumeric(rankCompare) Then
                If rankBase < rankCompare Then
                    newSheet.Cells(outputRow, 1).Value = rankBase
                    newSheet.Cells(outputRow, 2).Value = keyword
                    newSheet.Cells(outputRow, 3).Value = rankCompare - rankBase
                    newSheet.Cells(outputRow, 4).Value = rankCompare
                    outputRow = outputRow + 1
                End If
            End If
        End If
    Next keyword

    Call ApplySheetFormatting(newSheet)
End Sub

' ▶ 기준 기간에만 있는 신규 키워드 추출 (비교기간 1개 대상)
Sub ExtractNewKeywordsVsPeriod(ws As Worksheet, lastRow As Long, basePeriod As String, comparePeriod As String)
    Dim newSheet As Worksheet, i As Long, keyword As Variant, outputRow As Long
    Dim keywordsCompare As Object, keywordsBase As Object
    Dim sheetName As String

    Set keywordsCompare = CreateObject("Scripting.Dictionary")
    Set keywordsBase = CreateObject("Scripting.Dictionary")

    For i = 2 To lastRow
        If ws.Cells(i, 3).Value = comparePeriod Then
            keywordsCompare(ws.Cells(i, 2).Value) = True
        End If
    Next i

    For i = 2 To lastRow
        If ws.Cells(i, 3).Value = basePeriod Then
            keywordsBase(ws.Cells(i, 2).Value) = ws.Cells(i, 1).Value
        End If
    Next i

    sheetName = comparePeriod & " 대비신규검색어"
    Set newSheet = CreateResultSheet(sheetName)
    newSheet.Cells(1, 1).Value = "순위"
    newSheet.Cells(1, 2).Value = "인기검색어"
    outputRow = 2

    For Each keyword In keywordsBase.Keys
        If Not keywordsCompare.exists(keyword) Then
            newSheet.Cells(outputRow, 1).Value = keywordsBase(keyword)
            newSheet.Cells(outputRow, 2).Value = keyword
            outputRow = outputRow + 1
        End If
    Next keyword

    Call ApplySheetFormatting(newSheet)
End Sub

' ▶ 기준 대비 비교 기간의 순위 상승 키워드 추출
Sub ExtractRisingKeywordsVsPeriod(ws As Worksheet, lastRow As Long, basePeriod As String, comparePeriod As String)
    Dim newSheet As Worksheet, i As Long, keyword As Variant, outputRow As Long
    Dim keywordsCompare As Object, keywordsBase As Object
    Dim rankCompare As Long, rankBase As Long
    Dim sheetName As String

    Set keywordsCompare = CreateObject("Scripting.Dictionary")
    Set keywordsBase = CreateObject("Scripting.Dictionary")

    For i = 2 To lastRow
        If ws.Cells(i, 3).Value = comparePeriod Then
            keywordsCompare(ws.Cells(i, 2).Value) = ws.Cells(i, 1).Value
        End If
    Next i

    For i = 2 To lastRow
        If ws.Cells(i, 3).Value = basePeriod Then
            keywordsBase(ws.Cells(i, 2).Value) = ws.Cells(i, 1).Value
        End If
    Next i

    sheetName = comparePeriod & " 대비순위상승검색어"
    Set newSheet = CreateResultSheet(sheetName)
    newSheet.Cells(1, 1).Value = basePeriod & "_순위"
    newSheet.Cells(1, 2).Value = "인기검색어"
    newSheet.Cells(1, 3).Value = "순위변동"
    newSheet.Cells(1, 4).Value = comparePeriod & "_순위"
    outputRow = 2

    For Each keyword In keywordsBase.Keys
        If keywordsCompare.exists(keyword) Then
            rankBase = keywordsBase(keyword)
            rankCompare = keywordsCompare(keyword)
            If IsNumeric(rankBase) And IsNumeric(rankCompare) Then
                If rankBase < rankCompare Then
                    newSheet.Cells(outputRow, 1).Value = rankBase
                    newSheet.Cells(outputRow, 2).Value = keyword
                    newSheet.Cells(outputRow, 3).Value = rankCompare - rankBase
                    newSheet.Cells(outputRow, 4).Value = rankCompare
                    outputRow = outputRow + 1
                End If
            End If
        End If
    Next keyword

    Call ApplySheetFormatting(newSheet)
End Sub

' ▶ 비교2 → 비교1 → 기준 순으로 꾸준히 상승한 키워드 추출
Sub ExtractRisingFromPastToCurrent(ws As Worksheet, lastRow As Long, basePeriod As String, comparePeriod1 As String, comparePeriod2 As String)
    Dim newSheet As Worksheet
    Dim i As Long, outputRow As Long
    Dim keyword As Variant
    Dim keywordsBase As Object, keywordsCompare1 As Object, keywordsCompare2 As Object
    Dim rankBase As Long, rank1 As Long, rank2 As Long
    Dim sheetName As String

    Set keywordsBase = CreateObject("Scripting.Dictionary")
    Set keywordsCompare1 = CreateObject("Scripting.Dictionary")
    Set keywordsCompare2 = CreateObject("Scripting.Dictionary")

    ' --- 기준 시점 키워드
    For i = 2 To lastRow
        If ws.Cells(i, 3).Value = basePeriod Then
            keywordsBase(ws.Cells(i, 2).Value) = ws.Cells(i, 1).Value
        End If
    Next i

    ' --- 비교1
    For i = 2 To lastRow
        If ws.Cells(i, 3).Value = comparePeriod1 Then
            keywordsCompare1(ws.Cells(i, 2).Value) = ws.Cells(i, 1).Value
        End If
    Next i

    ' --- 비교2
    For i = 2 To lastRow
        If ws.Cells(i, 3).Value = comparePeriod2 Then
            keywordsCompare2(ws.Cells(i, 2).Value) = ws.Cells(i, 1).Value
        End If
    Next i

    ' --- 시트 생성
    sheetName = "과거→현재_순위상승"
    Set newSheet = CreateResultSheet(sheetName)
    newSheet.Cells(1, 1).Value = comparePeriod2 & "_순위"
    newSheet.Cells(1, 2).Value = comparePeriod1 & "_순위"
    newSheet.Cells(1, 3).Value = basePeriod & "_순위"
    newSheet.Cells(1, 4).Value = "인기검색어"
    newSheet.Cells(1, 5).Value = "총상승폭"
    outputRow = 2

    ' --- 순위가 과거→현재로 갈수록 상승 (숫자 작아짐)
    For Each keyword In keywordsBase.Keys
        If keywordsCompare1.exists(keyword) And keywordsCompare2.exists(keyword) Then
            rankBase = keywordsBase(keyword)
            rank1 = keywordsCompare1(keyword)
            rank2 = keywordsCompare2(keyword)
            If IsNumeric(rankBase) And IsNumeric(rank1) And IsNumeric(rank2) Then
                If rank2 > rank1 And rank1 > rankBase Then
                    newSheet.Cells(outputRow, 1).Value = rank2
                    newSheet.Cells(outputRow, 2).Value = rank1
                    newSheet.Cells(outputRow, 3).Value = rankBase
                    newSheet.Cells(outputRow, 4).Value = keyword
                    newSheet.Cells(outputRow, 5).Value = rank2 - rankBase
                    outputRow = outputRow + 1
                End If
            End If
        End If
    Next keyword

    Call ApplySheetFormatting(newSheet)
End Sub




' ▶ 결과 시트 생성
Function CreateResultSheet(sheetName As String) As Worksheet
    Dim ws As Worksheet, cleanName As String
    cleanName = Replace(sheetName, "\", "")
    cleanName = Replace(cleanName, "/", "")
    cleanName = Replace(cleanName, "*", "")
    cleanName = Replace(cleanName, "[", "")
    cleanName = Replace(cleanName, "]", "")
    cleanName = Replace(cleanName, ":", "")
    cleanName = Replace(cleanName, "?", "")
    If Len(cleanName) > 31 Then cleanName = Left(cleanName, 31)

    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets(cleanName).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set ws = ActiveWorkbook.Sheets.Add
    ws.Name = cleanName
    Set CreateResultSheet = ws
End Function

' ▶ 결과 시트 서식 및 필터 적용 + 탭 색상 변경 + 테두리
Sub ApplySheetFormatting(ws As Worksheet, Optional isNewKeywordSheet As Boolean = False)
    Dim lastRow As Long, lastCol As Long
    Dim dataRange As Range

    With ws
        ' 마지막 행/열 계산
        lastRow = .Cells(.Rows.count, 1).End(xlUp).Row
        lastCol = .Cells(1, .Columns.count).End(xlToLeft).column

        If lastRow < 1 Or lastCol < 1 Then Exit Sub

        Set dataRange = .Range(.Cells(1, 1), .Cells(lastRow, lastCol))

        ' 제목 행 서식
        With .Range(.Cells(1, 1), .Cells(1, lastCol))
            .Font.Bold = True
            .Interior.Color = RGB(197, 217, 241) ' 파스텔 블루
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
        End With

        ' 자동 필터 적용
        .Range(.Cells(1, 1), .Cells(1, lastCol)).AutoFilter

        ' 열 너비 자동 맞춤
        .Columns("A:" & Split(.Cells(1, lastCol).Address, "$")(1)).AutoFit

        ' 테두리 적용
        With dataRange.Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .colorIndex = xlAutomatic
        End With

        ' ▶ 완전신규어 시트는 탭 색상만 빨간색 계열로 지정
        If isNewKeywordSheet Then
            .Tab.Color = RGB(255, 0, 0) ' 탭 색상 빨간색
        End If
    End With
End Sub


' ▶ 요약 보고서 시트 생성 및 하이퍼링크 포함 정리
Sub CreateFormattedSummaryReport()
    Dim wsSummary As Worksheet, ws As Worksheet
    Dim rowIndex As Long, keywordCount As Long
    Dim 기준시점 As String, 비교1 As String, 비교2 As String
    Dim periodList() As Variant, dictPeriods As Object
    Dim i As Long, lastRow As Long
    Dim wb As Workbook
    Dim 대표키워드 As String

    Set wb = ActiveWorkbook
    Set dictPeriods = CreateObject("Scripting.Dictionary")
    Set ws = wb.Sheets("데이터")
    lastRow = ws.Cells(ws.Rows.count, "C").End(xlUp).Row

    For i = 2 To lastRow
        If Len(Trim(ws.Cells(i, 3).Value)) > 0 Then
            dictPeriods(Trim(ws.Cells(i, 3).Value)) = True
        End If
    Next i

    If dictPeriods.count < 3 Then
        MsgBox "기간이 3개 이상 존재하지 않습니다.", vbExclamation
        Exit Sub
    End If

    periodList = dictPeriods.Keys
    SortDescending periodList
    기준시점 = periodList(0): 비교1 = periodList(1): 비교2 = periodList(2)

    ' 기존 요약 보고서 삭제
    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets("요약보고서").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set wsSummary = Worksheets.Add
    wsSummary.Name = "요약보고서"

    ' 상단 비교시점 테이블 생성
    With wsSummary
        .Range("A1").Value = "비교시점"
        .Range("B1").Value = "기간"
        .Range("A1:B1").Interior.Color = RGB(169, 209, 142)
        .Range("A1:B1").Font.Bold = True
        .Range("A1:B1").HorizontalAlignment = xlCenter
        .Range("A2").Value = "기준시점": .Range("B2").Value = 기준시점
        .Range("A3").Value = "1. 비교시점": .Range("B3").Value = 비교1
        .Range("A4").Value = "2. 비교시점": .Range("B4").Value = 비교2
        .Range("A1:A4").Borders.Weight = xlThin
        .Range("B1:B4").Borders.Weight = xlThin
    End With

    ' 분석 결과 테이블 헤더 (5열 구성)
    rowIndex = 6
    With wsSummary
        .Range("A" & rowIndex).Value = "분석 항목"
        .Range("B" & rowIndex).Value = "키워드 수"
        .Range("C" & rowIndex).Value = "대표 키워드"
        .Range("D" & rowIndex).Value = "분석 시트로 이동"
        .Range("E" & rowIndex).Value = "비고"
        .Range("A" & rowIndex & ":E" & rowIndex).Interior.Color = RGB(244, 176, 132)
        .Range("A" & rowIndex & ":E" & rowIndex).Font.Bold = True
        .Range("A" & rowIndex & ":E" & rowIndex).HorizontalAlignment = xlCenter
    End With

    ' 결과 시트 순회 및 요약정보 입력
    rowIndex = rowIndex + 1
    For Each ws In wb.Worksheets
        If ws.Name <> "데이터" And ws.Name <> "요약보고서" Then
            keywordCount = Application.WorksheetFunction.CountA(ws.Range("A:A")) - 1
            
            ' 대표 키워드: 일반 시트는 B2, 과거→현재_순위상승 시트는 D2
            If keywordCount > 0 Then
                If ws.Name = "과거→현재_순위상승" Then
                    대표키워드 = ws.Range("D2").Value
                Else
                    대표키워드 = ws.Range("B2").Value
                End If
                If Len(대표키워드) = 0 Then 대표키워드 = "(대표 키워드 없음)"
            Else
                대표키워드 = "(데이터 없음)"
            End If

            
            ' 값 삽입
            wsSummary.Cells(rowIndex, 1).Value = ws.Name
            wsSummary.Cells(rowIndex, 2).Value = keywordCount
            wsSummary.Cells(rowIndex, 3).Value = 대표키워드
            wsSummary.Hyperlinks.Add Anchor:=wsSummary.Cells(rowIndex, 4), Address:="", _
                SubAddress:="'" & ws.Name & "'!A1", TextToDisplay:="이동"

            ' 비고 설명
            Select Case True
                Case InStr(ws.Name, "완전신규") > 0
                    wsSummary.Cells(rowIndex, 5).Value = "기준시점에 완전 새롭게 등장한 키워드"
                Case InStr(ws.Name, "대비신규") > 0
                    If InStr(ws.Name, 비교1) > 0 Then
                        wsSummary.Cells(rowIndex, 5).Value = "1. 비교시점 대비 기준시점 신규 키워드"
                    ElseIf InStr(ws.Name, 비교2) > 0 Then
                        wsSummary.Cells(rowIndex, 5).Value = "2. 비교시점 대비 기준시점 신규 키워드"
                    End If
                Case InStr(ws.Name, "대비순위상승") > 0
                    If InStr(ws.Name, 비교1) > 0 Then
                        wsSummary.Cells(rowIndex, 5).Value = "1. 비교시점 대비 순위 상승"
                    ElseIf InStr(ws.Name, 비교2) > 0 Then
                        wsSummary.Cells(rowIndex, 5).Value = "2. 비교시점 대비 순위 상승"
                    End If
                    
                Case ws.Name = "과거→현재_순위상승"
                        wsSummary.Cells(rowIndex, 5).Value = "과거 → 현재로 순위가 꾸준히 상승한 키워드"
            End Select

            rowIndex = rowIndex + 1
        End If
    Next ws

    ' 열 너비 자동조정
    wsSummary.Columns("A:E").AutoFit
End Sub


