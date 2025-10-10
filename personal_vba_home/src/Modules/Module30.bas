Attribute VB_Name = "Module30"
Option Explicit

Sub AnalyzeKeywords_2개분석()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim basePeriod As String
    Dim comparePeriod As String
    Dim proceed As VbMsgBoxResult
    Dim wb As Workbook

    ' 현재 활성화된 워크북에서 작업
    Set wb = ActiveWorkbook
    On Error Resume Next
    Set ws = wb.Sheets("데이터")
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "데이터 시트를 찾을 수 없습니다. '데이터'라는 이름의 시트를 확인하세요.", vbExclamation
        Exit Sub
    End If
    
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row

    ' 사용자로부터 기간 입력 받기
    basePeriod = InputBox("분석할 기준 기간을 입력하세요 (예: 2024년11월):", "기준 기간 입력")
    If basePeriod = "" Then Exit Sub ' 입력 취소 시 종료

    comparePeriod = InputBox("비교할 기간을 입력하세요 (예: 2023년11월):", "비교 기간 입력")
    If comparePeriod = "" Then Exit Sub ' 입력 취소 시 종료

    ' 사용자 확인
    proceed = MsgBox("기준 기간: " & basePeriod & vbCrLf & _
                     "비교 기간: " & comparePeriod & vbCrLf & _
                     "분석을 진행하시겠습니까?", vbYesNo + vbQuestion)
    If proceed = vbNo Then Exit Sub

    ' 각 서브루틴 실행
    ExtractNewKeywords_Personal ws, lastRow, basePeriod, comparePeriod
    ExtractRisingKeywords_Personal ws, lastRow, basePeriod, comparePeriod

    MsgBox "모든 분석이 완료되었습니다.", vbInformation
End Sub

Sub ExtractNewKeywords_Personal(ws As Worksheet, lastRow As Long, basePeriod As String, comparePeriod As String)
    Dim newSheet As Worksheet
    Dim i As Long
    Dim keyword As Variant
    Dim outputRow As Long
    Dim keywordsPrev As Object
    Dim keywordsBase As Object
    Dim sheetName As String

    Set keywordsPrev = CreateObject("Scripting.Dictionary")
    Set keywordsBase = CreateObject("Scripting.Dictionary")

    ' 비교 기간의 키워드 로드
    For i = 2 To lastRow
        If ws.Cells(i, 3).Value = comparePeriod Then
            keyword = ws.Cells(i, 2).Value
            keywordsPrev(keyword) = True
        End If
    Next i

    ' 기준 기간의 키워드 로드
    For i = 2 To lastRow
        If ws.Cells(i, 3).Value = basePeriod Then
            keyword = ws.Cells(i, 2).Value
            keywordsBase(keyword) = ws.Cells(i, 1).Value ' 순위 저장
        End If
    Next i

    ' 결과 시트 생성
    sheetName = basePeriod & " 신규 검색어"
    Set newSheet = CreateResultSheet_Personal(sheetName)
    newSheet.Cells(1, 1).Value = "순위"
    newSheet.Cells(1, 2).Value = "인기검색어"
    outputRow = 2

    ' 신규 검색어 추출
    For Each keyword In keywordsBase.Keys
        If Not keywordsPrev.exists(keyword) Then
            newSheet.Cells(outputRow, 1).Value = keywordsBase(keyword) ' 순위
            newSheet.Cells(outputRow, 2).Value = keyword ' 인기검색어
            outputRow = outputRow + 1
        End If
    Next keyword

    MsgBox "총 " & (outputRow - 2) & "개의 신규 키워드가 '" & sheetName & "' 시트에 작성되었습니다.", vbInformation
End Sub

Sub ExtractRisingKeywords_Personal(ws As Worksheet, lastRow As Long, basePeriod As String, comparePeriod As String)
    Dim newSheet As Worksheet
    Dim i As Long
    Dim keyword As Variant
    Dim outputRow As Long
    Dim keywordsCompare As Object
    Dim keywordsBase As Object
    Dim rankCompare As Long
    Dim rankBase As Long
    Dim sheetName As String

    Set keywordsCompare = CreateObject("Scripting.Dictionary")
    Set keywordsBase = CreateObject("Scripting.Dictionary")

    ' 비교 기간의 키워드와 순위 로드
    For i = 2 To lastRow
        If ws.Cells(i, 3).Value = comparePeriod Then
            keyword = ws.Cells(i, 2).Value
            keywordsCompare(keyword) = ws.Cells(i, 1).Value ' 순위 저장
        End If
    Next i

    ' 기준 기간의 키워드와 순위 로드
    For i = 2 To lastRow
        If ws.Cells(i, 3).Value = basePeriod Then
            keyword = ws.Cells(i, 2).Value
            keywordsBase(keyword) = ws.Cells(i, 1).Value ' 순위 저장
        End If
    Next i

    ' 결과 시트 생성
    sheetName = comparePeriod & " 대비 순위 상승 검색어"
    Set newSheet = CreateResultSheet_Personal(sheetName)
    newSheet.Cells(1, 1).Value = basePeriod & "_순위"
    newSheet.Cells(1, 2).Value = "인기검색어"
    newSheet.Cells(1, 3).Value = "순위변동"
    outputRow = 2

    ' 순위 상승 검색어 추출
    For Each keyword In keywordsBase.Keys
        If keywordsCompare.exists(keyword) Then
            rankCompare = keywordsCompare(keyword)
            rankBase = keywordsBase(keyword)
            If IsNumeric(rankCompare) And IsNumeric(rankBase) Then
                If rankBase < rankCompare Then
                    newSheet.Cells(outputRow, 1).Value = rankBase ' 기준 기간 순위
                    newSheet.Cells(outputRow, 2).Value = keyword ' 인기검색어
                    newSheet.Cells(outputRow, 3).Value = rankCompare - rankBase ' 순위변동
                    outputRow = outputRow + 1
                End If
            End If
        End If
    Next keyword

    MsgBox "총 " & (outputRow - 2) & "개의 순위 상승 키워드가 '" & sheetName & "' 시트에 작성되었습니다.", vbInformation
End Sub

Function CreateResultSheet_Personal(sheetName As String) As Worksheet
    Dim ws As Worksheet
    ' 기존 시트가 있으면 삭제
    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets(sheetName).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    ' 새로운 시트 생성
    Set ws = ActiveWorkbook.Sheets.Add
    ws.Name = sheetName
    Set CreateResultSheet_Personal = ws
End Function




