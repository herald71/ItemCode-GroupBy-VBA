Attribute VB_Name = "Module37"
Sub CopyColoredCellsFromColumnB()
    ' --------------------------------------------------------
    ' 프로그램명 : B열의 채우기 색 있는 셀을 B열검색어 시트로 복사
    ' 작성일자   : 2025-07-11
    ' 설명       : 주요 키워드 시트들에서 B열의 채우기 색 있는 셀만 추출하여
    '              'B열검색어' 시트에 일련번호, 값, 시트명으로 정리
    '              복사 순서는 지정된 순서에 따름
    ' --------------------------------------------------------

    Dim destSheet As Worksheet
    Dim ws As Worksheet
    Dim cell As Range
    Dim r As Long
    Dim lastRow As Long
    Dim destRow As Long: destRow = 2
    Dim wb As Workbook
    Set wb = ActiveWorkbook

    Application.ScreenUpdating = False

    ' 기존 시트 초기화 또는 생성
    On Error Resume Next
    Set destSheet = wb.Sheets("B열검색어")
    If destSheet Is Nothing Then
        Set destSheet = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.count))
        destSheet.Name = "B열검색어"
    Else
        destSheet.Cells.Clear
    End If
    On Error GoTo 0

    ' 제목 입력
    With destSheet
        .Cells(1, 1).Value = "일련번호"
        .Cells(1, 2).Value = "인기검색어"
        .Cells(1, 3).Value = "가져온 시트"
        .Rows(1).Font.Bold = True
    End With

    ' 모든 시트 순회
    For Each ws In wb.Worksheets
        ' 결과 시트 제외
        If ws.Name <> destSheet.Name Then
            With ws
                lastRow = .Cells(.Rows.count, "B").End(xlUp).Row
                For r = 2 To lastRow
                    Set cell = .Cells(r, "B")
                    If cell.Interior.colorIndex <> xlNone And cell.Interior.colorIndex <> -4142 Then
                        If Trim(cell.Value) <> "" Then
                            destSheet.Cells(destRow, 1).Value = destRow - 1        ' 일련번호
                            destSheet.Cells(destRow, 2).Value = cell.Value         ' 인기검색어
                            destSheet.Cells(destRow, 3).Value = ws.Name            ' 시트명
                            destRow = destRow + 1
                        End If
                    End If
                Next r
            End With
        End If
    Next ws

    destSheet.Columns("A:C").AutoFit

    Application.ScreenUpdating = True

    MsgBox "B열검색어 시트로 복사 완료! 총 " & destRow - 2 & "개 항목", vbInformation

End Sub

