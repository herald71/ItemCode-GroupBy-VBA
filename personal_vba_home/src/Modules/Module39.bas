Attribute VB_Name = "Module39"
Sub 월별빈도수_순위_비고_테두리_그래프()
    Dim ws As Worksheet
    Dim summaryWs As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim monthVal As Variant
    Dim monthCount(1 To 12) As Long
    Dim rankVal As Integer
    Dim lastDataRow As Long
    Dim chartObj As ChartObject

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.count, "N").End(xlUp).Row

    ' 월별 빈도수 집계
    For i = 2 To lastRow
        monthVal = ws.Cells(i, "N").Value
        If IsNumeric(monthVal) Then
            If monthVal >= 1 And monthVal <= 12 Then
                monthCount(monthVal) = monthCount(monthVal) + 1
            End If
        End If
    Next i

    ' 기존 시트 삭제
    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets("월별 빈도수").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    ' 새 시트 생성
    Set summaryWs = Worksheets.Add
    summaryWs.Name = "월별 빈도수"

    ' 헤더 작성
    With summaryWs
        .Range("A1").Value = "월"
        .Range("B1").Value = "빈도수"
        .Range("C1").Value = "순위"
        .Range("D1").Value = "비고"
        .Range("A1:D1").Font.Bold = True
    End With

    ' 데이터 입력
    For i = 1 To 12
        summaryWs.Cells(i + 1, 1).Value = i & "월"
        summaryWs.Cells(i + 1, 2).Value = monthCount(i)
    Next i

    ' 빈도수 서식
    summaryWs.Range("B2:B13").NumberFormat = "#,##0"

    ' 순위 계산
    summaryWs.Range("C2:C13").FormulaR1C1 = "=RANK(RC[-1], R2C2:R13C2, 0)"

    ' 비고 및 색상 적용
    For i = 2 To 13
        rankVal = summaryWs.Cells(i, 3).Value
        Select Case rankVal
            Case 1 To 3
                summaryWs.Range("A" & i & ":D" & i).Interior.Color = RGB(255, 204, 0) ' 성수기
                summaryWs.Cells(i, 4).Value = "성수기"
            Case 10 To 12
                summaryWs.Range("A" & i & ":D" & i).Interior.Color = RGB(204, 255, 255) ' 비수기
                summaryWs.Cells(i, 4).Value = "비수기"
        End Select
    Next i

    ' 실제 데이터 범위 계산
    lastDataRow = summaryWs.Cells(summaryWs.Rows.count, "A").End(xlUp).Row

    ' 테두리 적용
    With summaryWs.Range("A1:D" & lastDataRow)
        .Columns.AutoFit
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
    End With

    ' ? 차트 삽입 (꺾은선형)
    Set chartObj = summaryWs.ChartObjects.Add(Left:=300, Width:=500, Top:=30, Height:=300)
    With chartObj.Chart
        .ChartType = xlLine
        .SetSourceData Source:=summaryWs.Range("A1:B13")
        .HasTitle = True
        .ChartTitle.text = "빈도수"
        .Axes(xlCategory).HasTitle = False
        .Axes(xlValue).HasTitle = False
    End With

    MsgBox "? 모든 작업 완료: 시트 + 성수기/비수기 + 꺾은선형 차트까지 생성되었습니다!", vbInformation
End Sub

