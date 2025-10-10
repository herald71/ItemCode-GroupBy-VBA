Attribute VB_Name = "Module13"
Sub 데이타열너비자동조절및테두리치기()
    ' 데이터 열 너비 자동 조절, 테두리 추가, 제목행 서식 적용

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim headerRange As Range

    ' 현재 활성화된 워크시트를 가져옵니다.
    Set ws = ActiveSheet

    ' 데이터 범위의 마지막 행과 열을 찾습니다.
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).column

    ' 모든 열의 너비를 자동으로 조정합니다.
    ws.Columns("A:" & Split(Cells(, lastCol).Address, "$")(1)).AutoFit

    ' 데이터 범위에 테두리를 추가합니다.
    With ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Borders
        .LineStyle = xlContinuous
        .colorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With

    ' 제목행 서식 적용
    Set headerRange = ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))
    With headerRange
        .Font.Bold = True
        .Font.Size = 13
        .HorizontalAlignment = xlCenter       ' 가운데 정렬
        .VerticalAlignment = xlCenter         ' 가운데 맞춤
        .WrapText = True                      ' 자동 줄바꿈
        .Interior.Color = RGB(189, 215, 238)  ' 파스텔 블루
    End With
End Sub



