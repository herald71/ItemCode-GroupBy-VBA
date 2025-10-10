Attribute VB_Name = "Module16"
Sub MergeSheetsIntoOne()
    Dim ws As Worksheet
    Dim wsMaster As Worksheet
    Dim lastRow As Long
    Dim PasteRow As Long
    Dim i As Integer
    Dim DataStartRow As Long
    Dim wb As Workbook

    ' 현재 활성화된 워크북 참조
    Set wb = ActiveWorkbook

    ' 기존에 "Master" 시트가 있다면 삭제
    On Error Resume Next
    Application.DisplayAlerts = False
    wb.Sheets("Master").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    ' 새로운 마스터 시트 생성
    Set wsMaster = wb.Sheets.Add
    wsMaster.Name = "Master"
    
    ' 첫 번째 시트의 제목행을 형식 포함하여 마스터 시트에 복사
    With wb.Sheets(1)
        .Rows(1).Copy
        wsMaster.Cells(1, 1).PasteSpecial Paste:=xlPasteAll ' 값과 형식 모두 복사
    End With

    ' 첫 번째 데이터가 들어갈 행 번호 설정 (2번째 행부터 시작)
    PasteRow = 2
    
    ' 첫 번째 시트의 모든 데이터 복사 (제목 행 제외)
    Set ws = wb.Sheets(1)
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    If lastRow > 1 Then
        ws.Range(ws.Rows(2), ws.Rows(lastRow)).Copy
        wsMaster.Cells(PasteRow, 1).PasteSpecial Paste:=xlPasteValues
        PasteRow = wsMaster.Cells(wsMaster.Rows.count, 1).End(xlUp).Row + 1
    End If

    ' 2번째 시트부터 마지막 시트까지 순회 (첫 번째 시트는 제외)
    For i = 2 To wb.Sheets.count
        If wb.Sheets(i).Name <> "Master" Then ' "Master" 시트는 제외
            Set ws = wb.Sheets(i)
            
            ' 현재 시트의 마지막 행 계산 (제목행을 제외한 데이터만 복사)
            lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
            
            ' 시트에 데이터가 있는 경우만 실행
            If lastRow > 1 Then
                DataStartRow = 2 ' 제목행을 제외한 데이터는 2번째 행부터 시작
                ws.Range(ws.Rows(DataStartRow), ws.Rows(lastRow)).Copy
                wsMaster.Cells(PasteRow, 1).PasteSpecial Paste:=xlPasteValues
                
                ' 붙여넣기 후 다음 행으로 이동
                PasteRow = wsMaster.Cells(wsMaster.Rows.count, 1).End(xlUp).Row + 1
            End If
        End If
    Next i

    ' 복사 후 클립보드 비우기
    Application.CutCopyMode = False
    
    MsgBox "모든 시트를 'Master' 시트에 성공적으로 합쳤습니다!"
End Sub

