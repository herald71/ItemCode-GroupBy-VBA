Attribute VB_Name = "Module16"
Sub MergeSheetsIntoOne()
    Dim ws As Worksheet
    Dim wsMaster As Worksheet
    Dim lastRow As Long
    Dim PasteRow As Long
    Dim i As Integer
    Dim DataStartRow As Long
    Dim wb As Workbook

    ' ���� Ȱ��ȭ�� ��ũ�� ����
    Set wb = ActiveWorkbook

    ' ������ "Master" ��Ʈ�� �ִٸ� ����
    On Error Resume Next
    Application.DisplayAlerts = False
    wb.Sheets("Master").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    ' ���ο� ������ ��Ʈ ����
    Set wsMaster = wb.Sheets.Add
    wsMaster.Name = "Master"
    
    ' ù ��° ��Ʈ�� �������� ���� �����Ͽ� ������ ��Ʈ�� ����
    With wb.Sheets(1)
        .Rows(1).Copy
        wsMaster.Cells(1, 1).PasteSpecial Paste:=xlPasteAll ' ���� ���� ��� ����
    End With

    ' ù ��° �����Ͱ� �� �� ��ȣ ���� (2��° ����� ����)
    PasteRow = 2
    
    ' ù ��° ��Ʈ�� ��� ������ ���� (���� �� ����)
    Set ws = wb.Sheets(1)
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    If lastRow > 1 Then
        ws.Range(ws.Rows(2), ws.Rows(lastRow)).Copy
        wsMaster.Cells(PasteRow, 1).PasteSpecial Paste:=xlPasteValues
        PasteRow = wsMaster.Cells(wsMaster.Rows.count, 1).End(xlUp).Row + 1
    End If

    ' 2��° ��Ʈ���� ������ ��Ʈ���� ��ȸ (ù ��° ��Ʈ�� ����)
    For i = 2 To wb.Sheets.count
        If wb.Sheets(i).Name <> "Master" Then ' "Master" ��Ʈ�� ����
            Set ws = wb.Sheets(i)
            
            ' ���� ��Ʈ�� ������ �� ��� (�������� ������ �����͸� ����)
            lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
            
            ' ��Ʈ�� �����Ͱ� �ִ� ��츸 ����
            If lastRow > 1 Then
                DataStartRow = 2 ' �������� ������ �����ʹ� 2��° ����� ����
                ws.Range(ws.Rows(DataStartRow), ws.Rows(lastRow)).Copy
                wsMaster.Cells(PasteRow, 1).PasteSpecial Paste:=xlPasteValues
                
                ' �ٿ��ֱ� �� ���� ������ �̵�
                PasteRow = wsMaster.Cells(wsMaster.Rows.count, 1).End(xlUp).Row + 1
            End If
        End If
    Next i

    ' ���� �� Ŭ������ ����
    Application.CutCopyMode = False
    
    MsgBox "��� ��Ʈ�� 'Master' ��Ʈ�� ���������� ���ƽ��ϴ�!"
End Sub

