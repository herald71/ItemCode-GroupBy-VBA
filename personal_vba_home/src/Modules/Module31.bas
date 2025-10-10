Attribute VB_Name = "Module31"
Sub �ݺ��۾�()
'*******************************************************************************
' ��ũ�θ�: �ݺ��۾�
' �ۼ�����: 2025-04-21
' ��ɼ���:
' 1. �����Ͱ� �ִ� ������ �׵θ� ����
' 2. ù ��° �� ���� ����
'    - ��� ����
'    - �۲� ũ�� 12
'    - �� ���� 30
'    - ���� �����
'    - ���� �۲�
' 3. �ڵ� ���� ����
' 4. C�� ���� �������� ����
' 5. A���� ������ȣ �ο�
'
' �ٷ� ���� Ű: Ctrl+Shift+K
'
' �����̷�:
' 2025-04-21: ���� �ۼ�
' - �⺻ ���� ���� �� ���� ��� ����
' - ���� �ڵ� ���� �� ������ ��� �߰�
' - ���� ó�� �߰�
' 2025-04-21: ��� �߰�
' - ������ ���� �ڵ� ���� �߰�
' - A�� ���� �ڵ� �ο� �߰�
' - ���� ó�� ��ȭ
'*******************************************************************************

    On Error GoTo ErrorHandler
    
    '-------------------------------------------
    ' �ʱ� ����
    '-------------------------------------------
    ' ���� Ȱ��ȭ�� ��ũ��Ʈ ���� ����
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' ���� ����
    Dim lastRow As Long
    Dim i As Long
    Dim dataRange As Range
    Dim headerRange As Range
    
    ' ������ �� ã�� (B�� ����)
    lastRow = ws.Cells(ws.Rows.count, "B").End(xlUp).Row
    
    ' �����Ͱ� ���� ��� ó��
    If lastRow < 2 Then
        MsgBox "ó���� �����Ͱ� �����ϴ�.", vbExclamation
        Exit Sub
    End If
    
    ' ���� ����
    Set dataRange = ws.Range("A1:D" & lastRow)
    Set headerRange = ws.Range("A1:D1")
    
    ' ȭ�� ������Ʈ ���� (ó�� �ӵ� ���)
    Application.ScreenUpdating = False
    
    ' ������ ����� ���Ͱ� �ִٸ� ����
    If ws.AutoFilterMode Then
        ws.AutoFilterMode = False
    End If
    
    '-------------------------------------------
    ' �׵θ� ����
    '-------------------------------------------
    ' �����Ͱ� �ִ� ������ �׵θ� ����
    With dataRange
        .Borders.LineStyle = xlContinuous
        .Borders.colorIndex = 0
        .Borders.TintAndShade = 0
        .Borders.Weight = xlThin
    End With
    
    '-------------------------------------------
    ' ��� ��(1��) ���� ����
    '-------------------------------------------
    With headerRange
        ' ���� ����
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        
        ' �۲� ����
        With .Font
            .Size = 12
            .Bold = True
        End With
        
        ' ���� ����
        With .Interior
            .Color = 65535  ' �����
            .TintAndShade = 0
        End With
    End With
    
    '-------------------------------------------
    ' �� ���� �� ���� ����
    '-------------------------------------------
    ' 1���� ���̸� 30���� ����
    headerRange.RowHeight = 30
    
    '-------------------------------------------
    ' ������ ����
    '-------------------------------------------
    ' C�� ���� �������� ����
    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add2 key:=Range("C2:C" & lastRow), _
        SortOn:=xlSortOnValues, _
        Order:=xlDescending, _
        DataOption:=xlSortNormal
        
    With ws.Sort
        .SetRange dataRange
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ' �ڵ� ���� ����
    If Not ws.AutoFilterMode Then
        dataRange.AutoFilter
    End If
    
    '-------------------------------------------
    ' ���� �ο�
    '-------------------------------------------
    ' A���� ������ȣ �ο� (2����� ������ �����)
    With ws.Range("A2:A" & lastRow)
        ' ���� �ο�
        For i = 2 To lastRow
            ws.Cells(i, 1).Value = i - 1
        Next i
        
        ' ��� ����
        .HorizontalAlignment = xlCenter
    End With

ExitSub:
    ' ȭ�� ������Ʈ �簳
    Application.ScreenUpdating = True
    
    ' �۾� �Ϸ� �� F1 ���� �̵�
    ws.Range("F1").Select
    Exit Sub

ErrorHandler:
    ' ���� �߻� �� ó��
    MsgBox "������ �߻��߽��ϴ�." & vbNewLine & _
           "���� ����: " & Err.Description, vbCritical
    
    ' ȭ�� ������Ʈ �簳
    Application.ScreenUpdating = True
    Exit Sub
End Sub


