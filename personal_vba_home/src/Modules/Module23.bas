Attribute VB_Name = "Module23"
Sub ���α���_Ű���庰_�м�()
    Dim ws As Worksheet, analysisWs As Worksheet
    Dim wb As Workbook
    Dim lastRow As Long, analysisRow As Long
    Dim dict As Object
    Dim key As String
    Dim rng As Range, cell As Range
    
    ' �۾��� ��ũ�� ����
    On Error Resume Next
    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        MsgBox "Ȱ��ȭ�� ��ũ���� �����ϴ�. ���� �м��� ������ ���� �ּ���.", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    
    ' "Sheet1"�� �м� ������� �ڵ� ����
    On Error Resume Next
    Set ws = wb.Sheets("Sheet1")
    On Error GoTo 0
    
    ' ��Ʈ ���� ���� Ȯ��
    If ws Is Nothing Then
        MsgBox """Sheet1""�� �������� �ʽ��ϴ�. �ùٸ� ������ ���� �ּ���.", vbCritical
        Exit Sub
    End If
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' ������ ���� ����
    lastRow = ws.Cells(ws.Rows.count, "M").End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "Ű���� �����Ͱ� �����ϴ�.", vbExclamation
        Exit Sub
    End If
    Set rng = ws.Range("M2:M" & lastRow) ' "Ű����"�� M���� �ִٰ� ����
    
    ' ������ ���� �� ����
    Dim dataArr As Variant
    For Each cell In rng
        key = Trim(cell.Value) ' ���� ������ Ű����
        
        If key <> "" Then ' �� �� ����
            If Not dict.exists(key) Then
                dict.Add key, Array(0, 0, 0, 0, 0) ' �����, Ŭ����, �����, �ֹ���, ��ȯ����� �ʱ�ȭ
            End If
            
            dataArr = dict(key)
            
            ' ������ �� ���� (N��: �����, O��: Ŭ����, P��: �����, R��: �ֹ���, X��: ��ȯ�����)
            dataArr(0) = dataArr(0) + CDbl(ws.Cells(cell.Row, 14).Value) ' ����� (N��)
            dataArr(1) = dataArr(1) + CDbl(ws.Cells(cell.Row, 15).Value) ' Ŭ���� (O��)
            dataArr(2) = dataArr(2) + CDbl(ws.Cells(cell.Row, 16).Value) ' ����� (P��)
            dataArr(3) = dataArr(3) + CDbl(ws.Cells(cell.Row, 18).Value) ' �ֹ��� (R��)
            dataArr(4) = dataArr(4) + CDbl(ws.Cells(cell.Row, 24).Value) ' ��ȯ����� (X��)
            
            dict(key) = dataArr
        End If
    Next cell
    
    ' ���� "Ű���� �м�" ��Ʈ ���� �� ���� ����
    On Error Resume Next
    Application.DisplayAlerts = False
    If Not wb.Sheets("Ű���� �м�") Is Nothing Then
        wb.Sheets("Ű���� �м�").Delete
    End If
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Set analysisWs = wb.Sheets.Add
    analysisWs.Name = "Ű���� �м�"
    
    ' ��� �߰� �� ���� ����
    With analysisWs.Range("A1:J1")
        .Value = Array("Ű����", "�����", "Ŭ����", "Ŭ����(%)", "�ֹ���", "��ȯ��(%)", "CPC", "�����", "�������", "ROAS(%)")
        .Font.Bold = True
        .Borders.LineStyle = xlContinuous
        .Interior.Color = RGB(220, 220, 220) ' ���� ȸ�� ���
    End With
    
    ' ������ ���
    analysisRow = 2
    Dim item As Variant
    For Each item In dict.Keys
        Dim impressions As Double, clicks As Double, cost As Double, orders As Double, revenue As Double
        
        dataArr = dict(item)
        
        impressions = dataArr(0)
        clicks = dataArr(1)
        cost = dataArr(2) * 1.1 ' VAT ���� �����
        orders = dataArr(3)
        revenue = dataArr(4)
        
        analysisWs.Cells(analysisRow, 1).Value = item
        analysisWs.Cells(analysisRow, 2).Value = impressions
        analysisWs.Cells(analysisRow, 3).Value = clicks
        
        ' Ŭ���� ��� (0���� ������ ����)
        If impressions > 0 Then
            analysisWs.Cells(analysisRow, 4).Value = (clicks / impressions) * 100 ' Ŭ����(%)
        Else
            analysisWs.Cells(analysisRow, 4).Value = 0
        End If

        analysisWs.Cells(analysisRow, 5).Value = orders
        
        ' ��ȯ�� ���
        If clicks > 0 Then
            analysisWs.Cells(analysisRow, 6).Value = (orders / clicks) * 100 ' ��ȯ��(%)
        Else
            analysisWs.Cells(analysisRow, 6).Value = 0
        End If
        
        ' CPC ���
        If clicks > 0 Then
            analysisWs.Cells(analysisRow, 7).Value = cost / clicks ' CPC
        Else
            analysisWs.Cells(analysisRow, 7).Value = 0
        End If

        analysisWs.Cells(analysisRow, 8).Value = cost ' �����(VAT ����)
        analysisWs.Cells(analysisRow, 9).Value = revenue ' �������
        
        ' ROAS ���
        If cost > 0 Then
            analysisWs.Cells(analysisRow, 10).Value = (revenue / cost) * 100 ' ROAS(%)
        Else
            analysisWs.Cells(analysisRow, 10).Value = 0
        End If
        
        analysisRow = analysisRow + 1
    Next item
    
    ' �׵θ� �� �� �ʺ� ����
    With analysisWs.Range("A1:J" & analysisRow - 1)
        .Borders.LineStyle = xlContinuous
        .Columns.AutoFit
    End With
    
    ' ���� (����� ���� ��������)
    analysisWs.Range("A1:J" & analysisRow - 1).Sort Key1:=analysisWs.Range("B1"), Order1:=xlDescending, Header:=xlYes
    
    MsgBox "Ű���� �м� �Ϸ�!", vbInformation
End Sub


