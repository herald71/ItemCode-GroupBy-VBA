Attribute VB_Name = "Module28"
Sub ���α����������м�()
    Dim ws As Worksheet, analysisWs As Worksheet
    Dim lastRow As Long, analysisRow As Long
    Dim dict As Object
    Dim key As String
    Dim rng As Range, cell As Range

    ' ���� ��Ʈ ���� (Ȱ��ȭ�� ��ũ���� Sheet1)
    Set ws = ActiveWorkbook.Sheets("Sheet1")
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' ������ ���� ����
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    Set rng = ws.Range("L2:L" & lastRow) ' "���� ���� ����"�� L���� �ִٰ� ����
    
    ' ������ ���� �� ����
    For Each cell In rng
        key = ws.Cells(cell.Row, 6).Value & "|" & cell.Value ' ķ���θ�(F��)�� ���� ���� ����(L��) ����
        
        If Not dict.exists(key) Then
            dict.Add key, Array(0, 0, 0, 0, 0, 0) ' �����, Ŭ����, �����, �ֹ���, ��ȯ����� �ʱ�ȭ
        End If
        
        Dim dataArr As Variant
        dataArr = dict(key)
        
        ' ������ �� ���� (N��: �����, O��: Ŭ����, P��: �����, R��: �ֹ���, X��: ��ȯ�����)
        dataArr(0) = dataArr(0) + val(ws.Cells(cell.Row, 14).Value) ' �����
        dataArr(1) = dataArr(1) + val(ws.Cells(cell.Row, 15).Value) ' Ŭ����
        dataArr(2) = dataArr(2) + val(ws.Cells(cell.Row, 16).Value) ' �����
        dataArr(3) = dataArr(3) + val(ws.Cells(cell.Row, 18).Value) ' �ֹ���
        dataArr(4) = dataArr(4) + val(ws.Cells(cell.Row, 24).Value) ' ��ȯ�����
        
        dict(key) = dataArr
    Next cell
    
    ' ���� ��Ʈ ���� �� ���� ����
    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets("����������� �м�").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Set analysisWs = ActiveWorkbook.Sheets.Add
    analysisWs.Name = "����������� �м�"
    
    ' ��� �߰� �� ���� ����
    With analysisWs.Range("A1:N1")
        .Value = Array("ķ���θ�", "���� ���� ����", "�����", "Ŭ����", "�ֹ���", "Ŭ����(%)", "��ȯ��(%)", "CPM", "CPC", "�����", "�������", "ROAS(%)", "��ȯ����", "���ܰ�")
        .Font.Bold = True
        .Borders.LineStyle = xlContinuous
        .Interior.Color = RGB(220, 220, 220) ' ���� ȸ�� ���
    End With
    
    ' ������ ���
    analysisRow = 2
    Dim item As Variant
    For Each item In dict.Keys
        Dim splitKey As Variant
        Dim campaign As String, exposure As String
        Dim impressions As Double, clicks As Double, cost As Double, orders As Double, revenue As Double
        
        splitKey = Split(item, "|")
        campaign = splitKey(0)
        exposure = splitKey(1)
        
        impressions = dict(item)(0)
        clicks = dict(item)(1)
        cost = dict(item)(2) * 1.1 ' VAT ���� �����
        orders = dict(item)(3)
        revenue = dict(item)(4)
        
        analysisWs.Cells(analysisRow, 1).Value = campaign
        analysisWs.Cells(analysisRow, 2).Value = exposure
        analysisWs.Cells(analysisRow, 3).Value = impressions
        analysisWs.Cells(analysisRow, 4).Value = clicks
        analysisWs.Cells(analysisRow, 5).Value = orders
        analysisWs.Cells(analysisRow, 6).Value = IIf(impressions > 0, (clicks / impressions) * 100, 0) ' Ŭ����
        analysisWs.Cells(analysisRow, 7).Value = IIf(clicks > 0, (orders / clicks) * 100, 0) ' ��ȯ��
        analysisWs.Cells(analysisRow, 8).Value = IIf(impressions > 0, (cost / impressions) * 1000, 0) ' CPM
        analysisWs.Cells(analysisRow, 9).Value = IIf(clicks > 0, cost / clicks, 0) ' CPC
        analysisWs.Cells(analysisRow, 10).Value = cost ' �����(VAT ����)
        analysisWs.Cells(analysisRow, 11).Value = revenue ' �������
        analysisWs.Cells(analysisRow, 12).Value = IIf(cost > 0, (revenue / cost) * 100, 0) ' ROAS
        analysisWs.Cells(analysisRow, 13).Value = IIf(orders > 0, cost / orders, 0) ' ��ȯ����
        analysisWs.Cells(analysisRow, 14).Value = IIf(orders > 0, revenue / orders, 0) ' ���ܰ�
        
        analysisRow = analysisRow + 1
    Next item
    
    ' �׵θ� �� �� �ʺ� ����
    With analysisWs.Range("A1:N" & analysisRow - 1)
        .Borders.LineStyle = xlContinuous
        .Columns.AutoFit
    End With
    
    ' ���� (����� ���� ��������)
    analysisWs.Range("A1:N" & analysisRow - 1).Sort Key1:=analysisWs.Range("C1"), Order1:=xlDescending, Header:=xlYes
    
    MsgBox "���� ���� ���� �м� �Ϸ�!", vbInformation
End Sub


