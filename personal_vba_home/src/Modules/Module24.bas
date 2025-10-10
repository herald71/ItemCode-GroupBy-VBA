Attribute VB_Name = "Module24"
Sub ���α��������ǰ�м�()
    Dim ws As Worksheet, analysisWs As Worksheet
    Dim lastRow As Long, analysisRow As Long
    Dim dict As Object
    Dim key As String
    Dim rng As Range, cell As Range
    Dim wb As Workbook

    ' ���� Ȱ��ȭ�� ��ũ���� ����
    Set wb = ActiveWorkbook
    Set ws = wb.Sheets("Sheet1") ' ���� �����Ͱ� ��� �ִ� ��Ʈ

    Set dict = CreateObject("Scripting.Dictionary")

    ' ������ ���� ����
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    Set rng = ws.Range("H2:H" & lastRow) ' "�������� ��ǰ��"�� H���� �ִٰ� ����

    ' ������ ���� �� ����
    For Each cell In rng
        key = ws.Cells(cell.Row, 8).Value & "|" & ws.Cells(cell.Row, 9).Value

        If Not dict.exists(key) Then
            dict.Add key, Array(0#, 0#, 0#, 0#, 0#) ' �ֹ���, �����, �������, �����, Ŭ���� �ʱ�ȭ
        End If

        Dim dataArr As Variant
        dataArr = dict(key)

        ' ������ �� ���� (Nz �Լ� ����)
        dataArr(0) = dataArr(0) + Nz(ws.Cells(cell.Row, 18).Value, 0) ' �ֹ��� (R��)
        dataArr(1) = dataArr(1) + Nz(ws.Cells(cell.Row, 16).Value, 0) ' ����� (P��)
        dataArr(2) = dataArr(2) + Nz(ws.Cells(cell.Row, 24).Value, 0) ' ������� (X��)
        dataArr(3) = dataArr(3) + Nz(ws.Cells(cell.Row, 14).Value, 0) ' ����� (N��)
        dataArr(4) = dataArr(4) + Nz(ws.Cells(cell.Row, 15).Value, 0) ' Ŭ���� (O��)

        dict(key) = dataArr
    Next cell

    ' ���� �м� ��Ʈ ���� �� ���� ����
    On Error Resume Next
    Application.DisplayAlerts = False
    wb.Sheets("�������� ��ǰ�м�").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set analysisWs = wb.Sheets.Add
    analysisWs.Name = "�������� ��ǰ�м�"

    ' ��� �߰�
    With analysisWs.Range("A1:J1")
        .Value = Array("�������� ��ǰ��", "�������� �ɼ�ID", "�ֹ���", "�����", "�������", "ROAS(%)", "�����", "Ŭ����", "Ŭ����(%)", "��ȯ��(%)")
        .Font.Bold = True
        .Borders.LineStyle = xlContinuous
        .Interior.Color = RGB(220, 220, 220)
    End With

    ' ������ ���
    analysisRow = 2
    Dim item As Variant
    For Each item In dict.Keys
        Dim splitKey As Variant
        Dim productName As String, optionID As String
        Dim orders As Double, cost As Double, revenue As Double, impressions As Double, clicks As Double

        splitKey = Split(item, "|")
        productName = splitKey(0)
        optionID = splitKey(1)

        orders = dict(item)(0)
        cost = dict(item)(1) * 1.1 ' VAT ���� �����
        revenue = dict(item)(2)
        impressions = dict(item)(3)
        clicks = dict(item)(4)

        analysisWs.Cells(analysisRow, 1).Value = productName
        analysisWs.Cells(analysisRow, 2).Value = optionID
        analysisWs.Cells(analysisRow, 3).Value = orders
        analysisWs.Cells(analysisRow, 4).Value = cost
        analysisWs.Cells(analysisRow, 5).Value = revenue

        ' ROAS ���
        If cost > 0 Then
            analysisWs.Cells(analysisRow, 6).Value = Round((revenue / cost) * 100, 2)
        Else
            analysisWs.Cells(analysisRow, 6).Value = 0
        End If

        analysisWs.Cells(analysisRow, 7).Value = impressions
        analysisWs.Cells(analysisRow, 8).Value = clicks

        ' Ŭ���� ���
        If impressions > 0 Then
            analysisWs.Cells(analysisRow, 9).Value = Round((clicks / impressions) * 100, 2)
        Else
            analysisWs.Cells(analysisRow, 9).Value = 0
        End If

        ' ��ȯ�� ���
        If clicks > 0 Then
            analysisWs.Cells(analysisRow, 10).Value = Round((orders / clicks) * 100, 2)
        Else
            analysisWs.Cells(analysisRow, 10).Value = 0
        End If

        analysisRow = analysisRow + 1
    Next item

    MsgBox "�������� ��ǰ�м� �Ϸ�!", vbInformation
End Sub


Function Nz(Value, Default As Double) As Double
    If Not IsNumeric(Value) Or IsError(Value) Or IsEmpty(Value) Then
        Nz = Default
    Else
        Nz = CDbl(Value)
    End If
End Function


