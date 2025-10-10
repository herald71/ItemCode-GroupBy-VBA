Attribute VB_Name = "Module29"

Sub ���α�����ȯ��ǰ�м�()
    Dim ws As Worksheet, analysisWs As Worksheet
    Dim lastRow As Long, analysisRow As Long
    Dim dict As Object
    Dim key As String
    Dim rng As Range, cell As Range
    Dim wb As Workbook
    
    ' ���� Ȱ��ȭ�� ������ �������� ���� (ThisWorkbook �� ActiveWorkbook ����)
    Set wb = ActiveWorkbook
    
    ' Sheet1�� �ƴ� ����ڰ� ������ ��Ʈ�� �������� �����ϵ��� ����
    On Error Resume Next
    Set ws = wb.ActiveSheet
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "Ȱ�� ��Ʈ�� �����ϴ�. ������ �� �� �ٽ� �õ��ϼ���.", vbExclamation
        Exit Sub
    End If
    
    ' Dictionary ���� (Ű: ��ǰ�� + �ɼ�ID, ��: Ű����, �������, �ֹ���)
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' ������ ���� ����
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    Set rng = ws.Range("J2:J" & lastRow) ' "������ȯ����߻� ��ǰ��"�� J���� �ִٰ� ����
    
    ' ������ ���� �� ����
    For Each cell In rng
        key = ws.Cells(cell.Row, 10).Value & "|" & ws.Cells(cell.Row, 11).Value ' ��ǰ��(J��) + �ɼ�ID(K��)
        
        If Not dict.exists(key) Then
            dict.Add key, Array("", 0#, 0#) ' Ű���� �ʱ�ȭ (���ڿ�), ������� (Double), �ֹ��� (Double)
        End If
        
        Dim dataArr As Variant
        dataArr = dict(key)
        
        ' Ű���� (ù ��° ���� ����)
        If dataArr(0) = "" Then
            dataArr(0) = ws.Cells(cell.Row, 13).Value ' Ű���� (M��)
        End If
        
        ' ������� ���� (X��)
        dataArr(1) = dataArr(1) + CDbl(Nz(ws.Cells(cell.Row, 24).Value, 0)) ' ������� (X��)
        
        ' �ֹ��� ���� (R��)
        dataArr(2) = dataArr(2) + CDbl(Nz(ws.Cells(cell.Row, 18).Value, 0)) ' �ֹ��� (R��)
        
        dict(key) = dataArr
    Next cell
    
    ' ���� �м� ��Ʈ ���� �� ���� ����
    On Error Resume Next
    Application.DisplayAlerts = False
    If Not wb.Sheets("������ȯ �м�") Is Nothing Then wb.Sheets("������ȯ �м�").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    ' �� ��Ʈ ����
    Set analysisWs = wb.Sheets.Add
    analysisWs.Name = "������ȯ �м�"
    
    ' ��� �߰� �� ���� ����
    With analysisWs.Range("A1:E1")
        .Value = Array("������ȯ����߻� ��ǰ��", "������ȯ����߻� �ɼ�ID", "Ű����", "�������", "�ֹ���")
        .Font.Bold = True
        .Borders.LineStyle = xlContinuous
        .Interior.Color = RGB(220, 220, 220) ' ���� ȸ�� ���
    End With
    
    ' ������ ���
    analysisRow = 2
    Dim item As Variant
    For Each item In dict.Keys
        Dim splitKey As Variant
        Dim productName As String, optionID As String, keyword As String
        Dim revenue As Double, orders As Double
        
        splitKey = Split(item, "|")
        productName = splitKey(0)
        optionID = splitKey(1)
        keyword = dict(item)(0)
        revenue = dict(item)(1)
        orders = dict(item)(2)
        
        analysisWs.Cells(analysisRow, 1).Value = productName
        analysisWs.Cells(analysisRow, 2).Value = optionID
        analysisWs.Cells(analysisRow, 3).Value = keyword
        analysisWs.Cells(analysisRow, 4).Value = revenue
        analysisWs.Cells(analysisRow, 5).Value = orders
        
        analysisRow = analysisRow + 1
    Next item
    
    ' �׵θ� �� �� �ʺ� ����
    With analysisWs.Range("A1:E" & analysisRow - 1)
        .Borders.LineStyle = xlContinuous
        .Columns.AutoFit
    End With
    
    ' ���� (������� ���� ��������)
    analysisWs.Range("A1:E" & analysisRow - 1).Sort _
        Key1:=analysisWs.Range("D1"), Order1:=xlDescending, Header:=xlYes
    
    MsgBox "������ȯ �м� �Ϸ�!", vbInformation
End Sub

Function Nz(Value As Variant, Default As Double) As Double
    If IsNumeric(Value) Then
        Nz = CDbl(Value)
    Else
        Nz = Default
    End If
End Function




