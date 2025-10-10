Attribute VB_Name = "Module25"
Sub �ܾ����������()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim colNum As Integer
    Dim searchWords As String
    Dim keywords() As String
    Dim i As Long, j As Integer
    Dim cellValue As String
    Dim deleteCount As Integer
    
    ' ���� Ȱ��ȭ�� ��Ʈ�� ������� ��
    Set ws = ActiveSheet
    
    ' ����ڿ��� �� ��ȣ �Է¹ޱ�
    colNum = Application.InputBox("�˻��� �� ��ȣ�� �Է��ϼ��� (��: 2�� B��)", Type:=1)
    If colNum < 1 Then Exit Sub ' �Է��� �߸��Ǹ� ����
    
    ' ����ڿ��� ������ �ܾ� �Է¹ޱ� (��ǥ�� ����)
    searchWords = Application.InputBox("������ �ܾ���� �Է��ϼ��� (��ǥ�� ����)", Type:=2)
    If searchWords = "" Then Exit Sub ' �Է��� ������ ����
    
    ' �ܾ� �迭�� ��ȯ
    keywords = Split(searchWords, ",")
    
    ' ������ �� ã��
    lastRow = ws.Cells(ws.Rows.count, colNum).End(xlUp).Row
    
    ' ������ �� ���� �ʱ�ȭ
    deleteCount = 0

    ' �ڿ��� ������ �����ؾ� ���� ����
    For i = lastRow To 1 Step -1
        cellValue = ws.Cells(i, colNum).Value
        
        ' �Էµ� �ܾ� �� �ϳ��� ���ԵǸ� �� ����
        For j = LBound(keywords) To UBound(keywords)
            If InStr(1, cellValue, Trim(keywords(j)), vbTextCompare) > 0 Then
                ws.Rows(i).Delete
                deleteCount = deleteCount + 1 ' ������ �� ���� ����
                Exit For ' �� �� �����Ǹ� �ش� ���� �� �̻� �˻��� �ʿ� ����
            End If
        Next j
    Next i
    
    ' ������ �� �� ���
    MsgBox deleteCount & "���� ���� �����Ǿ����ϴ�.", vbInformation, "���� �Ϸ�"
End Sub

