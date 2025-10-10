Attribute VB_Name = "Module8"
Sub �����۸�ũ�����_����ü()
    Dim sourceCol As String
    Dim targetCol As String
    Dim lastRow As Long
    Dim ws As Worksheet
    Dim i As Long
    Dim hyperlinkAddress As String
    
    ' Ȱ�� ��Ʈ ����
    Set ws = ActiveSheet
    
    ' �����۸�ũ �ּҰ� �ִ� �� ��ȣ �Է� �ޱ�
    sourceCol = InputBox("�����۸�ũ �ּҰ� �ִ� �� ��ȣ�� �Է��ϼ��� (��: A, B, C ��):", "�����۸�ũ �ҽ� �� ��ȣ �Է�")
    If sourceCol = "" Then
        MsgBox "�ҽ� �� �Է��� ��ҵǾ����ϴ�."
        Exit Sub
    End If
    
    ' "�ٷΰ���"�� ������ �� ��ȣ �Է� �ޱ�
    targetCol = InputBox("�����۸�ũ�� ������ �� ��ȣ�� �Է��ϼ��� (��: A, B, C ��):", "�����۸�ũ Ÿ�� �� ��ȣ �Է�")
    If targetCol = "" Then
        MsgBox "Ÿ�� �� �Է��� ��ҵǾ����ϴ�."
        Exit Sub
    End If
    
    ' �Էµ� �ҽ� ������ ������ �� ã��
    lastRow = ws.Cells(ws.Rows.count, sourceCol).End(xlUp).Row
    
    ' �ҽ� ���� �����۸�ũ �ּҸ� �а� Ÿ�� ���� "�ٷΰ���" ����
    For i = 1 To lastRow
        hyperlinkAddress = ws.Cells(i, sourceCol).Value
        If hyperlinkAddress Like "http*://*" Then
            ' Ÿ�� ���� ���� �࿡ "�ٷΰ���" �ؽ�Ʈ�� �����۸�ũ �߰�
            With ws.Cells(i, targetCol)
                .Value = "�ٷΰ���"
                ws.Hyperlinks.Add Anchor:=.Cells, Address:=hyperlinkAddress, TextToDisplay:="�ٷΰ���"
            End With
        End If
    Next i
    
    MsgBox "�����۸�ũ ���� �۾��� �Ϸ�Ǿ����ϴ�."
End Sub


