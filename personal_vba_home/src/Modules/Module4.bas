Attribute VB_Name = "Module4"
Sub �ؽ�Ʈ���ڸ����ڷκ�ȯ_�ܰ�()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim columnLetter As String
    
    ' ���� Ȱ�� ��Ʈ ����
    Set ws = ActiveSheet
    
    ' ����ڷκ��� �� ���ڸ� �Է¹���
    columnLetter = InputBox("��ȯ�� ���� ���ڸ� �Է��ϼ��� (��: A, B, C, ...)", "�� ����")
    
    ' �Է¹��� ���� ������ ���� ���� (���� �� ����)
    Set rng = ws.Range(columnLetter & "2:" & columnLetter & ws.Cells(ws.Rows.count, columnLetter).End(xlUp).Row)
    
    ' ���� ���� �� ���� ���� �ؽ�Ʈ�� ���ڷ� ��ȯ
    For Each cell In rng
        ' �� ���� ���ڷ� ��ȯ �������� Ȯ�� �� ��ȯ
        If IsNumeric(cell.Value) Then
            cell.Value = val(cell.Value)
        End If
    Next cell
    
    MsgBox "�ؽ�Ʈ ���ڰ� ���ڷ� ��ȯ�Ǿ����ϴ�.", vbInformation
End Sub
