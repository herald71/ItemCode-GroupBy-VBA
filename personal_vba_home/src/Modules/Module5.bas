Attribute VB_Name = "Module5"
Sub ConvertTextPercentToNumbers()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim columnLetters As String
    Dim arrColumns As Variant
    Dim column As Variant
    Dim lastRow As Long
    
    ' ���� Ȱ�� ��Ʈ ����
    Set ws = ActiveSheet
    
    ' ����ڷκ��� ��ǥ�� ���е� �� ���ڸ� �Է¹���
    columnLetters = InputBox("��ȯ�� ���� ���ڸ� ��ǥ�� �����Ͽ� �Է��ϼ��� (��: A,B,C)", "�� ����")
    
    ' �Է¹��� �� ���ڸ� ��ǥ�� �и��Ͽ� �迭�� ����
    arrColumns = Split(columnLetters, ",")
    
    ' �迭�� �� ���� ���� �۾� ����
    For Each column In arrColumns
        column = Trim(column) ' �յ� ���� ����
        If ws.Columns(column).column <= ws.Columns.count Then
            lastRow = ws.Cells(ws.Rows.count, column).End(xlUp).Row
            
            ' ������ ���� ���� ��(1��)���� Ŭ ��쿡�� �۾� ����
            If lastRow > 1 Then
                Set rng = ws.Range(column & "2:" & column & lastRow)
                
                ' ���� ���� �� ���� ���� �۾� ����
                For Each cell In rng
                    If InStr(cell.Value, "%") > 0 Then
                        ' �ؽ�Ʈ ������ ����� ���� ���� ���ڷ� ��ȯ
                        cell.Value = CDbl(Replace(cell.Value, "%", "")) / 100
                        ' �� ������ ������� ����
                        cell.NumberFormat = "0.00%"
                    ElseIf IsNumeric(cell.Value) Then
                        ' �� ���� ���ڷ� ��ȯ ������ ��� ��ȯ
                        cell.Value = val(cell.Value)
                    End If
                Next cell
            End If
        Else
            MsgBox column & "�� ��ȿ���� ���� �� �����Դϴ�.", vbCritical
            Exit Sub ' ��ȿ���� ���� �Է¿� ���� �ڵ� ���� ����
        End If
    Next column
    
    MsgBox "������ ���� �ؽ�Ʈ ������ ����� �� ���ڰ� ���� �������� ��ȯ�Ǿ����ϴ�.", vbInformation
End Sub

