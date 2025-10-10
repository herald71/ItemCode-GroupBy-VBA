Attribute VB_Name = "Module13"
Sub ����Ÿ���ʺ��ڵ��������׵θ�ġ��()
    ' ������ �� �ʺ� �ڵ� ����, �׵θ� �߰�, ������ ���� ����

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim headerRange As Range

    ' ���� Ȱ��ȭ�� ��ũ��Ʈ�� �����ɴϴ�.
    Set ws = ActiveSheet

    ' ������ ������ ������ ��� ���� ã���ϴ�.
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).column

    ' ��� ���� �ʺ� �ڵ����� �����մϴ�.
    ws.Columns("A:" & Split(Cells(, lastCol).Address, "$")(1)).AutoFit

    ' ������ ������ �׵θ��� �߰��մϴ�.
    With ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Borders
        .LineStyle = xlContinuous
        .colorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With

    ' ������ ���� ����
    Set headerRange = ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))
    With headerRange
        .Font.Bold = True
        .Font.Size = 13
        .HorizontalAlignment = xlCenter       ' ��� ����
        .VerticalAlignment = xlCenter         ' ��� ����
        .WrapText = True                      ' �ڵ� �ٹٲ�
        .Interior.Color = RGB(189, 215, 238)  ' �Ľ��� ���
    End With
End Sub



