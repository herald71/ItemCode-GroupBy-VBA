Attribute VB_Name = "URL���ϻ輼"

Sub URL��������ǥ����_�ؽ�Ʈ����()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim questionMarkPos As Long
    Dim targetColumn As String
    Dim lastRow As Long
    
    ' ���� Ȱ��ȭ�� ��ũ��Ʈ ����
    Set ws = ActiveSheet
    
    ' ����ڷκ��� �� �Է¹ޱ� (��: "A"�� ���� �Է�)
    targetColumn = InputBox("URL�� ���Ե� ���� �Է��ϼ��� (��: A)")
    
    ' �Է��� ��ȿ���� ������ ����
    If targetColumn = "" Then
        MsgBox "��ȿ�� ���� �Է��ϼ���."
        Exit Sub
    End If
    
    ' ������ �� ��ȣ ã�� (�ش� ������)
    lastRow = ws.Cells(ws.Rows.count, targetColumn).End(xlUp).Row
    
    ' ������ ���� ��� �� ���� ����
    Set rng = ws.Range(targetColumn & "1:" & targetColumn & lastRow)
    
    ' �� ���� ��ȸ�ϸ� '?' ���� �ؽ�Ʈ ����
    For Each cell In rng
        If InStr(cell.Value, "?") > 0 Then
            questionMarkPos = InStr(cell.Value, "?")
            cell.Value = Left(cell.Value, questionMarkPos - 1)
        End If
    Next cell
End Sub


