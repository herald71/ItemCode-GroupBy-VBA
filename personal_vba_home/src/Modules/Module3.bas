Attribute VB_Name = "Module3"
Sub �������Է¹޾Ƽ��ڷ���ȯ()
Attribute �������Է¹޾Ƽ��ڷ���ȯ.VB_Description = "������ �Է¹޾� ���ڸ� ����"
Attribute �������Է¹޾Ƽ��ڷ���ȯ.VB_ProcData.VB_Invoke_Func = "t\n14"
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Dim cell As Range
    Dim arrCols As Variant
    Dim col As Variant
    Dim i As Long
    Dim cleanValue As String
    Dim colInput As String
    
    ' ����ڷκ��� �˻��� ���� �Է� �޽��ϴ�.
    colInput = InputBox("�˻��� ���� �Է��ϼ���. ���� ��ǥ(,)�� �����ϼ���. ��: D,E,J,M")
    If colInput = "" Then Exit Sub ' ����ڰ� �Է��� ����ϸ� �ڵ带 �����մϴ�.
    
    ' �Է¹��� ���� �迭�� ��ȯ�մϴ�.
    arrCols = Split(colInput, ",")
    
    For Each col In arrCols
        For i = 2 To ws.Cells(ws.Rows.count, col).End(xlUp).Row
            Set cell = ws.Cells(i, col)
            cleanValue = ExtractNumbers(cell.Value)
            If cleanValue <> "" Then
                cell.Value = CDbl(cleanValue)
            Else
                cell.ClearContents
            End If
        Next i
    Next col
End Sub

Function ExtractNumbers(str As String) As String
    Dim output As String
    Dim pos As Integer

    For pos = 1 To Len(str)
        If Mid(str, pos, 1) Like "[0-9]" Then
            output = output & Mid(str, pos, 1)
        End If
    Next pos

    ExtractNumbers = output
End Function

