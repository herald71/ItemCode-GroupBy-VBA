Attribute VB_Name = "Module10"
Sub ����üó���ϱ�()
    Dim ws As Worksheet
    Dim cell As Range
    Dim startPos As Integer
    Dim textLength As Integer
    Dim searchText As String
    Dim boldText As String

    ' ���� Ȱ��ȭ�� ��Ʈ ����
    Set ws = ActiveSheet

    ' ã�� �ؽ�Ʈ�� ���� ó���� �ؽ�Ʈ ����
    searchText = "<b>"
    boldText = "</b>"

    ' ��Ʈ�� ��� ���� ��ȸ�ϸ� �˻�
    For Each cell In ws.UsedRange
        If InStr(cell.Value, searchText) > 0 And InStr(cell.Value, boldText) > 0 Then
            ' �ؽ�Ʈ ���۰� �� ��ġ ���
            startPos = InStr(cell.Value, searchText) + Len(searchText)
            textLength = InStr(cell.Value, boldText) - startPos
            
            ' �ؽ�Ʈ���� <b>�� </b> ����
            cell.Value = Replace(cell.Value, searchText, "")
            cell.Value = Replace(cell.Value, boldText, "")

            ' �ؽ�Ʈ�� ���� ó��
            With cell.Characters(Start:=startPos - Len(searchText), Length:=textLength)
                .Font.Bold = True
            End With
        End If
    Next cell
End Sub

