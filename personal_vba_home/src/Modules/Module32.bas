Attribute VB_Name = "Module32"
Sub ColorRowsByColumnGroup()
    ' ---------------------------------------------------------
    ' ���α׷��� : ColorRowsByColumnGroup
    ' �ѱۼ���   : ������ ���� ���� �������� �� ������ �׷캰�� ����
    ' �ۼ�����   : 2025-06-28
    ' ����       : ���� Ȱ��ȭ�� ��ũ��Ʈ���� ����ڰ� �Է��� ���� ���� ��������
    '              ���� ���� ���� ��鿡 ���� �Ľ����� ������ ���������� �����Ͽ�
    '              �ð������� �׷��� ������ �� �ֵ��� �����ִ� ��ũ���Դϴ�.
    ' ---------------------------------------------------------

    Dim ws As Worksheet
    Set ws = ActiveSheet  ' ���� Ȱ��ȭ�� ��Ʈ �������� ����

    Dim colInput As Variant
    Dim colNumber As Long

    ' ����ڿ��� �� ���� �Է� ��û (��: F)
    colInput = Application.InputBox( _
        prompt:="������ �� ���� �Է��ϼ��� (��: F)", _
        Title:="���� ���� ���� �� �Է�", Type:=2)

    If colInput = False Then Exit Sub ' ����ڰ� ������� ��� ����

    ' �� ���ڸ� �� ��ȣ�� ��ȯ
    On Error GoTo InvalidColumn
    colNumber = Range(colInput & "1").column
    On Error GoTo 0

    ' ���� ���� ������ ������ �� ��ȣ ã��
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, colNumber).End(xlUp).Row

    ' Dictionary ��ü ���� (���� ���� �����)
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim i As Long, key As String
    Dim colorList As Variant
    Dim colorIndex As Long

    ' ���� �Ľ����� RGB ���� ����Ʈ
    colorList = Array( _
        RGB(255, 235, 238), RGB(232, 245, 233), RGB(232, 234, 246), _
        RGB(255, 249, 196), RGB(224, 242, 241), RGB(241, 248, 233), _
        RGB(248, 237, 227), RGB(237, 231, 246), RGB(255, 236, 179), _
        RGB(225, 245, 254), RGB(240, 244, 195), RGB(255, 245, 238) _
    )

    colorIndex = 0

    ' 2����� ������ ����� �ݺ��ϸ� ���� ����
    For i = 2 To lastRow
        key = ws.Cells(i, colNumber).Value

        If Not dict.exists(key) Then
            dict.Add key, colorList(colorIndex Mod UBound(colorList) + 1)
            colorIndex = colorIndex + 1
        End If

        ws.Rows(i).Interior.Color = dict(key)
    Next i

    Exit Sub

InvalidColumn:
    MsgBox "�Է��� ���� ��ȿ���� �ʽ��ϴ�. �ٽ� Ȯ���� �ּ���.", vbExclamation
End Sub


