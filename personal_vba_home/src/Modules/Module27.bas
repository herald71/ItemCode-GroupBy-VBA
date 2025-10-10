Attribute VB_Name = "Module27"
Sub ���α���_��¥��_ķ���κ�_�м�()
    Dim ws As Worksheet, wsNew As Worksheet
    Dim lastRow As Long
    Dim dict As Object
    Dim i As Long
    Dim key As Variant

    ' ���� ������ ��Ʈ ���� (ActiveWorkbook ���)
    Set ws = ActiveWorkbook.Sheets("Sheet1") ' ���� ������ ��Ʈ ���� ����
    
    ' ������ �� ã��
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    
    ' ������ ������ Dictionary ����
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' ���� �����Ϳ��� ��¥ + ķ���θ� �׷�ȭ
    For i = 2 To lastRow
        Dim ��¥ As String, ķ���θ� As String
        Dim ����� As Double, Ŭ���� As Double, ����� As Double
        Dim �ֹ��� As Double, ��ȯ����� As Double
        
        ��¥ = ws.Cells(i, 1).Value ' ��¥ (A��)
        ķ���θ� = ws.Cells(i, 6).Value ' ķ���θ� (F��)
        ����� = ws.Cells(i, 14).Value ' ����� (N��)
        Ŭ���� = ws.Cells(i, 15).Value ' Ŭ���� (O��)
        ����� = ws.Cells(i, 16).Value ' ����� (P��)
        �ֹ��� = ws.Cells(i, 18).Value ' �� �ֹ���(1��) (R��)
        ��ȯ����� = ws.Cells(i, 24).Value ' �� ��ȯ�����(1��) (X��)
        
        ' Key: "��¥|ķ���θ�" �������� ����
        Dim dictKey As String
        dictKey = ��¥ & "|" & ķ���θ�
        
        ' ���� �����Ͱ� ������ �ʱ�ȭ
        If Not dict.exists(dictKey) Then
            dict.Add dictKey, Array(0, 0, 0, 0, 0) ' [�����, Ŭ����, �����, �ֹ���, ��ȯ�����]
        End If
        
        ' ���� �� ����
        Dim values As Variant
        values = dict(dictKey)
        
        values(0) = values(0) + �����
        values(1) = values(1) + Ŭ����
        values(2) = values(2) + �����
        values(3) = values(3) + �ֹ���
        values(4) = values(4) + ��ȯ�����
        
        dict(dictKey) = values
    Next i
    
    ' ���� ��Ʈ ���� �� ���ο� ��Ʈ ���� (ActiveWorkbook ����)
    On Error Resume Next
    Application.DisplayAlerts = False
    ActiveWorkbook.Sheets("��¥�� ķ���� �м�").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    ' ���ο� ��Ʈ �߰�
    Set wsNew = ActiveWorkbook.Sheets.Add
    wsNew.Name = "��¥�� ķ���� �м�"
    
    ' ��� �ۼ�
    Dim headers As Variant
    headers = Array("��¥", "ķ���θ�", "ROAS(%)", "CPC", "Ŭ����(%)", "��ȯ��(%)", "�ֹ���", "�����(VAT ����)", "�������", "��ȯ����", "���ܰ�")
    
    For i = 0 To UBound(headers)
        wsNew.Cells(1, i + 1).Value = headers(i)
        wsNew.Cells(1, i + 1).Font.Bold = True ' ��� �۾� ����
    Next i
    
    ' ������ �Է�
    Dim rowIndex As Integer
    rowIndex = 2
    
    For Each key In dict.Keys
        values = dict(key)
        
        ' Key���� ��¥�� ķ���θ� �и�
        Dim keyParts() As String
        keyParts = Split(key, "|")
        
        Dim ������ As Double, Ŭ���� As Double, ������� As Double
        Dim �ֹ��� As Double, ��ȯ������ As Double
        Dim VAT���Ա���� As Double, ROAS As Double
        Dim CPC As Double, Ŭ���� As Double, ��ȯ�� As Double
        Dim ��ȯ���� As Double, ���ܰ� As Double
        
        ������ = values(0)
        Ŭ���� = values(1)
        ������� = values(2)
        �ֹ��� = values(3)
        ��ȯ������ = values(4)
        
        ' VAT ���� ����� ���
        VAT���Ա���� = ������� * 1.1
        
        ' ��ǥ ��� (0���� ������ ��� ���� + �ݿø�)
        If VAT���Ա���� <> 0 Then
            ROAS = Round((��ȯ������ / VAT���Ա����) * 100, 2)
            CPC = Round(VAT���Ա���� / IIf(Ŭ���� = 0, 1, Ŭ����), 2)
        Else
            ROAS = 0
            CPC = 0
        End If
        
        If Ŭ���� <> 0 Then
            Ŭ���� = Round((Ŭ���� / ������) * 100, 2)
            ��ȯ�� = Round((�ֹ��� / Ŭ����) * 100, 2)
        Else
            Ŭ���� = 0
            ��ȯ�� = 0
        End If
        
        If �ֹ��� <> 0 Then
            ��ȯ���� = Round(VAT���Ա���� / �ֹ���, 2)
            ���ܰ� = Round(��ȯ������ / �ֹ���, 2)
        Else
            ��ȯ���� = 0
            ���ܰ� = 0
        End If
        
        ' ��Ʈ�� �� �Է�
        wsNew.Cells(rowIndex, 1).Value = keyParts(0) ' ��¥
        wsNew.Cells(rowIndex, 2).Value = keyParts(1) ' ķ���θ�
        wsNew.Cells(rowIndex, 3).Value = ROAS
        wsNew.Cells(rowIndex, 4).Value = CPC
        wsNew.Cells(rowIndex, 5).Value = Ŭ����
        wsNew.Cells(rowIndex, 6).Value = ��ȯ��
        wsNew.Cells(rowIndex, 7).Value = �ֹ���
        wsNew.Cells(rowIndex, 8).Value = VAT���Ա����
        wsNew.Cells(rowIndex, 9).Value = ��ȯ������
        wsNew.Cells(rowIndex, 10).Value = ��ȯ����
        wsNew.Cells(rowIndex, 11).Value = ���ܰ�
        
        rowIndex = rowIndex + 1
    Next key

    ' ���� ����
    With wsNew.Columns("C:K")
        .NumberFormat = "#,##0.00" ' ���� ���� ���� (�Ҽ��� 2�ڸ�)
        .AutoFit ' �� �ʺ� �ڵ� ����
    End With
    
    wsNew.Columns("A:B").AutoFit ' ��¥, ķ���θ� �� �ʺ� �ڵ� ����
    
    ' �Ϸ� �޽���
    MsgBox "��¥�� ķ���� �м��� �Ϸ�Ǿ����ϴ�!", vbInformation, "�Ϸ�"
End Sub


