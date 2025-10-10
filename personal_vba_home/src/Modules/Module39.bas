Attribute VB_Name = "Module39"
Sub �����󵵼�_����_���_�׵θ�_�׷���()
    Dim ws As Worksheet
    Dim summaryWs As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim monthVal As Variant
    Dim monthCount(1 To 12) As Long
    Dim rankVal As Integer
    Dim lastDataRow As Long
    Dim chartObj As ChartObject

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.count, "N").End(xlUp).Row

    ' ���� �󵵼� ����
    For i = 2 To lastRow
        monthVal = ws.Cells(i, "N").Value
        If IsNumeric(monthVal) Then
            If monthVal >= 1 And monthVal <= 12 Then
                monthCount(monthVal) = monthCount(monthVal) + 1
            End If
        End If
    Next i

    ' ���� ��Ʈ ����
    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets("���� �󵵼�").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    ' �� ��Ʈ ����
    Set summaryWs = Worksheets.Add
    summaryWs.Name = "���� �󵵼�"

    ' ��� �ۼ�
    With summaryWs
        .Range("A1").Value = "��"
        .Range("B1").Value = "�󵵼�"
        .Range("C1").Value = "����"
        .Range("D1").Value = "���"
        .Range("A1:D1").Font.Bold = True
    End With

    ' ������ �Է�
    For i = 1 To 12
        summaryWs.Cells(i + 1, 1).Value = i & "��"
        summaryWs.Cells(i + 1, 2).Value = monthCount(i)
    Next i

    ' �󵵼� ����
    summaryWs.Range("B2:B13").NumberFormat = "#,##0"

    ' ���� ���
    summaryWs.Range("C2:C13").FormulaR1C1 = "=RANK(RC[-1], R2C2:R13C2, 0)"

    ' ��� �� ���� ����
    For i = 2 To 13
        rankVal = summaryWs.Cells(i, 3).Value
        Select Case rankVal
            Case 1 To 3
                summaryWs.Range("A" & i & ":D" & i).Interior.Color = RGB(255, 204, 0) ' ������
                summaryWs.Cells(i, 4).Value = "������"
            Case 10 To 12
                summaryWs.Range("A" & i & ":D" & i).Interior.Color = RGB(204, 255, 255) ' �����
                summaryWs.Cells(i, 4).Value = "�����"
        End Select
    Next i

    ' ���� ������ ���� ���
    lastDataRow = summaryWs.Cells(summaryWs.Rows.count, "A").End(xlUp).Row

    ' �׵θ� ����
    With summaryWs.Range("A1:D" & lastDataRow)
        .Columns.AutoFit
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
    End With

    ' ? ��Ʈ ���� (��������)
    Set chartObj = summaryWs.ChartObjects.Add(Left:=300, Width:=500, Top:=30, Height:=300)
    With chartObj.Chart
        .ChartType = xlLine
        .SetSourceData Source:=summaryWs.Range("A1:B13")
        .HasTitle = True
        .ChartTitle.text = "�󵵼�"
        .Axes(xlCategory).HasTitle = False
        .Axes(xlValue).HasTitle = False
    End With

    MsgBox "? ��� �۾� �Ϸ�: ��Ʈ + ������/����� + �������� ��Ʈ���� �����Ǿ����ϴ�!", vbInformation
End Sub

