Attribute VB_Name = "Module9"

Sub ���ؽ���ǰ���ڵ�����()
    Dim ws As Worksheet
    Set ws = Sheets("data")
    
    ' 1. "data" ��Ʈ�� 1,2���� �����մϴ�.
    ws.Rows("1:2").Delete
    
' 2. E ���� �ʵ尪�� ���� �׸��� ã�Ƽ� �� ���� ��� �����͸� �����մϴ�.
    Dim lastRow As Long, i As Long
    lastRow = ws.Cells(ws.Rows.count, "E").End(xlUp).Row

    For i = lastRow To 1 Step -1
            If Trim(ws.Cells(i, "E").Value) = "" Then
            ws.Rows(i).Delete
    End If
    Next i

    
    ' 3. J������ S������ ��� ������ �����մϴ�.
    ws.Columns("J:S").Delete Shift:=xlToLeft
    
    ' 4. G, F, D ���� ������� �����մϴ�.
    ws.Columns("G").Delete Shift:=xlToLeft
    ws.Columns("F").Delete Shift:=xlToLeft
    ws.Columns("D").Delete Shift:=xlToLeft
    
    ' 5. A, B ���� �����͸� ���� �ٲߴϴ�. �ӽ� ���� ����ϴ� ������� �����մϴ�.
    ws.Columns("A").EntireColumn.Copy
    ws.Columns("XFD").EntireColumn.PasteSpecial Paste:=xlPasteValues
    ws.Columns("B").EntireColumn.Copy
    ws.Columns("A").EntireColumn.PasteSpecial Paste:=xlPasteValues
    ws.Columns("XFD").EntireColumn.Copy
    ws.Columns("B").EntireColumn.PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    ' �ӽ÷� ���� XFD ���� �����մϴ�.
        ws.Columns("XFD").Delete
    
    ' 6. A�� �տ� �� ���� �����մϴ�.
     ws.Columns("A").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
End Sub


