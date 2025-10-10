Attribute VB_Name = "Module19"
Public Sub ���Ų�_����Ʈ��ǰ����()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim Confirm As VbMsgBoxResult
    Dim cell As Range
    Dim img As Picture
    Dim imageUrl As String
    Dim TotalSheets As Long
    Dim SheetCount As Long
    Dim Progress As Double
    Dim wb As Workbook
    
    ' ���� Ȱ��ȭ�� ���� ���� ����
    Set wb = ActiveWorkbook
    
    ' ���� ���� Ȯ��
    Confirm = MsgBox("���Ų� ����Ʈ ��ǰ ���� �۾��� ���� �ұ��?", vbYesNo + vbQuestion, "�۾� Ȯ��")
    If Confirm = vbNo Then Exit Sub
    
    ' ȭ�� ������Ʈ �� �̺�Ʈ ��Ȱ��ȭ
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayStatusBar = True
    
    ' �� ��Ʈ �� Ȯ��
    TotalSheets = wb.Worksheets.count
    SheetCount = 0
    
    ' ���� ���� ������ ��� ��Ʈ���� �۾� ����
    For Each ws In wb.Worksheets
        SheetCount = SheetCount + 1
        
        ' ������ �� ã��
        lastRow = ws.Cells(ws.Rows.count, "G").End(xlUp).Row
        
        ' ����� ǥ�� ������Ʈ
        Progress = (SheetCount / TotalSheets) * 100
        Application.StatusBar = "�۾� �����: " & Format(Progress, "0.00") & "% �Ϸ� ��..."
        
        ' 1. G���� URL�� I���� "�ٷΰ���" �����۸�ũ�� ����
        For Each cell In ws.Range("G2:G" & lastRow)
            If cell.Value Like "http*://*" Then
                ws.Cells(cell.Row, "I").Value = "�ٷΰ���"
                ws.Hyperlinks.Add Anchor:=ws.Cells(cell.Row, "I"), Address:=cell.Value, TextToDisplay:="�ٷΰ���"
            End If
        Next cell
        
        ' 2. H���� �̹��� URL�� �����Ͽ� J���� �̹��� ����
        For Each cell In ws.Range("H2:H" & lastRow)
            imageUrl = cell.Value
            If imageUrl <> "" Then
                On Error Resume Next
                Set img = ws.Pictures.Insert(imageUrl)
                If Not img Is Nothing Then
                    With img
                        .Top = ws.Cells(cell.Row, "J").Top
                        .Left = ws.Cells(cell.Row, "J").Left
                        ws.Cells(cell.Row, "J").RowHeight = 50
                        
                        ' ������ �°� ũ�� ����
                        Dim origRatio As Double
                        origRatio = .Width / .Height
                        Dim cellRatio As Double
                        cellRatio = ws.Cells(cell.Row, "J").Width / ws.Cells(cell.Row, "J").Height
                        
                        If origRatio > cellRatio Then
                            .Width = ws.Cells(cell.Row, "J").Width
                            .Height = .Width / origRatio
                        Else
                            .Height = ws.Cells(cell.Row, "J").Height
                            .Width = .Height * origRatio
                        End If
                        
                        ' �̹��� �߾� ��ġ
                        .Top = ws.Cells(cell.Row, "J").Top + (ws.Cells(cell.Row, "J").Height - .Height) / 2
                        .Left = ws.Cells(cell.Row, "J").Left + (ws.Cells(cell.Row, "J").Width - .Width) / 2
                    End With
                End If
                On Error GoTo 0
            End If
        Next cell
        
        ' 3. G���� H�� ���� ó��
        ws.Columns("G").EntireColumn.Hidden = True
        ws.Columns("H").EntireColumn.Hidden = True
        
        ' 4. I1, J1�� �ؽ�Ʈ �߰�
        ws.Range("I1").Value = "�ٷΰ���"
        ws.Range("J1").Value = "�̹���"
        
        ' 5. �׵θ� ���� �� ����
        With ws.Range("A1:J" & lastRow)
            .Borders.LineStyle = xlContinuous
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        
        ' 6. D���� ���� ���� ���� ���� �� D1 �߾� ����
        ws.Range("D2:D" & lastRow).HorizontalAlignment = xlLeft
        ws.Range("D1").HorizontalAlignment = xlCenter
    Next ws
    
    ' �۾� ����
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "���Ų� ����Ʈ ��ǰ ���� �۾��� �Ϸ�Ǿ����ϴ�.", vbInformation, "�۾� �Ϸ�"
End Sub
