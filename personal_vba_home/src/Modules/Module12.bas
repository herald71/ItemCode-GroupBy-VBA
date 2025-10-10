Attribute VB_Name = "Module12"
Sub ����ϻ���()
    Dim urlColumn As String
    Dim insertColumn As String
    Dim cell As Range
    Dim img As Picture
    Dim imageUrl As String
    Dim regex As Object

    ' ���Խ��� ����Ͽ� �����ڸ� ����ϵ��� ����
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = False
    regex.Pattern = "^[A-Za-z]+$"

    ' ����ڷκ��� �̹��� URL�� �ִ� ���� �̹����� ������ �� ������ �Է¹޽��ϴ�.
    urlColumn = InputBox("�̹��� URL�� ���Ե� ���� �Է��ϼ���. ��: G")
    insertColumn = InputBox("�̹����� ������ ���� �Է��ϼ���. ��: J")

    ' �����ڸ� �ԷµǾ����� Ȯ���մϴ�.
    If Not regex.Test(urlColumn) Or Not regex.Test(insertColumn) Then
        MsgBox "�� �̸��� �ݵ�� �����ڸ� �Է��ؾ� �մϴ�.", vbExclamation
        Exit Sub
    End If

    On Error GoTo ErrorHandler ' ���� �߻� �� ������ ��ġ ����
    
    ' �Է¹��� ���� ������� URL�� �����մϴ�.
    For Each cell In Range(urlColumn & "2:" & urlColumn & Cells(Rows.count, urlColumn).End(xlUp).Row)
        imageUrl = cell.Value
        
        ' �̹��� URL�� �� ���ڿ��� �ƴ� ��쿡�� �̹��� ������ �õ��մϴ�.
        If imageUrl <> "" Then
            ' �̹����� ��ũ��Ʈ�� �����մϴ�.
            Set img = ActiveSheet.Pictures.Insert(imageUrl)
            
            With img
                ' �̹����� ����ڰ� ������ ���� ��ġ��ŵ�ϴ�.
                .Top = Cells(cell.Row, insertColumn).Top
                .Left = Cells(cell.Row, insertColumn).Left
                
                ' ���� ���̸� 50���� �����մϴ�.
                Cells(cell.Row, insertColumn).RowHeight = 50
                
                ' �̹����� ���� ���μ��� ������ ����մϴ�.
                Dim origRatio As Double
                origRatio = .Width / .Height
                
                ' ���� �� ���� ���μ��� ������ ����մϴ�.
                Dim cellRatio As Double
                cellRatio = Cells(cell.Row, insertColumn).Width / Cells(cell.Row, insertColumn).Height
                
                ' �� ������ ���� �̹����� ũ�⸦ �����մϴ�.
                If origRatio > cellRatio Then
                    .Width = Cells(cell.Row, insertColumn).Width
                    .Height = .Width / origRatio
                Else
                    .Height = Cells(cell.Row, insertColumn).Height
                    .Width = .Height * origRatio
                End If
                
                ' �̹����� ���� �߾ӿ� ��ġ�մϴ�.
                .Top = Cells(cell.Row, insertColumn).Top + (Cells(cell.Row, insertColumn).Height - .Height) / 2
                .Left = Cells(cell.Row, insertColumn).Left + (Cells(cell.Row, insertColumn).Width - .Width) / 2
            End With
        End If
    Next cell
    
    Exit Sub ' ���� ����

ErrorHandler:
    MsgBox "������ �߻��߽��ϴ�: " & Err.Description, vbCritical
End Sub

