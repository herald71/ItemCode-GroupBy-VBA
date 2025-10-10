Attribute VB_Name = "Module1"
Sub ����ȭ�Ͻ�Ʈ������()
    Dim folderPath As String
    Dim FileName As String
    Dim wbSource As Workbook
    Dim wsSource As Worksheet
    Dim wbDest As Workbook
    Dim wsDest As Worksheet
    Dim filePath As String
    Dim lastRow As Long
    Dim FileTitle As String
    Dim fd As FileDialog
    
    ' ����ڿ��� ���� ���� ��ȭ���� ǥ��
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.Title = "������ ���ϵ��� �ִ� ������ �����ϼ���."
    
    ' ����ڰ� ������ ������ ���
    If fd.Show = -1 Then
        folderPath = fd.SelectedItems(1) & "\" ' ���õ� ���� ��θ� ������
    Else
        MsgBox "������ �������� �ʾҽ��ϴ�. �۾��� ����մϴ�."
        Exit Sub
    End If
    
    ' ���յ� �����͸� ������ �� ��ũ�� ����
    Set wbDest = Workbooks.Add
    
    ' ���� �� ù ��° ���� ���� ã��
    FileName = Dir(folderPath & "*.xls*") ' ���� ����(.xls, .xlsx, .xlsm) Ȯ����
    
    ' ���� ���� �� ���� ���Ͽ� ���� �ݺ�
    Do While FileName <> ""
        ' �ҽ� ���� ����
        filePath = folderPath & FileName
        Set wbSource = Workbooks.Open(filePath)
        
        ' �� ������ ù ��° ��Ʈ ��������
        For Each wsSource In wbSource.Sheets
            ' �� ��Ʈ �߰�
            Set wsDest = wbDest.Sheets.Add(After:=wbDest.Sheets(wbDest.Sheets.count))
            ' �ҽ� ��Ʈ�� ��� �����͸� ����
            wsSource.UsedRange.Copy wsDest.Cells(1, 1)
            
            ' ��Ʈ �̸��� ���� �̸����� ���� (Ȯ���� ����)
            FileTitle = Left(FileName, InStrRev(FileName, ".") - 1)
            On Error Resume Next ' ��Ʈ �̸��� �ߺ��Ǹ� ���� ����
            wsDest.Name = FileTitle
            On Error GoTo 0
        Next wsSource
        
        ' �ҽ� ���� �ݱ�
        wbSource.Close False
        
        ' ���� ���Ϸ� �̵�
        FileName = Dir
    Loop
    
    ' ���յ� ���� ����
    Application.DisplayAlerts = False
    wbDest.Sheets(1).Delete ' �⺻���� ������ �� ��Ʈ ����
    Application.DisplayAlerts = True
    
    MsgBox "��� ������ ���������� ���յǾ����ϴ�!"
End Sub

