Attribute VB_Name = "Module17"
Sub MergeExcelFiles_Onesheet()
    Dim MyPath As String, FilesInPath As String
    Dim MyFiles() As String
    Dim SourceRcount As Long, FNum As Long
    Dim mybook As Workbook, BaseWks As Worksheet
    Dim sourceRange As Range, destrange As Range
    Dim rnum As Long, CalcMode As Long
    Dim fd As FileDialog ' ���� ��ȭ���� ���� ����

    ' ���� ���� ��ȭ���ڸ� ����Ͽ� ����ڿ��� ���� ������ ��û
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.Title = "������ �����ϼ���"

    ' ����ڰ� ������ �����ϸ� ��θ� MyPath�� ����
    If fd.Show = -1 Then
        MyPath = fd.SelectedItems(1) ' ���õ� ���� ���
    Else
        MsgBox "������ ���õ��� �ʾҽ��ϴ�. ��ũ�θ� �����մϴ�."
        Exit Sub
    End If

    ' ���� ��� ���� "\"�� ������ ������ "\" �߰�
    If Right(MyPath, 1) <> "\" Then MyPath = MyPath & "\"

    ' ���� ��ο� ��ġ�ϴ� ���� ����� �����ɴϴ�.
    FilesInPath = Dir(MyPath & "*.xls*")
    If FilesInPath = "" Then
        MsgBox "�ش� �������� ��ĥ Excel ������ �����ϴ�."
        Exit Sub
    End If

    ' ���� �迭�� ���� �̸��� ����
    FNum = 0
    Do While FilesInPath <> ""
        FNum = FNum + 1
        ReDim Preserve MyFiles(1 To FNum)
        MyFiles(FNum) = FilesInPath
        FilesInPath = Dir()
    Loop

    ' Ȱ�� ��ũ���� ù ��° ��Ʈ�� ���
    Set BaseWks = ActiveWorkbook.Sheets(1)
    rnum = 1

    ' �� ������ ���� ������ ���� ��Ĩ�ϴ�.
    For FNum = 1 To UBound(MyFiles)
        Set mybook = Workbooks.Open(MyPath & MyFiles(FNum))

        ' ù ��° ���Ͽ����� ��� �����͸� �����ϰ�, �� ���� ���Ͽ����� ù ��° ���� �����մϴ�.
        With mybook.Sheets(1)
            Set sourceRange = .UsedRange
            If FNum > 1 Then
                Set sourceRange = sourceRange.Offset(1, 0).Resize(sourceRange.Rows.count - 1, sourceRange.Columns.count)
            End If

            ' �����͸� �⺻ ��ũ��Ʈ�� ����
            If rnum + sourceRange.Rows.count > BaseWks.Rows.count Then
                MsgBox "����� Excel ��Ʈ�� �ִ� �� ���� �ʰ��մϴ�."
                mybook.Close SaveChanges:=False
                GoTo ExitTheSub
            Else
                Set destrange = BaseWks.Cells(rnum, "A")
                sourceRange.Copy Destination:=destrange
                rnum = rnum + sourceRange.Rows.count
            End If
        End With

        mybook.Close SaveChanges:=False
    Next FNum

ExitTheSub:
End Sub

