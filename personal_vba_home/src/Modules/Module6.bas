Attribute VB_Name = "Module6"
Sub ����̹�������()
    Dim pic As Picture
    ' ���� Ȱ��ȭ�� ��Ʈ�� ��� �̹��� ����
    For Each pic In ActiveSheet.Pictures
        pic.Delete
    Next pic
End Sub

