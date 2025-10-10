Attribute VB_Name = "Module11"
Sub �ܾ�˻��Ļ������ΰ����ϱ�()
    Dim rng As Range
    Dim cell As Range
    Dim searchText As String
    Dim colorText As String
    Dim startPos As Integer
    Dim count As Long ' �˻��� �ܾ��� ���� ���� ���� ����
    Dim searchTextLength As Long
    Dim fontColor As Long

    ' ����ڷκ��� �˻��� �ܾ�� ������ �Է¹���
    searchText = InputBox("���� ����ü�� ǥ���� �ܾ �Է��ϼ���:", "�˻��� �Է�")
    colorText = InputBox("���� ������ �Է��ϼ��� (��: Red �Ǵ� 255,0,0):", "���� �Է�")
    
    ' �Է°��� ������� ���� ��쿡�� ����
    If searchText <> "" And colorText <> "" Then
        searchText = LCase(searchText) ' ����� �Է��� �ҹ��ڷ� ��ȯ
        searchTextLength = Len(searchText) ' �˻� �ؽ�Ʈ�� ���� ���
        fontColor = ColorToRGB(colorText) ' �Է¹��� ������ RGB ������ ��ȯ
        Set rng = ActiveSheet.UsedRange ' Ȱ�� ��Ʈ�� ���� ������ ����
        count = 0 ' ī���� �ʱ�ȭ
        
        For Each cell In rng
            If Not IsError(cell.Value) And Not IsEmpty(cell.Value) Then
                If VarType(cell.Value) = vbString Then ' �� ���� ���ڿ��� ��쿡�� ó��
                    If InStr(1, LCase(cell.Value), searchText, vbTextCompare) > 0 Then
                        startPos = InStr(1, LCase(cell.Value), searchText, vbTextCompare)
                        While startPos > 0
                            With cell.Characters(startPos, searchTextLength).Font
                                .Bold = True
                                .Color = fontColor ' ����ڰ� ������ �������� ����
                            End With
                            startPos = InStr(startPos + searchTextLength, LCase(cell.Value), searchText, vbTextCompare)
                            count = count + 1
                        Wend
                    End If
                End If
            End If
        Next cell
        
        ' ��� �޽��� ǥ��
        If count > 0 Then
            MsgBox count & "���� ������ '" & searchText & "'(��)�� �˻��Ǿ� ������ �������� ���� ǥ�õǾ����ϴ�.", vbInformation, "�˻� ���"
        Else
            MsgBox "'" & searchText & "'(��)�� �����ϴ� ���� �����ϴ�.", vbInformation, "�˻� ���"
        End If
    Else
        MsgBox "�˻��� �ܾ� �Ǵ� ������ �Էµ��� �ʾҽ��ϴ�.", vbExclamation, "�Է� ����"
    End If
End Sub

Function ColorToRGB(colorText As String) As Long
    Dim colorParts() As String
    If InStr(colorText, ",") > 0 Then
        colorParts = Split(colorText, ",")
        If UBound(colorParts) = 2 Then
            ColorToRGB = RGB(val(colorParts(0)), val(colorParts(1)), val(colorParts(2)))
        Else
            ColorToRGB = RGB(255, 0, 0) ' �⺻���� ������
        End If
    Else
        Select Case LCase(colorText)
            Case "red"
                ColorToRGB = RGB(255, 0, 0)
            Case "green"
                ColorToRGB = RGB(0, 128, 0)
            Case "blue"
                ColorToRGB = RGB(0, 0, 255)
            Case Else
                ColorToRGB = RGB(0, 0, 0) ' �⺻���� ������
        End Select
    End If
End Function

