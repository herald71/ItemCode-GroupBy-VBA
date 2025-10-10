Attribute VB_Name = "Module26"
Sub ���͸�_��_����()
    Dim ws As Worksheet
    Dim newWs As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim filterRange As Range
    Dim copyRange As Range
    Dim wsName As String
    Dim i As Integer
    Dim userResponse As Variant
    Dim searchMonths As String
    Dim monthArray() As String
    Dim minVal As Double, maxVal As Double
    Dim filterSummary As String
    Dim recentSearchRange As String
    Dim coupangPriceRange As String
    Dim coupangReviewRange As String
    Dim coupangRocketRatio As String
    Dim coupangSellerRocketRatio As String
    
    ' ���� ���õ� ��Ʈ
    Set ws = ActiveSheet

    ' ������ ��� �� ã��
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).column

    ' �����Ͱ� ���� ��� ����
    If lastRow < 2 Then
        MsgBox "�����Ͱ� �����ϴ�. ������ �����մϴ�.", vbExclamation
        Exit Sub
    End If

    ' ���� AutoFilter ����
    If ws.AutoFilterMode Then ws.AutoFilterMode = False

    ' ���͸� ������ ��ü ���� ����
    Set filterRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
    
    ' AutoFilter Ȱ��ȭ
    filterRange.AutoFilter

    ' ���͸� ��� �ʱ�ȭ
    filterSummary = "���͸� ���� ���:" & vbLf

    ' �귣�� Ű����(O) ���� ���� Ȯ��
    userResponse = MsgBox("�귣�� Ű���尡 'O'�� �׸��� �����Ͻðڽ��ϱ�?", vbYesNo, "���͸� �ɼ�")
    If userResponse = vbYes Then
        filterRange.AutoFilter Field:=4, Criteria1:="<>O"
        filterSummary = filterSummary & "- �귣�� Ű����(O) ����" & vbLf
    End If

    ' ���μ� Ű����(X) ����
    filterRange.AutoFilter Field:=5, Criteria1:="<>X"
    filterSummary = filterSummary & "- ���μ� Ű����(X) ����" & vbLf

    ' �ֱ� 1���� �˻��� ���͸�
    recentSearchRange = InputBox("�ֱ� 1���� �˻��� ������ �Է��ϼ��� (��: 1000~100000)")
    If recentSearchRange <> "" And IsValidRange(recentSearchRange) Then
        minVal = Split(recentSearchRange, "~")(0)
        maxVal = Split(recentSearchRange, "~")(1)
        filterRange.AutoFilter Field:=7, Criteria1:=">=" & minVal, Criteria2:="<=" & maxVal
        filterSummary = filterSummary & "- �ֱ� 1���� �˻���: " & recentSearchRange & vbLf
    End If

    ' �۳� �ִ� �˻� �� ���͸�
    searchMonths = InputBox("�۳� �ִ� �˻� ���� �Է��ϼ��� (��: 4,5,6,7,8)")
    If searchMonths <> "" And IsValidMonthList(searchMonths) Then
        monthArray = Split(searchMonths, ",")
        filterRange.AutoFilter Field:=14, Criteria1:=monthArray, Operator:=xlFilterValues
        filterSummary = filterSummary & "- �۳� �ִ� �˻� ��: " & searchMonths & vbLf
    End If

    ' ���� ��հ� ���͸�
    coupangPriceRange = InputBox("���� ��հ� ������ �Է��ϼ��� (��: 9800~29999)")
    If coupangPriceRange <> "" And IsValidRange(coupangPriceRange) Then
        minVal = Split(coupangPriceRange, "~")(0)
        maxVal = Split(coupangPriceRange, "~")(1)
        filterRange.AutoFilter Field:=26, Criteria1:=">=" & minVal, Criteria2:="<=" & maxVal
        filterSummary = filterSummary & "- ���� ��հ�: " & coupangPriceRange & vbLf
    End If

    ' ���� ��ո���� ���͸�
    coupangReviewRange = InputBox("���� ��ո���� ������ �Է��ϼ��� (��: 0~200)")
    If coupangReviewRange <> "" And IsValidRange(coupangReviewRange) Then
        minVal = Split(coupangReviewRange, "~")(0)
        maxVal = Split(coupangReviewRange, "~")(1)
        filterRange.AutoFilter Field:=29, Criteria1:=">=" & minVal, Criteria2:="<=" & maxVal
        filterSummary = filterSummary & "- ���� ��ո����: " & coupangReviewRange & vbLf
    End If

    ' ���� ���Ϲ�ۺ��� ���͸�
    coupangRocketRatio = InputBox("���� ���Ϲ�ۺ��� ������ �Է��ϼ��� (��: 0~50)")
    If coupangRocketRatio <> "" And IsValidRange(coupangRocketRatio) Then
        minVal = CDbl(Split(coupangRocketRatio, "~")(0)) / 100
        maxVal = CDbl(Split(coupangRocketRatio, "~")(1)) / 100
        filterRange.AutoFilter Field:=30, Criteria1:=">=" & minVal, Criteria2:="<=" & maxVal
        filterSummary = filterSummary & "- ���� ���Ϲ�ۺ���: " & coupangRocketRatio & vbLf
    End If

    ' ���� �Ǹ��ڷ��� ��ۺ��� ���͸�
    coupangSellerRocketRatio = InputBox("���� �Ǹ��ڷ��� ��ۺ��� ������ �Է��ϼ��� (��: 0~50)")
    If coupangSellerRocketRatio <> "" And IsValidRange(coupangSellerRocketRatio) Then
        minVal = CDbl(Split(coupangSellerRocketRatio, "~")(0)) / 100
        maxVal = CDbl(Split(coupangSellerRocketRatio, "~")(1)) / 100
        filterRange.AutoFilter Field:=31, Criteria1:=">=" & minVal, Criteria2:="<=" & maxVal
        filterSummary = filterSummary & "- ���� �Ǹ��ڷ��� ��ۺ���: " & coupangSellerRocketRatio & vbLf
    End If

    ' ���͸��� �����͸� �����Ͽ� ���ο� ��Ʈ�� ����
    i = 1
    wsName = "����Ű����"
    Do While SheetExists(wsName & i)
        i = i + 1
    Loop
    wsName = wsName & i
    
    Set newWs = ActiveWorkbook.Sheets.Add
    newWs.Name = wsName
    
    On Error Resume Next
    Set copyRange = filterRange.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0

    If copyRange Is Nothing Then
        MsgBox "���͸��� ����� �����ϴ�. ������ �����մϴ�.", vbExclamation
        ws.AutoFilterMode = False
        Exit Sub
    End If
    
    copyRange.Copy
    newWs.Range("A1").PasteSpecial Paste:=xlPasteValues

    ' ���͸� ���� ����� ������ ���� B���� �߰�
    lastRow = newWs.Cells(newWs.Rows.count, "A").End(xlUp).Row
    newWs.Cells(lastRow + 1, 2).Value = filterSummary
    newWs.Cells(lastRow + 1, 2).WrapText = True
    
    ' ���͸� ����� B1 ���� ��Ʈ(�޸�)�� �߰�
    With newWs.Range("B1")
        .ClearComments ' ���� ��Ʈ ����
        .AddComment.text text:=filterSummary
    End With

    ws.AutoFilterMode = False

    MsgBox "���͸� �� ������ ���簡 �Ϸ�Ǿ����ϴ�. ����� '" & wsName & "' ��Ʈ�� ����Ǿ����ϴ�.", vbInformation

End Sub

Function SheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

Function IsValidRange(ByVal userInput As String) As Boolean
    IsValidRange = (InStr(userInput, "~") > 0)
End Function

Function IsValidMonthList(ByVal userInput As String) As Boolean
    IsValidMonthList = (userInput <> "")
End Function

