Attribute VB_Name = "Module22"
' ==============================================
' ���α׷��� : URL�����ϰ������۸�ũ�����
' ���� :
'   - B��(2��° ��)�� �ִ� �˻�� ������� ���̹�, ���̹�����, ����, ��Ʃ��, ƽ��, ���ζ��̺�, �˻���Ʈ����, ���Ų��� �˻� URL�� �ڵ� �����մϴ�.
'   - �� URL�� �̵��� �� �ִ� �����۸�ũ�� ���ο� ���� �����մϴ�.
'   - URL�� �� ���� �����, ��ü �����Ϳ� �׵θ��� �߰��մϴ�.
'   - ���� �����Ϳ� ������ ���� �ʵ��� ù ��° �� ������ ����� �ۼ��մϴ�.
'   - URL ���ڵ��� ���� �ѱ�/Ư������ �˻�� ���� ó���մϴ�.
'   - �۾� ���� ��Ȳ�� ����ǥ���ٿ� ǥ���ϰ�, �Ϸ� �� �޽��� �ڽ��� ���ϴ�.
' ���� :
'   1. B��(2��° ��)�� �˻�� �Է��մϴ�(2�����).
'   2. ��Ʈ�� Ȱ��ȭ�� ���¿��� �� ��ũ�θ� �����մϴ�.
'   3. �ڵ����� URL �� �����۸�ũ�� �����˴ϴ�.
' ==============================================
'
' ====== [EUC-KR ���ڵ� �Լ� �߰�] ======
#If VBA7 Then
    Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" ( _
        ByVal CodePage As Long, ByVal dwFlags As Long, _
        ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long, _
        ByVal lpMultiByteStr As LongPtr, ByVal cbMultiByte As Long, _
        ByVal lpDefaultChar As LongPtr, ByVal lpUsedDefaultChar As LongPtr) As Long
#Else
    Private Declare Function WideCharToMultiByte Lib "kernel32" ( _
        ByVal CodePage As Long, ByVal dwFlags As Long, _
        ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, _
        ByVal lpMultiByteStr As Long, ByVal cbMultiByte As Long, _
        ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
#End If

Function EncodeEUC_KR(str As String) As String
    Dim l As Long, i As Long
    Dim arr() As Byte
    Dim result As String
    Dim b As Byte
    Dim tempStr As String

    l = WideCharToMultiByte(949, 0, StrPtr(str), -1, 0, 0, 0, 0)
    If l > 1 Then
        ReDim arr(l - 2)
        WideCharToMultiByte 949, 0, StrPtr(str), -1, VarPtr(arr(0)), l - 1, 0, 0
        For i = 0 To UBound(arr)
            b = arr(i)
            If (b >= &H30 And b <= &H39) Or (b >= &H41 And b <= &H5A) Or (b >= &H61 And b <= &H7A) Then
                tempStr = Chr(b)
            Else
                tempStr = "%" & Right("0" & Hex(b), 2)
            End If
            result = result & tempStr
        Next i
    End If
    EncodeEUC_KR = result
End Function

Sub URL�����ϰ������۸�ũ�����()
    Dim ws As Worksheet ' �۾��� ��Ʈ ��ü
    Dim lastRow As Long ' B���� ������ ������ �� ��ȣ
    Dim startCol As Long ' ù ��° �� �� ��ȣ
    Dim startColURL As Long ' URL�� �� ���� �� ��ȣ
    Dim startColLink As Long ' �����۸�ũ�� �� ���� �� ��ȣ
    Dim i As Long ' �ݺ��� �ε���
    Dim dataRange As Range ' �׵θ� ���� ����
    
    ' ���� Ȱ��ȭ�� ��Ʈ ����
    On Error Resume Next
    Set ws = ActiveSheet
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "Ȱ��ȭ�� ��Ʈ�� �����ϴ�. ���� �۾��� ��Ʈ�� Ȱ��ȭ�ϼ���!", vbExclamation
        Exit Sub
    End If
    
    ' B�� �����Ͱ� �ִ� ������ �� ã�� (2����� �����Ͱ� �ִٰ� ����)
    lastRow = ws.Cells(ws.Rows.count, "B").End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "B���� �����Ͱ� �����ϴ�. Ȯ�����ּ���!", vbExclamation
        Exit Sub
    End If
    
    ' ���� ǥ���� �ʱ�ȭ
    Application.StatusBar = "�۾� ����: URL ���� ��..."
    
    ' ù ��° �� �� ã�� (���� �����Ϳ� ���� ������)
    startCol = 1
    Do While Application.WorksheetFunction.CountA(ws.Columns(startCol)) > 0
        startCol = startCol + 1
    Loop
    
    ' URL �� �����۸�ũ �߰��� �� ��ġ ����
    startColURL = startCol ' URL �� ����
    startColLink = startColURL + 8 ' �����۸�ũ �� ���� (URL �� 8�� ��)
    
    ' URL �� ���� �߰�
    ws.Cells(1, startColURL).Value = "���̹��˻� URL"
    ws.Cells(1, startColURL + 1).Value = "���̹����ΰ˻� URL"
    ws.Cells(1, startColURL + 2).Value = "����URL"
    ws.Cells(1, startColURL + 3).Value = "��Ʃ��URL"
    ws.Cells(1, startColURL + 4).Value = "ƽ��URL"
    ws.Cells(1, startColURL + 5).Value = "���ζ��̺�URL"
    ws.Cells(1, startColURL + 6).Value = "�˻���Ʈ���� URL"
    ws.Cells(1, startColURL + 7).Value = "���Ųڰ˻� URL"
    
    ' �� �ึ�� �˻���� URL ���� �� �߰�
    For i = 2 To lastRow
        Application.StatusBar = "URL ���� ��... (" & i - 1 & "/" & lastRow - 1 & ")"
        
        If ws.Cells(i, 2).Value <> "" Then ' B���� �˻�� ���� ����
            ' ���̹� �˻� URL
            ws.Cells(i, startColURL).Value = "https://search.naver.com/search.naver?where=nexearch&sm=top_hty&fbm=0&ie=utf8&query=" & WorksheetFunction.EncodeURL(ws.Cells(i, 2).Value)
            ' ���̹����� �˻� URL
            ws.Cells(i, startColURL + 1).Value = "https://search.shopping.naver.com/ns/search?query=" & WorksheetFunction.EncodeURL(ws.Cells(i, 2).Value)
            ' ���� �˻� URL
            ws.Cells(i, startColURL + 2).Value = "https://www.coupang.com/np/search?q=" & WorksheetFunction.EncodeURL(ws.Cells(i, 2).Value)
            ' ��Ʃ�� �˻� URL
            ws.Cells(i, startColURL + 3).Value = "https://www.youtube.com/results?search_query=" & WorksheetFunction.EncodeURL(ws.Cells(i, 2).Value)
            ' ƽ�� �˻� URL
            ws.Cells(i, startColURL + 4).Value = "https://www.tiktok.com/search?q=" & WorksheetFunction.EncodeURL(ws.Cells(i, 2).Value)
            ' ���ζ��̺� �˻� URL
            ws.Cells(i, startColURL + 5).Value = "https://shoppinglive.naver.com/search/lives?query=" & WorksheetFunction.EncodeURL(ws.Cells(i, 2).Value)
            ' �˻���Ʈ���� URL (����)
            ws.Cells(i, startColURL + 6).Value = "https://datalab.naver.com/keyword/trendSearch.naver"
            ' ���Ų� �˻� URL (EUC-KR ���ڵ� ����)
            ws.Cells(i, startColURL + 7).Value = "https://domeggook.com/main/item/itemList.php?sfc=ttl&sf=ttl&sw=" & EncodeEUC_KR(ws.Cells(i, 2).Value)
        End If
    Next i
    
    ' �����۸�ũ �� ���� �߰�
    ws.Cells(1, startColLink).Value = "���̹��˻�"
    ws.Cells(1, startColLink + 1).Value = "���̹����ΰ˻�"
    ws.Cells(1, startColLink + 2).Value = "����"
    ws.Cells(1, startColLink + 3).Value = "��Ʃ��"
    ws.Cells(1, startColLink + 4).Value = "ƽ��"
    ws.Cells(1, startColLink + 5).Value = "���ζ��̺�"
    ws.Cells(1, startColLink + 6).Value = "�˻���Ʈ����"
    ws.Cells(1, startColLink + 7).Value = "���Ųڰ˻�"
    
    ' �� �ึ�� �����۸�ũ �߰�
    For i = 2 To lastRow
        Application.StatusBar = "�����۸�ũ �߰� ��... (" & i - 1 & "/" & lastRow - 1 & ")"
        
        ' �� URL�� ���� ���� �����۸�ũ ����
        If ws.Cells(i, startColURL).Value <> "" Then
            ws.Hyperlinks.Add ws.Cells(i, startColLink), ws.Cells(i, startColURL).Value, , , "�ٷΰ���"
        End If
        If ws.Cells(i, startColURL + 1).Value <> "" Then
            ws.Hyperlinks.Add ws.Cells(i, startColLink + 1), ws.Cells(i, startColURL + 1).Value, , , "�ٷΰ���"
        End If
        If ws.Cells(i, startColURL + 2).Value <> "" Then
            ws.Hyperlinks.Add ws.Cells(i, startColLink + 2), ws.Cells(i, startColURL + 2).Value, , , "�ٷΰ���"
        End If
        If ws.Cells(i, startColURL + 3).Value <> "" Then
            ws.Hyperlinks.Add ws.Cells(i, startColLink + 3), ws.Cells(i, startColURL + 3).Value, , , "�ٷΰ���"
        End If
        If ws.Cells(i, startColURL + 4).Value <> "" Then
            ws.Hyperlinks.Add ws.Cells(i, startColLink + 4), ws.Cells(i, startColURL + 4).Value, , , "�ٷΰ���"
        End If
        If ws.Cells(i, startColURL + 5).Value <> "" Then
            ws.Hyperlinks.Add ws.Cells(i, startColLink + 5), ws.Cells(i, startColURL + 5).Value, , , "�ٷΰ���"
        End If
        ' �˻���Ʈ����� ���� URL�� �����۸�ũ ����
        ws.Hyperlinks.Add ws.Cells(i, startColLink + 6), "https://datalab.naver.com/keyword/trendSearch.naver", , , "�ٷΰ���"
        ' ���Ų� �����۸�ũ ����
        If ws.Cells(i, startColURL + 7).Value <> "" Then
            ws.Hyperlinks.Add ws.Cells(i, startColLink + 7), ws.Cells(i, startColURL + 7).Value, , , "�ٷΰ���"
        End If
    Next i
    
    ' URL �� ����� (����ڿ��� URL�� ������ �ʵ���)
    ws.Range(ws.Columns(startColURL), ws.Columns(startColURL + 7)).EntireColumn.Hidden = True
    
    ' ��ü �����Ϳ� �׵θ� �߰�
    Set dataRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, startColLink + 7))
    With dataRange.Borders
        .LineStyle = xlContinuous
        .Color = RGB(0, 0, 0)
        .Weight = xlThin
    End With
    
    ' �Ϸ� �޽��� �� ����ǥ���� �ʱ�ȭ
    Application.StatusBar = "�۾� �Ϸ�!"
    Application.StatusBar = False
    MsgBox "URL�� �����۸�ũ�� �����ǰ�, URL ���� ���������� �׵θ��� �߰��Ǿ����ϴ�!", vbInformation
End Sub




