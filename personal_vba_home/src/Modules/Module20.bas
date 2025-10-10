Attribute VB_Name = "Module20"
Option Explicit

Sub AnalyzeKeywords()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim basePeriod As String
    Dim comparePeriod1 As String
    Dim comparePeriod2 As String
    Dim proceed As VbMsgBoxResult
    Dim wb As Workbook

    ' ���� Ȱ��ȭ�� ��ũ�Ͽ��� �۾�
    Set wb = ActiveWorkbook
    On Error Resume Next
    Set ws = wb.Sheets("������")
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "������ ��Ʈ�� ã�� �� �����ϴ�. '������'��� �̸��� ��Ʈ�� Ȯ���ϼ���.", vbExclamation
        Exit Sub
    End If
    
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row

    ' ����ڷκ��� �Ⱓ �Է� �ޱ�
    basePeriod = InputBox("�м��� ���� �Ⱓ�� �Է��ϼ��� (��: 2024��11��):", "���� �Ⱓ �Է�")
    If basePeriod = "" Then Exit Sub ' �Է� ��� �� ����

    comparePeriod1 = InputBox("���� ù ��° �Ⱓ�� �Է��ϼ��� (��: 2023��11��):", "�� �Ⱓ �Է�")
    If comparePeriod1 = "" Then Exit Sub ' �Է� ��� �� ����

    comparePeriod2 = InputBox("���� �� ��° �Ⱓ�� �Է��ϼ��� (��: 2024��10��):", "�� �Ⱓ �Է�")
    If comparePeriod2 = "" Then Exit Sub ' �Է� ��� �� ����

    ' ����� Ȯ��
    proceed = MsgBox("���� �Ⱓ: " & basePeriod & vbCrLf & _
                     "�� �Ⱓ 1: " & comparePeriod1 & vbCrLf & _
                     "�� �Ⱓ 2: " & comparePeriod2 & vbCrLf & _
                     "�м��� �����Ͻðڽ��ϱ�?", vbYesNo + vbQuestion)
    If proceed = vbNo Then Exit Sub

    ' �� �����ƾ ����
    ' ���� �Ⱓ���� �����ϴ� �ű� �˻�� �����մϴ�.
    ' �� �Ⱓ 1�� �� �Ⱓ 2�� �������� ���ο� Ű���带 ã���ϴ�.
    ExtractNewKeywordsBase ws, lastRow, basePeriod, comparePeriod1, comparePeriod2
    
    ' ���� �Ⱓ���� ������ ����� �˻�� �����մϴ�.
    ' �� ����� �� �Ⱓ 1�Դϴ�.
    ExtractRisingKeywordsBase ws, lastRow, basePeriod, comparePeriod1
    
    ' ���� �Ⱓ ��� �� �Ⱓ 2���� �ű� �˻�� �����մϴ�.
    ' ���� �Ⱓ���� �����ϴ� Ű���带 Ȯ���մϴ�.
    ExtractNewKeywordsVsPeriod ws, lastRow, basePeriod, comparePeriod2
    
    ' ���� �Ⱓ ��� �� �Ⱓ 1���� �ű� �˻�� �����մϴ�.
    ' ���� �Ⱓ���� �����ϴ� Ű���带 Ȯ���մϴ�.
    ExtractNewKeywordsVsPeriod ws, lastRow, basePeriod, comparePeriod1
    
    ' ���� �Ⱓ���� �� �Ⱓ 2�� ���Ͽ� ������ ����� �˻�� �����մϴ�.
    ' ���� ������ ����Ͽ� ����� Ű���常 ����մϴ�.
    ExtractRisingKeywordsVsPeriod ws, lastRow, basePeriod, comparePeriod2

    MsgBox "��� �м��� �Ϸ�Ǿ����ϴ�.", vbInformation
End Sub

Sub ExtractNewKeywordsBase(ws As Worksheet, lastRow As Long, basePeriod As String, comparePeriod1 As String, comparePeriod2 As String)
    Dim newSheet As Worksheet
    Dim i As Long
    Dim keyword As Variant
    Dim outputRow As Long
    Dim keywordsPrev As Object
    Dim keywordsBase As Object
    Dim sheetName As String

    Set keywordsPrev = CreateObject("Scripting.Dictionary")
    Set keywordsBase = CreateObject("Scripting.Dictionary")

    ' �� �Ⱓ�� Ű���� �ε�
    For i = 2 To lastRow
        If ws.Cells(i, 3).Value = comparePeriod1 Or ws.Cells(i, 3).Value = comparePeriod2 Then
            keyword = ws.Cells(i, 2).Value
            keywordsPrev(keyword) = True
        End If
    Next i

    ' ���� �Ⱓ�� Ű���� �ε�
    For i = 2 To lastRow
        If ws.Cells(i, 3).Value = basePeriod Then
            keyword = ws.Cells(i, 2).Value
            keywordsBase(keyword) = ws.Cells(i, 1).Value ' ���� ����
        End If
    Next i

    ' ��� ��Ʈ ����
    sheetName = basePeriod & " �ű� �˻���"
    Set newSheet = CreateResultSheet(sheetName)
    newSheet.Cells(1, 1).Value = "����"
    newSheet.Cells(1, 2).Value = "�α�˻���"
    outputRow = 2

    ' �ű� �˻��� ����
    For Each keyword In keywordsBase.Keys
        If Not keywordsPrev.exists(keyword) Then
            newSheet.Cells(outputRow, 1).Value = keywordsBase(keyword) ' ����
            newSheet.Cells(outputRow, 2).Value = keyword ' �α�˻���
            outputRow = outputRow + 1
        End If
    Next keyword

    MsgBox "�� " & (outputRow - 2) & "���� �ű� Ű���尡 '" & sheetName & "' ��Ʈ�� �ۼ��Ǿ����ϴ�.", vbInformation
End Sub

Sub ExtractRisingKeywordsBase(ws As Worksheet, lastRow As Long, basePeriod As String, comparePeriod As String)
    Dim newSheet As Worksheet
    Dim i As Long
    Dim keyword As Variant
    Dim outputRow As Long
    Dim keywordsCompare As Object
    Dim keywordsBase As Object
    Dim rankCompare As Long
    Dim rankBase As Long
    Dim sheetName As String

    Set keywordsCompare = CreateObject("Scripting.Dictionary")
    Set keywordsBase = CreateObject("Scripting.Dictionary")

    ' �� �Ⱓ�� Ű����� ���� �ε�
    For i = 2 To lastRow
        If ws.Cells(i, 3).Value = comparePeriod Then
            keyword = ws.Cells(i, 2).Value
            keywordsCompare(keyword) = ws.Cells(i, 1).Value ' ���� ����
        End If
    Next i

    ' ���� �Ⱓ�� Ű����� ���� �ε�
    For i = 2 To lastRow
        If ws.Cells(i, 3).Value = basePeriod Then
            keyword = ws.Cells(i, 2).Value
            keywordsBase(keyword) = ws.Cells(i, 1).Value ' ���� ����
        End If
    Next i

    ' ��� ��Ʈ ����
    sheetName = comparePeriod & " ��� ���� ��� �˻���"
    Set newSheet = CreateResultSheet(sheetName)
    newSheet.Cells(1, 1).Value = basePeriod & "_����"
    newSheet.Cells(1, 2).Value = "�α�˻���"
    newSheet.Cells(1, 3).Value = "��������"
    outputRow = 2

    ' ���� ��� �˻��� ����
    For Each keyword In keywordsBase.Keys
        If keywordsCompare.exists(keyword) Then
            rankCompare = keywordsCompare(keyword)
            rankBase = keywordsBase(keyword)
            If IsNumeric(rankCompare) And IsNumeric(rankBase) Then
                If rankBase < rankCompare Then
                    newSheet.Cells(outputRow, 1).Value = rankBase ' ���� �Ⱓ ����
                    newSheet.Cells(outputRow, 2).Value = keyword ' �α�˻���
                    newSheet.Cells(outputRow, 3).Value = rankCompare - rankBase ' ��������
                    outputRow = outputRow + 1
                End If
            End If
        End If
    Next keyword

    MsgBox "�� " & (outputRow - 2) & "���� ���� ��� Ű���尡 '" & sheetName & "' ��Ʈ�� �ۼ��Ǿ����ϴ�.", vbInformation
End Sub

Sub ExtractNewKeywordsVsPeriod(ws As Worksheet, lastRow As Long, basePeriod As String, comparePeriod As String)
    Dim newSheet As Worksheet
    Dim i As Long
    Dim keyword As Variant
    Dim outputRow As Long
    Dim keywordsCompare As Object
    Dim keywordsBase As Object
    Dim sheetName As String

    Set keywordsCompare = CreateObject("Scripting.Dictionary")
    Set keywordsBase = CreateObject("Scripting.Dictionary")

    ' �� �Ⱓ�� Ű���� �ε�
    For i = 2 To lastRow
        If ws.Cells(i, 3).Value = comparePeriod Then
            keyword = ws.Cells(i, 2).Value
            keywordsCompare(keyword) = True
        End If
    Next i

    ' ���� �Ⱓ�� Ű���� �ε�
    For i = 2 To lastRow
        If ws.Cells(i, 3).Value = basePeriod Then
            keyword = ws.Cells(i, 2).Value
            keywordsBase(keyword) = ws.Cells(i, 1).Value ' ���� ����
        End If
    Next i

    ' ��� ��Ʈ ����
    sheetName = comparePeriod & " ��� �ű� �˻���"
    Set newSheet = CreateResultSheet(sheetName)
    newSheet.Cells(1, 1).Value = "����"
    newSheet.Cells(1, 2).Value = "�α�˻���"
    outputRow = 2

    ' �ű� �˻��� ����
    For Each keyword In keywordsBase.Keys
        If Not keywordsCompare.exists(keyword) Then
            newSheet.Cells(outputRow, 1).Value = keywordsBase(keyword) ' ����
            newSheet.Cells(outputRow, 2).Value = keyword ' �α�˻���
            outputRow = outputRow + 1
        End If
    Next keyword

    MsgBox "�� " & (outputRow - 2) & "���� �ű� Ű���尡 '" & sheetName & "' ��Ʈ�� �ۼ��Ǿ����ϴ�.", vbInformation
End Sub

Sub ExtractRisingKeywordsVsPeriod(ws As Worksheet, lastRow As Long, basePeriod As String, comparePeriod As String)
    Dim newSheet As Worksheet
    Dim i As Long
    Dim keyword As Variant
    Dim outputRow As Long
    Dim keywordsCompare As Object
    Dim keywordsBase As Object
    Dim rankCompare As Long
    Dim rankBase As Long
    Dim sheetName As String

    Set keywordsCompare = CreateObject("Scripting.Dictionary")
    Set keywordsBase = CreateObject("Scripting.Dictionary")

    ' �� �Ⱓ�� Ű����� ���� �ε�
    For i = 2 To lastRow
        If ws.Cells(i, 3).Value = comparePeriod Then
            keyword = ws.Cells(i, 2).Value
            keywordsCompare(keyword) = ws.Cells(i, 1).Value ' ���� ����
        End If
    Next i

    ' ���� �Ⱓ�� Ű����� ���� �ε�
    For i = 2 To lastRow
        If ws.Cells(i, 3).Value = basePeriod Then
            keyword = ws.Cells(i, 2).Value
            keywordsBase(keyword) = ws.Cells(i, 1).Value ' ���� ����
        End If
    Next i

    ' ��� ��Ʈ ����
    sheetName = comparePeriod & " ��� ���� ��� �˻���"
    Set newSheet = CreateResultSheet(sheetName)
    newSheet.Cells(1, 1).Value = basePeriod & "_����"
    newSheet.Cells(1, 2).Value = "�α�˻���"
    newSheet.Cells(1, 3).Value = "��������"
    outputRow = 2

    ' ���� ��� �˻��� ����
    For Each keyword In keywordsBase.Keys
        If keywordsCompare.exists(keyword) Then
            rankCompare = keywordsCompare(keyword)
            rankBase = keywordsBase(keyword)
            If IsNumeric(rankCompare) And IsNumeric(rankBase) Then
                If rankBase < rankCompare Then
                    newSheet.Cells(outputRow, 1).Value = rankBase ' ���� �Ⱓ ����
                    newSheet.Cells(outputRow, 2).Value = keyword ' �α�˻���
                    newSheet.Cells(outputRow, 3).Value = rankCompare - rankBase ' ��������
                    outputRow = outputRow + 1
                End If
            End If
        End If
    Next keyword

    MsgBox "�� " & (outputRow - 2) & "���� ���� ��� Ű���尡 '" & sheetName & "' ��Ʈ�� �ۼ��Ǿ����ϴ�.", vbInformation
End Sub

Function CreateResultSheet(sheetName As String) As Worksheet
    Dim ws As Worksheet
    ' ���� ��Ʈ�� ������ ����
    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets(sheetName).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    ' ���ο� ��Ʈ ����
    Set ws = ActiveWorkbook.Sheets.Add
    ws.Name = sheetName
    Set CreateResultSheet = ws
End Function




