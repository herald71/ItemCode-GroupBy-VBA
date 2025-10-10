Attribute VB_Name = "Module30"
Option Explicit

Sub AnalyzeKeywords_2���м�()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim basePeriod As String
    Dim comparePeriod As String
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

    comparePeriod = InputBox("���� �Ⱓ�� �Է��ϼ��� (��: 2023��11��):", "�� �Ⱓ �Է�")
    If comparePeriod = "" Then Exit Sub ' �Է� ��� �� ����

    ' ����� Ȯ��
    proceed = MsgBox("���� �Ⱓ: " & basePeriod & vbCrLf & _
                     "�� �Ⱓ: " & comparePeriod & vbCrLf & _
                     "�м��� �����Ͻðڽ��ϱ�?", vbYesNo + vbQuestion)
    If proceed = vbNo Then Exit Sub

    ' �� �����ƾ ����
    ExtractNewKeywords_Personal ws, lastRow, basePeriod, comparePeriod
    ExtractRisingKeywords_Personal ws, lastRow, basePeriod, comparePeriod

    MsgBox "��� �м��� �Ϸ�Ǿ����ϴ�.", vbInformation
End Sub

Sub ExtractNewKeywords_Personal(ws As Worksheet, lastRow As Long, basePeriod As String, comparePeriod As String)
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
        If ws.Cells(i, 3).Value = comparePeriod Then
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
    Set newSheet = CreateResultSheet_Personal(sheetName)
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

Sub ExtractRisingKeywords_Personal(ws As Worksheet, lastRow As Long, basePeriod As String, comparePeriod As String)
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
    Set newSheet = CreateResultSheet_Personal(sheetName)
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

Function CreateResultSheet_Personal(sheetName As String) As Worksheet
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
    Set CreateResultSheet_Personal = ws
End Function




