Attribute VB_Name = "Module36"
' ===============================================
' ���α׷��� : Ű���� ��ȭ �ڵ� �м� ����
' �ۼ�����   : 2025-07-19
' ����       : v2.1 (��� ��Ʈ ���� + �����űԾ� ���� ����)
' ����       : '������' ��Ʈ�� C��(�Ⱓ)�� �ڵ� �ν��Ͽ�
'              ����/�񱳱Ⱓ�� ���� �ű�/��� Ű���带 �м��ϰ�
'              ��� ��Ʈ�� ���� �� ��� ���� ��Ʈ�� ����
' ===============================================

Option Explicit

' �� ���� ���� ���ν���: C������ 3�� �ֽ� �Ⱓ�� �ν��� �м� ����
Sub AnalyzeKeywords_AutoPeriod()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim wb As Workbook
    Dim dictPeriods As Object
    Dim periodList() As Variant
    Dim i As Long
    Dim cellValue As String
    Dim basePeriod As String, comparePeriod1 As String, comparePeriod2 As String

    Set wb = ActiveWorkbook
    Set ws = wb.Sheets("������")
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row

    ' --- C���� �Ⱓ ���� ���� ---
    Set dictPeriods = CreateObject("Scripting.Dictionary")
    For i = 2 To lastRow
        cellValue = Trim(ws.Cells(i, 3).Value)
        If Len(cellValue) > 0 Then dictPeriods(cellValue) = True
    Next i

    ' --- �Ⱓ�� 3�� �̻� �־�� �м� ���� ---
    If dictPeriods.count < 3 Then
        MsgBox "�Ⱓ ������ 3�� �̻� �������� �ʽ��ϴ�.", vbExclamation
        Exit Sub
    End If

    ' --- �������� ����: �ֽ� ������ �����Ͽ� ���ء��񱳽��� ���� ---
    periodList = dictPeriods.Keys
    SortDescending periodList
    basePeriod = periodList(0)
    comparePeriod1 = periodList(1)
    comparePeriod2 = periodList(2)

    MsgBox "���� �Ⱓ: " & basePeriod & vbCrLf & _
           "�� �Ⱓ 1: " & comparePeriod1 & vbCrLf & _
           "�� �Ⱓ 2: " & comparePeriod2, vbInformation

    ' --- 5���� �м� ���� ---
    ExtractNewKeywordsBase ws, lastRow, basePeriod, comparePeriod1, comparePeriod2
    ExtractRisingKeywordsBase ws, lastRow, basePeriod, comparePeriod1
    ExtractNewKeywordsVsPeriod ws, lastRow, basePeriod, comparePeriod2
    ExtractNewKeywordsVsPeriod ws, lastRow, basePeriod, comparePeriod1
    ExtractRisingKeywordsVsPeriod ws, lastRow, basePeriod, comparePeriod2
    ExtractRisingFromPastToCurrent ws, lastRow, basePeriod, comparePeriod1, comparePeriod2

    ' --- ��� ���� ���� ---
    Call CreateFormattedSummaryReport
    MsgBox "��� �м��� �Ϸ�Ǿ����ϴ�.", vbInformation
End Sub

' �� ���ڿ� �迭 �������� ���� �Լ�
Sub SortDescending(arr() As Variant)
    Dim i As Long, j As Long, temp As String
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i) < arr(j) Then
                temp = arr(i): arr(i) = arr(j): arr(j) = temp
            End If
        Next j
    Next i
End Sub

' �� ���� �Ⱓ���� �ִ� Ű���� ���� ("�����űԾ�")
Sub ExtractNewKeywordsBase(ws As Worksheet, lastRow As Long, basePeriod As String, comparePeriod1 As String, comparePeriod2 As String)
    Dim newSheet As Worksheet, i As Long, keyword As Variant, outputRow As Long
    Dim keywordsPrev As Object, keywordsBase As Object
    Dim sheetName As String

    Set keywordsPrev = CreateObject("Scripting.Dictionary")
    Set keywordsBase = CreateObject("Scripting.Dictionary")

    ' --- �� �Ⱓ 1, 2���� ������ Ű���� ���� ---
    For i = 2 To lastRow
        If ws.Cells(i, 3).Value = comparePeriod1 Or ws.Cells(i, 3).Value = comparePeriod2 Then
            keyword = ws.Cells(i, 2).Value
            keywordsPrev(keyword) = True
        End If
    Next i

    ' --- ���� �Ⱓ Ű���� ���� ---
    For i = 2 To lastRow
        If ws.Cells(i, 3).Value = basePeriod Then
            keyword = ws.Cells(i, 2).Value
            keywordsBase(keyword) = ws.Cells(i, 1).Value
        End If
    Next i

    ' --- ��� ��Ʈ ���� �� ��� ---
    sheetName = basePeriod & " �����űԾ�"
    Set newSheet = CreateResultSheet(sheetName)
    newSheet.Cells(1, 1).Value = "����"
    newSheet.Cells(1, 2).Value = "�α�˻���"
    outputRow = 2

    For Each keyword In keywordsBase.Keys
        If Not keywordsPrev.exists(keyword) Then
            newSheet.Cells(outputRow, 1).Value = keywordsBase(keyword)
            newSheet.Cells(outputRow, 2).Value = keyword
            outputRow = outputRow + 1
        End If
    Next keyword

    Call ApplySheetFormatting(newSheet, True)
End Sub

' �� ���� �Ⱓ���� ������ ����� Ű���� ����
Sub ExtractRisingKeywordsBase(ws As Worksheet, lastRow As Long, basePeriod As String, comparePeriod As String)
    Dim newSheet As Worksheet, i As Long, keyword As Variant, outputRow As Long
    Dim keywordsCompare As Object, keywordsBase As Object
    Dim rankCompare As Long, rankBase As Long
    Dim sheetName As String

    Set keywordsCompare = CreateObject("Scripting.Dictionary")
    Set keywordsBase = CreateObject("Scripting.Dictionary")

    For i = 2 To lastRow
        If ws.Cells(i, 3).Value = comparePeriod Then
            keywordsCompare(ws.Cells(i, 2).Value) = ws.Cells(i, 1).Value
        End If
    Next i

    For i = 2 To lastRow
        If ws.Cells(i, 3).Value = basePeriod Then
            keywordsBase(ws.Cells(i, 2).Value) = ws.Cells(i, 1).Value
        End If
    Next i

    sheetName = comparePeriod & " ��������°˻���"
    Set newSheet = CreateResultSheet(sheetName)
    newSheet.Cells(1, 1).Value = basePeriod & "_����"
    newSheet.Cells(1, 2).Value = "�α�˻���"
    newSheet.Cells(1, 3).Value = "��������"
    newSheet.Cells(1, 4).Value = comparePeriod & "_����"
    outputRow = 2

    For Each keyword In keywordsBase.Keys
        If keywordsCompare.exists(keyword) Then
            rankBase = keywordsBase(keyword)
            rankCompare = keywordsCompare(keyword)
            If IsNumeric(rankBase) And IsNumeric(rankCompare) Then
                If rankBase < rankCompare Then
                    newSheet.Cells(outputRow, 1).Value = rankBase
                    newSheet.Cells(outputRow, 2).Value = keyword
                    newSheet.Cells(outputRow, 3).Value = rankCompare - rankBase
                    newSheet.Cells(outputRow, 4).Value = rankCompare
                    outputRow = outputRow + 1
                End If
            End If
        End If
    Next keyword

    Call ApplySheetFormatting(newSheet)
End Sub

' �� ���� �Ⱓ���� �ִ� �ű� Ű���� ���� (�񱳱Ⱓ 1�� ���)
Sub ExtractNewKeywordsVsPeriod(ws As Worksheet, lastRow As Long, basePeriod As String, comparePeriod As String)
    Dim newSheet As Worksheet, i As Long, keyword As Variant, outputRow As Long
    Dim keywordsCompare As Object, keywordsBase As Object
    Dim sheetName As String

    Set keywordsCompare = CreateObject("Scripting.Dictionary")
    Set keywordsBase = CreateObject("Scripting.Dictionary")

    For i = 2 To lastRow
        If ws.Cells(i, 3).Value = comparePeriod Then
            keywordsCompare(ws.Cells(i, 2).Value) = True
        End If
    Next i

    For i = 2 To lastRow
        If ws.Cells(i, 3).Value = basePeriod Then
            keywordsBase(ws.Cells(i, 2).Value) = ws.Cells(i, 1).Value
        End If
    Next i

    sheetName = comparePeriod & " ���ű԰˻���"
    Set newSheet = CreateResultSheet(sheetName)
    newSheet.Cells(1, 1).Value = "����"
    newSheet.Cells(1, 2).Value = "�α�˻���"
    outputRow = 2

    For Each keyword In keywordsBase.Keys
        If Not keywordsCompare.exists(keyword) Then
            newSheet.Cells(outputRow, 1).Value = keywordsBase(keyword)
            newSheet.Cells(outputRow, 2).Value = keyword
            outputRow = outputRow + 1
        End If
    Next keyword

    Call ApplySheetFormatting(newSheet)
End Sub

' �� ���� ��� �� �Ⱓ�� ���� ��� Ű���� ����
Sub ExtractRisingKeywordsVsPeriod(ws As Worksheet, lastRow As Long, basePeriod As String, comparePeriod As String)
    Dim newSheet As Worksheet, i As Long, keyword As Variant, outputRow As Long
    Dim keywordsCompare As Object, keywordsBase As Object
    Dim rankCompare As Long, rankBase As Long
    Dim sheetName As String

    Set keywordsCompare = CreateObject("Scripting.Dictionary")
    Set keywordsBase = CreateObject("Scripting.Dictionary")

    For i = 2 To lastRow
        If ws.Cells(i, 3).Value = comparePeriod Then
            keywordsCompare(ws.Cells(i, 2).Value) = ws.Cells(i, 1).Value
        End If
    Next i

    For i = 2 To lastRow
        If ws.Cells(i, 3).Value = basePeriod Then
            keywordsBase(ws.Cells(i, 2).Value) = ws.Cells(i, 1).Value
        End If
    Next i

    sheetName = comparePeriod & " ��������°˻���"
    Set newSheet = CreateResultSheet(sheetName)
    newSheet.Cells(1, 1).Value = basePeriod & "_����"
    newSheet.Cells(1, 2).Value = "�α�˻���"
    newSheet.Cells(1, 3).Value = "��������"
    newSheet.Cells(1, 4).Value = comparePeriod & "_����"
    outputRow = 2

    For Each keyword In keywordsBase.Keys
        If keywordsCompare.exists(keyword) Then
            rankBase = keywordsBase(keyword)
            rankCompare = keywordsCompare(keyword)
            If IsNumeric(rankBase) And IsNumeric(rankCompare) Then
                If rankBase < rankCompare Then
                    newSheet.Cells(outputRow, 1).Value = rankBase
                    newSheet.Cells(outputRow, 2).Value = keyword
                    newSheet.Cells(outputRow, 3).Value = rankCompare - rankBase
                    newSheet.Cells(outputRow, 4).Value = rankCompare
                    outputRow = outputRow + 1
                End If
            End If
        End If
    Next keyword

    Call ApplySheetFormatting(newSheet)
End Sub

' �� ��2 �� ��1 �� ���� ������ ������ ����� Ű���� ����
Sub ExtractRisingFromPastToCurrent(ws As Worksheet, lastRow As Long, basePeriod As String, comparePeriod1 As String, comparePeriod2 As String)
    Dim newSheet As Worksheet
    Dim i As Long, outputRow As Long
    Dim keyword As Variant
    Dim keywordsBase As Object, keywordsCompare1 As Object, keywordsCompare2 As Object
    Dim rankBase As Long, rank1 As Long, rank2 As Long
    Dim sheetName As String

    Set keywordsBase = CreateObject("Scripting.Dictionary")
    Set keywordsCompare1 = CreateObject("Scripting.Dictionary")
    Set keywordsCompare2 = CreateObject("Scripting.Dictionary")

    ' --- ���� ���� Ű����
    For i = 2 To lastRow
        If ws.Cells(i, 3).Value = basePeriod Then
            keywordsBase(ws.Cells(i, 2).Value) = ws.Cells(i, 1).Value
        End If
    Next i

    ' --- ��1
    For i = 2 To lastRow
        If ws.Cells(i, 3).Value = comparePeriod1 Then
            keywordsCompare1(ws.Cells(i, 2).Value) = ws.Cells(i, 1).Value
        End If
    Next i

    ' --- ��2
    For i = 2 To lastRow
        If ws.Cells(i, 3).Value = comparePeriod2 Then
            keywordsCompare2(ws.Cells(i, 2).Value) = ws.Cells(i, 1).Value
        End If
    Next i

    ' --- ��Ʈ ����
    sheetName = "���š�����_�������"
    Set newSheet = CreateResultSheet(sheetName)
    newSheet.Cells(1, 1).Value = comparePeriod2 & "_����"
    newSheet.Cells(1, 2).Value = comparePeriod1 & "_����"
    newSheet.Cells(1, 3).Value = basePeriod & "_����"
    newSheet.Cells(1, 4).Value = "�α�˻���"
    newSheet.Cells(1, 5).Value = "�ѻ����"
    outputRow = 2

    ' --- ������ ���š������ ������ ��� (���� �۾���)
    For Each keyword In keywordsBase.Keys
        If keywordsCompare1.exists(keyword) And keywordsCompare2.exists(keyword) Then
            rankBase = keywordsBase(keyword)
            rank1 = keywordsCompare1(keyword)
            rank2 = keywordsCompare2(keyword)
            If IsNumeric(rankBase) And IsNumeric(rank1) And IsNumeric(rank2) Then
                If rank2 > rank1 And rank1 > rankBase Then
                    newSheet.Cells(outputRow, 1).Value = rank2
                    newSheet.Cells(outputRow, 2).Value = rank1
                    newSheet.Cells(outputRow, 3).Value = rankBase
                    newSheet.Cells(outputRow, 4).Value = keyword
                    newSheet.Cells(outputRow, 5).Value = rank2 - rankBase
                    outputRow = outputRow + 1
                End If
            End If
        End If
    Next keyword

    Call ApplySheetFormatting(newSheet)
End Sub




' �� ��� ��Ʈ ����
Function CreateResultSheet(sheetName As String) As Worksheet
    Dim ws As Worksheet, cleanName As String
    cleanName = Replace(sheetName, "\", "")
    cleanName = Replace(cleanName, "/", "")
    cleanName = Replace(cleanName, "*", "")
    cleanName = Replace(cleanName, "[", "")
    cleanName = Replace(cleanName, "]", "")
    cleanName = Replace(cleanName, ":", "")
    cleanName = Replace(cleanName, "?", "")
    If Len(cleanName) > 31 Then cleanName = Left(cleanName, 31)

    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets(cleanName).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set ws = ActiveWorkbook.Sheets.Add
    ws.Name = cleanName
    Set CreateResultSheet = ws
End Function

' �� ��� ��Ʈ ���� �� ���� ���� + �� ���� ���� + �׵θ�
Sub ApplySheetFormatting(ws As Worksheet, Optional isNewKeywordSheet As Boolean = False)
    Dim lastRow As Long, lastCol As Long
    Dim dataRange As Range

    With ws
        ' ������ ��/�� ���
        lastRow = .Cells(.Rows.count, 1).End(xlUp).Row
        lastCol = .Cells(1, .Columns.count).End(xlToLeft).column

        If lastRow < 1 Or lastCol < 1 Then Exit Sub

        Set dataRange = .Range(.Cells(1, 1), .Cells(lastRow, lastCol))

        ' ���� �� ����
        With .Range(.Cells(1, 1), .Cells(1, lastCol))
            .Font.Bold = True
            .Interior.Color = RGB(197, 217, 241) ' �Ľ��� ���
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
        End With

        ' �ڵ� ���� ����
        .Range(.Cells(1, 1), .Cells(1, lastCol)).AutoFilter

        ' �� �ʺ� �ڵ� ����
        .Columns("A:" & Split(.Cells(1, lastCol).Address, "$")(1)).AutoFit

        ' �׵θ� ����
        With dataRange.Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .colorIndex = xlAutomatic
        End With

        ' �� �����űԾ� ��Ʈ�� �� ���� ������ �迭�� ����
        If isNewKeywordSheet Then
            .Tab.Color = RGB(255, 0, 0) ' �� ���� ������
        End If
    End With
End Sub


' �� ��� ���� ��Ʈ ���� �� �����۸�ũ ���� ����
Sub CreateFormattedSummaryReport()
    Dim wsSummary As Worksheet, ws As Worksheet
    Dim rowIndex As Long, keywordCount As Long
    Dim ���ؽ��� As String, ��1 As String, ��2 As String
    Dim periodList() As Variant, dictPeriods As Object
    Dim i As Long, lastRow As Long
    Dim wb As Workbook
    Dim ��ǥŰ���� As String

    Set wb = ActiveWorkbook
    Set dictPeriods = CreateObject("Scripting.Dictionary")
    Set ws = wb.Sheets("������")
    lastRow = ws.Cells(ws.Rows.count, "C").End(xlUp).Row

    For i = 2 To lastRow
        If Len(Trim(ws.Cells(i, 3).Value)) > 0 Then
            dictPeriods(Trim(ws.Cells(i, 3).Value)) = True
        End If
    Next i

    If dictPeriods.count < 3 Then
        MsgBox "�Ⱓ�� 3�� �̻� �������� �ʽ��ϴ�.", vbExclamation
        Exit Sub
    End If

    periodList = dictPeriods.Keys
    SortDescending periodList
    ���ؽ��� = periodList(0): ��1 = periodList(1): ��2 = periodList(2)

    ' ���� ��� ���� ����
    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets("��ຸ��").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set wsSummary = Worksheets.Add
    wsSummary.Name = "��ຸ��"

    ' ��� �񱳽��� ���̺� ����
    With wsSummary
        .Range("A1").Value = "�񱳽���"
        .Range("B1").Value = "�Ⱓ"
        .Range("A1:B1").Interior.Color = RGB(169, 209, 142)
        .Range("A1:B1").Font.Bold = True
        .Range("A1:B1").HorizontalAlignment = xlCenter
        .Range("A2").Value = "���ؽ���": .Range("B2").Value = ���ؽ���
        .Range("A3").Value = "1. �񱳽���": .Range("B3").Value = ��1
        .Range("A4").Value = "2. �񱳽���": .Range("B4").Value = ��2
        .Range("A1:A4").Borders.Weight = xlThin
        .Range("B1:B4").Borders.Weight = xlThin
    End With

    ' �м� ��� ���̺� ��� (5�� ����)
    rowIndex = 6
    With wsSummary
        .Range("A" & rowIndex).Value = "�м� �׸�"
        .Range("B" & rowIndex).Value = "Ű���� ��"
        .Range("C" & rowIndex).Value = "��ǥ Ű����"
        .Range("D" & rowIndex).Value = "�м� ��Ʈ�� �̵�"
        .Range("E" & rowIndex).Value = "���"
        .Range("A" & rowIndex & ":E" & rowIndex).Interior.Color = RGB(244, 176, 132)
        .Range("A" & rowIndex & ":E" & rowIndex).Font.Bold = True
        .Range("A" & rowIndex & ":E" & rowIndex).HorizontalAlignment = xlCenter
    End With

    ' ��� ��Ʈ ��ȸ �� ������� �Է�
    rowIndex = rowIndex + 1
    For Each ws In wb.Worksheets
        If ws.Name <> "������" And ws.Name <> "��ຸ��" Then
            keywordCount = Application.WorksheetFunction.CountA(ws.Range("A:A")) - 1
            
            ' ��ǥ Ű����: �Ϲ� ��Ʈ�� B2, ���š�����_������� ��Ʈ�� D2
            If keywordCount > 0 Then
                If ws.Name = "���š�����_�������" Then
                    ��ǥŰ���� = ws.Range("D2").Value
                Else
                    ��ǥŰ���� = ws.Range("B2").Value
                End If
                If Len(��ǥŰ����) = 0 Then ��ǥŰ���� = "(��ǥ Ű���� ����)"
            Else
                ��ǥŰ���� = "(������ ����)"
            End If

            
            ' �� ����
            wsSummary.Cells(rowIndex, 1).Value = ws.Name
            wsSummary.Cells(rowIndex, 2).Value = keywordCount
            wsSummary.Cells(rowIndex, 3).Value = ��ǥŰ����
            wsSummary.Hyperlinks.Add Anchor:=wsSummary.Cells(rowIndex, 4), Address:="", _
                SubAddress:="'" & ws.Name & "'!A1", TextToDisplay:="�̵�"

            ' ��� ����
            Select Case True
                Case InStr(ws.Name, "�����ű�") > 0
                    wsSummary.Cells(rowIndex, 5).Value = "���ؽ����� ���� ���Ӱ� ������ Ű����"
                Case InStr(ws.Name, "���ű�") > 0
                    If InStr(ws.Name, ��1) > 0 Then
                        wsSummary.Cells(rowIndex, 5).Value = "1. �񱳽��� ��� ���ؽ��� �ű� Ű����"
                    ElseIf InStr(ws.Name, ��2) > 0 Then
                        wsSummary.Cells(rowIndex, 5).Value = "2. �񱳽��� ��� ���ؽ��� �ű� Ű����"
                    End If
                Case InStr(ws.Name, "���������") > 0
                    If InStr(ws.Name, ��1) > 0 Then
                        wsSummary.Cells(rowIndex, 5).Value = "1. �񱳽��� ��� ���� ���"
                    ElseIf InStr(ws.Name, ��2) > 0 Then
                        wsSummary.Cells(rowIndex, 5).Value = "2. �񱳽��� ��� ���� ���"
                    End If
                    
                Case ws.Name = "���š�����_�������"
                        wsSummary.Cells(rowIndex, 5).Value = "���� �� ����� ������ ������ ����� Ű����"
            End Select

            rowIndex = rowIndex + 1
        End If
    Next ws

    ' �� �ʺ� �ڵ�����
    wsSummary.Columns("A:E").AutoFit
End Sub


