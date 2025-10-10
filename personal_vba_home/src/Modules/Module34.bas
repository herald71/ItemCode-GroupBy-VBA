Attribute VB_Name = "Module34"
Sub RunKeywordAnalysis_MultiYear()
    ' ----------------------------------------------------------
    ' ���α׷��� : RunKeywordAnalysis_MultiYear
    ' ����       : ���� �� ���� ���� ���� ���/�϶�, �ű�, �����, �ֽŻ�� Ű���� �м�
    ' �ۼ�����   : 2025-07-02
    ' ��������   : �м� ��Ʈ�� "�Ϸù�ȣ" �� �߰� �� �ڵ� ä��, �ֽŻ��Ű���� ��Ʈ �߰�
    ' ----------------------------------------------------------

    Dim wsData As Worksheet: Set wsData = ActiveWorkbook.Sheets("������")
    Dim lastRow As Long: lastRow = wsData.Cells(wsData.Rows.count, "B").End(xlUp).Row

    Dim rankData As Object: Set rankData = CreateObject("Scripting.Dictionary")
    Dim allYears As Object: Set allYears = CreateObject("Scripting.Dictionary")

    Dim r As Long, keyword As Variant, yearStr As String, year As Integer, rank As Long
    Dim data As Object    ' �ڡڡ� data ������ ���⼭ �� ���� ����! �ڡڡ�

    ' ------------------------------
    ' 1. ������ ����
    ' ------------------------------
    For r = 2 To lastRow
        keyword = Trim(wsData.Cells(r, "B").Value)
        yearStr = wsData.Cells(r, "C").Value
        If Len(yearStr) >= 4 Then year = val(Left(yearStr, 4)) Else year = 0
        rank = wsData.Cells(r, "A").Value

        If year > 0 Then
            If Not rankData.exists(keyword) Then
                Set rankData(keyword) = CreateObject("Scripting.Dictionary")
            End If
            rankData(keyword)(year) = rank
            allYears(year) = True
        End If
    Next r

    ' ���� ����
    Dim sortedYears() As Long
    ReDim sortedYears(0 To allYears.count - 1)
    Dim i As Integer: i = 0
    Dim y As Variant
    For Each y In allYears.Keys
        sortedYears(i) = y
        i = i + 1
    Next y
    Call QuickSortLong(sortedYears, LBound(sortedYears), UBound(sortedYears))
    Dim latestYear As Long: latestYear = sortedYears(UBound(sortedYears))

    ' ------------------------------
    ' 2. ���� ��Ʈ ����
    ' ------------------------------
    Dim shtNames: shtNames = Array("�������Ű����", "�����϶�Ű����", "�����Ű����", "�ű�Ű����", "Ű����м�_��ຸ��", "�ֱٻ��Ű����")
    Application.DisplayAlerts = False
    For Each y In shtNames
        On Error Resume Next: Sheets(y).Delete: On Error GoTo 0
    Next y
    Application.DisplayAlerts = True

    ' ------------------------------
    ' 3. ��� ��Ʈ ����
    ' ------------------------------
    Dim wsUp As Worksheet: Set wsUp = Sheets.Add: wsUp.Name = "�������Ű����"
    Dim wsDown As Worksheet: Set wsDown = Sheets.Add: wsDown.Name = "�����϶�Ű����"
    Dim wsGone As Worksheet: Set wsGone = Sheets.Add: wsGone.Name = "�����Ű����"
    Dim wsNew As Worksheet: Set wsNew = Sheets.Add: wsNew.Name = "�ű�Ű����"

    ' ------------------------------
    ' 3-1. �Ϸù�ȣ �� �߰� �� ���� �ۼ�
    ' ------------------------------
    Dim wsList As Variant: wsList = Array(wsUp, wsDown, wsGone, wsNew)
    For i = 0 To 3
        With wsList(i)
            .Columns("A:A").Insert Shift:=xlToRight
            .Cells(1, 1).Value = "�Ϸù�ȣ"
        End With
    Next i

    wsUp.Cells(1, 2).Value = "�α�˻���"
    wsDown.Cells(1, 2).Value = "�α�˻���"
    For i = 0 To UBound(sortedYears)
        wsUp.Cells(1, i + 3).Value = sortedYears(i) & " ����"
        wsDown.Cells(1, i + 3).Value = sortedYears(i) & " ����"
    Next i
    wsUp.Cells(1, i + 3).Value = "���� ������"
    wsDown.Cells(1, i + 3).Value = "���� �϶���"

    wsGone.Range("B1:C1").Value = Array("�α�˻���", "������ ����⵵")
    wsNew.Range("B1:C1").Value = Array("�ű� Ű����", latestYear & " ����")

    ' ------------------------------
    ' 4. �м�
    ' ------------------------------
    Dim iUp As Long: iUp = 2
    Dim iDown As Long: iDown = 2
    Dim iGone As Long: iGone = 2
    Dim iNew As Long: iNew = 2

    Dim goneDict As Object: Set goneDict = CreateObject("Scripting.Dictionary")

    For Each keyword In rankData.Keys
        Set data = rankData(keyword)
        Dim available() As Long: ReDim available(0 To data.count - 1)
        i = 0
        For Each y In sortedYears
            If data.exists(y) Then
                available(i) = y
                i = i + 1
            End If
        Next y

        If i >= 3 Then
            Dim upFlag As Boolean: upFlag = True
            Dim downFlag As Boolean: downFlag = True
            Dim j As Integer
            For j = 1 To i - 1
                If data(available(j - 1)) <= data(available(j)) Then upFlag = False
                If data(available(j - 1)) >= data(available(j)) Then downFlag = False
            Next j
            If upFlag Then
                wsUp.Cells(iUp, 1).Value = iUp - 1
                wsUp.Cells(iUp, 2).Value = keyword
                For j = 0 To UBound(sortedYears)
                    If data.exists(sortedYears(j)) Then
                        wsUp.Cells(iUp, j + 3).Value = data(sortedYears(j))
                    End If
                Next j
                wsUp.Cells(iUp, j + 3).Value = data(available(0)) - data(available(i - 1))
                iUp = iUp + 1
            ElseIf downFlag Then
                wsDown.Cells(iDown, 1).Value = iDown - 1
                wsDown.Cells(iDown, 2).Value = keyword
                For j = 0 To UBound(sortedYears)
                    If data.exists(sortedYears(j)) Then
                        wsDown.Cells(iDown, j + 3).Value = data(sortedYears(j))
                    End If
                Next j
                wsDown.Cells(iDown, j + 3).Value = data(available(i - 1)) - data(available(0))
                iDown = iDown + 1
            End If
        End If

        If Not data.exists(latestYear) Then
            Dim maxY As Long: maxY = 0
            For Each y In data.Keys
                If y > maxY Then maxY = y
            Next y
            goneDict(keyword) = maxY
        End If

        If data.count = 1 And data.exists(latestYear) Then
            wsNew.Cells(iNew, 1).Value = iNew - 1
            wsNew.Cells(iNew, 2).Value = keyword
            wsNew.Cells(iNew, 3).Value = data(latestYear)
            iNew = iNew + 1
        End If
    Next keyword

    For Each keyword In goneDict.Keys
        wsGone.Cells(iGone, 1).Value = iGone - 1
        wsGone.Cells(iGone, 2).Value = keyword
        wsGone.Cells(iGone, 3).Value = goneDict(keyword)
        iGone = iGone + 1
    Next keyword

    ' ------------------------------
    ' 4-1. �ֽų⵵�� �� �� ���̶� ������ ���� Ű���� ��Ʈ
    ' ------------------------------
    Dim wsRecentUp As Worksheet: Set wsRecentUp = Sheets.Add: wsRecentUp.Name = "�ֱٻ��Ű����"
    wsRecentUp.Cells(1, 1).Value = "�Ϸù�ȣ"
    wsRecentUp.Cells(1, 2).Value = "�α�˻���"
    wsRecentUp.Cells(1, 3).Value = "���� �ְ� ����(����)"
    wsRecentUp.Cells(1, 4).Value = latestYear & " ����"
    wsRecentUp.Cells(1, 5).Value = "���� ������"

    Dim iRecentUp As Long: iRecentUp = 2

    For Each keyword In rankData.Keys
        Set data = rankData(keyword)
        If data.exists(latestYear) And data.count > 1 Then
            Dim bestOldRank As Variant: bestOldRank = 1000000
            For Each y In data.Keys
                If y <> latestYear Then
                    If data(y) < bestOldRank Then bestOldRank = data(y)
                End If
            Next y
            If data(latestYear) < bestOldRank Then
                wsRecentUp.Cells(iRecentUp, 1).Value = iRecentUp - 1
                wsRecentUp.Cells(iRecentUp, 2).Value = keyword
                wsRecentUp.Cells(iRecentUp, 3).Value = bestOldRank
                wsRecentUp.Cells(iRecentUp, 4).Value = data(latestYear)
                wsRecentUp.Cells(iRecentUp, 5).Value = bestOldRank - data(latestYear)
                iRecentUp = iRecentUp + 1
            End If
        End If
    Next keyword

    ' ���� ���� ����
    With wsRecentUp.Range("A1:E1")
        .Font.Bold = True
        .Interior.Color = RGB(204, 255, 229)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    wsRecentUp.Columns.AutoFit
    wsRecentUp.Range("A1:E" & iRecentUp - 1).Borders.LineStyle = xlContinuous

    ' ------------------------------
    ' 5. ��� ����
    ' ------------------------------
    Dim wsReport As Worksheet: Set wsReport = Sheets.Add: wsReport.Name = "Ű����м�_��ຸ��"
    wsReport.Range("A1").Value = "�� �ٳⰣ ���̹� ����Ÿ�� Ű���� �м� ��� ����"
    wsReport.Range("A1").Font.Bold = True
    wsReport.Range("A1").Font.Size = 14
    wsReport.Range("A3:E3").Value = Array("�׸�", "Ű���� ��", "��ǥ Ű����", "�м� ��Ʈ�� �̵�", "���")
    wsReport.Range("A3:E3").Font.Bold = True

    Dim counts(1 To 5) As Long, examples(1 To 5) As String
    Dim snames: snames = Array("�������Ű����", "�����϶�Ű����", "�����Ű����", "�ű�Ű����", "�ֱٻ��Ű����")
    For i = 0 To 4
        With Worksheets(snames(i))
            counts(i + 1) = .Cells(.Rows.count, "B").End(xlUp).Row - 1
            If counts(i + 1) > 0 Then
                examples(i + 1) = .Cells(2, 2).Value
            Else
                examples(i + 1) = "(������ ����)"
            End If
        End With
    Next i

    For i = 0 To 4
        wsReport.Cells(i + 4, 1).Value = Choose(i + 1, "���� ��� Ű����", "���� �϶� Ű����", "����� Ű����", "�ű� Ű����", "�ֱ� ��� Ű����")
        wsReport.Cells(i + 4, 2).Value = counts(i + 1)
        wsReport.Cells(i + 4, 3).Value = examples(i + 1)
        wsReport.Hyperlinks.Add Anchor:=wsReport.Cells(i + 4, 4), _
            Address:="", SubAddress:="'" & snames(i) & "'!A1", _
            TextToDisplay:="�̵�"
        wsReport.Cells(i + 4, 5).Value = ""
    Next i

    With wsReport.Range("A3:E8")
        .Columns.AutoFit
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' ------------------------------
    ' 6. ��Ʈ ���� ���� (���� + ��ü �׵θ� + �ڵ� �ʺ�)
    ' ------------------------------
    Dim pastelColors As Variant
    pastelColors = Array(RGB(204, 255, 229), RGB(255, 230, 255), RGB(255, 255, 204), RGB(221, 235, 247), RGB(204, 255, 229)) ' 5�� ��

    Dim wsTitles As Variant
    wsTitles = Array(wsUp, wsDown, wsGone, wsNew, wsRecentUp)

    Dim colCount As Long, rowCount As Long
    For i = 0 To 4
        With wsTitles(i)
            colCount = .Cells(1, .Columns.count).End(xlToLeft).column
            rowCount = .Cells(.Rows.count, "A").End(xlUp).Row

            ' ���� �� ����
            With .Range(.Cells(1, 1), .Cells(1, colCount))
                .Font.Bold = True
                .Interior.Color = pastelColors(i)
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With

            ' ��ü �� �׵θ�
            With .Range(.Cells(1, 1), .Cells(rowCount, colCount))
                .Borders.LineStyle = xlContinuous
            End With

            .Columns.AutoFit
        End With
    Next i

    MsgBox "�� ���� ���� ���� �м� �� ���� ���� �Ϸ�!", vbInformation
End Sub

' ���� ���Ŀ� ����Ʈ �Լ�
Sub QuickSortLong(arr() As Long, ByVal first As Long, ByVal last As Long)
    Dim low As Long, high As Long, pivot As Long, temp As Long
    low = first: high = last
    pivot = arr((first + last) \ 2)
    Do While low <= high
        Do While arr(low) < pivot: low = low + 1: Loop
        Do While arr(high) > pivot: high = high - 1: Loop
        If low <= high Then
            temp = arr(low): arr(low) = arr(high): arr(high) = temp
            low = low + 1: high = high - 1
        End If
    Loop
    If first < high Then QuickSortLong arr, first, high
    If low < last Then QuickSortLong arr, low, last
End Sub


