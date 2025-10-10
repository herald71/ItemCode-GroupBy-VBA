Attribute VB_Name = "Module37"
Sub CopyColoredCellsFromColumnB()
    ' --------------------------------------------------------
    ' ���α׷��� : B���� ä��� �� �ִ� ���� B���˻��� ��Ʈ�� ����
    ' �ۼ�����   : 2025-07-11
    ' ����       : �ֿ� Ű���� ��Ʈ�鿡�� B���� ä��� �� �ִ� ���� �����Ͽ�
    '              'B���˻���' ��Ʈ�� �Ϸù�ȣ, ��, ��Ʈ������ ����
    '              ���� ������ ������ ������ ����
    ' --------------------------------------------------------

    Dim destSheet As Worksheet
    Dim ws As Worksheet
    Dim cell As Range
    Dim r As Long
    Dim lastRow As Long
    Dim destRow As Long: destRow = 2
    Dim wb As Workbook
    Set wb = ActiveWorkbook

    Application.ScreenUpdating = False

    ' ���� ��Ʈ �ʱ�ȭ �Ǵ� ����
    On Error Resume Next
    Set destSheet = wb.Sheets("B���˻���")
    If destSheet Is Nothing Then
        Set destSheet = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.count))
        destSheet.Name = "B���˻���"
    Else
        destSheet.Cells.Clear
    End If
    On Error GoTo 0

    ' ���� �Է�
    With destSheet
        .Cells(1, 1).Value = "�Ϸù�ȣ"
        .Cells(1, 2).Value = "�α�˻���"
        .Cells(1, 3).Value = "������ ��Ʈ"
        .Rows(1).Font.Bold = True
    End With

    ' ��� ��Ʈ ��ȸ
    For Each ws In wb.Worksheets
        ' ��� ��Ʈ ����
        If ws.Name <> destSheet.Name Then
            With ws
                lastRow = .Cells(.Rows.count, "B").End(xlUp).Row
                For r = 2 To lastRow
                    Set cell = .Cells(r, "B")
                    If cell.Interior.colorIndex <> xlNone And cell.Interior.colorIndex <> -4142 Then
                        If Trim(cell.Value) <> "" Then
                            destSheet.Cells(destRow, 1).Value = destRow - 1        ' �Ϸù�ȣ
                            destSheet.Cells(destRow, 2).Value = cell.Value         ' �α�˻���
                            destSheet.Cells(destRow, 3).Value = ws.Name            ' ��Ʈ��
                            destRow = destRow + 1
                        End If
                    End If
                Next r
            End With
        End If
    Next ws

    destSheet.Columns("A:C").AutoFit

    Application.ScreenUpdating = True

    MsgBox "B���˻��� ��Ʈ�� ���� �Ϸ�! �� " & destRow - 2 & "�� �׸�", vbInformation

End Sub

