Attribute VB_Name = "Module2"
Sub 네이버쇼핑_카테고리검색_데이타정리()
    ' 모든 작업을 순서대로 실행
    Call RenameHeadersInActiveSheet
    Call DeleteOtherColumns
    Call ExtractAndSaveNumbers
    Call FormatAndSaveDatesCorrectly
End Sub

Sub RenameHeadersInActiveSheet()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    With ws
        .Cells(1, 2).Value = "상품URL"
        .Cells(1, 3).Value = "썸네일URL"
        .Cells(1, 7).Value = "상품명"
        .Cells(1, 9).Value = "판매가격"
        .Cells(1, 10).Value = "배송비"
        .Cells(1, 12).Value = "대분류"
        .Cells(1, 13).Value = "소분류"
        .Cells(1, 14).Value = "중분류"
        .Cells(1, 15).Value = "세부항목"
        .Cells(1, 23).Value = "리뷰수"
        .Cells(1, 24).Value = "판매처"
        .Cells(1, 34).Value = "상품등록일"
        .Cells(1, 35).Value = "찜하기수"
    End With
End Sub

Sub DeleteOtherColumns()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Dim arrKeepCols As Variant
    Dim dictKeepCols As Object
    Dim i As Long
    Dim maxCol As Long

    arrKeepCols = Array(2, 3, 7, 9, 10, 12, 13, 14, 15, 23, 24, 34, 35) ' B, C, G, I, J, L, M, N, O, W, X, AH, AI
    Set dictKeepCols = CreateObject("Scripting.Dictionary")

    For i = LBound(arrKeepCols) To UBound(arrKeepCols)
        dictKeepCols.Add arrKeepCols(i), Nothing
    Next i

    maxCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).column

    For i = maxCol To 1 Step -1
        If Not dictKeepCols.exists(i) Then
            ws.Columns(i).Delete
        End If
    Next i
End Sub

Sub ExtractAndSaveNumbers()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Dim cell As Range
    Dim arrCols As Variant
    arrCols = Array("D", "E", "J", "M")  ' 검사할 열 지정
    Dim col As Variant
    Dim i As Long
    Dim cleanValue As String

    For Each col In arrCols
        For i = 2 To ws.Cells(ws.Rows.count, col).End(xlUp).Row
            Set cell = ws.Cells(i, col)
            cleanValue = ExtractNumbers(cell.Value)
            If cleanValue <> "" Then
                cell.Value = CDbl(cleanValue)
            Else
                cell.ClearContents
            End If
        Next i
    Next col
End Sub

Function ExtractNumbers(str As String) As String
    Dim output As String
    Dim pos As Integer

    For pos = 1 To Len(str)
        If Mid(str, pos, 1) Like "[0-9]" Then
            output = output & Mid(str, pos, 1)
        End If
    Next pos

    ExtractNumbers = output
End Function

Sub FormatAndSaveDatesCorrectly()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Dim cell As Range
    Dim lastRow As Long
    Dim dateStr As String
    Dim yearPart As Integer
    Dim monthPart As Integer

    lastRow = ws.Cells(ws.Rows.count, "L").End(xlUp).Row

    For Each cell In ws.Range("L2:L" & lastRow)
        dateStr = Trim(Replace(cell.Value, "등록일", ""))

        If dateStr Like "####.##.*" Then
            yearPart = val(Left(dateStr, 4))
            monthPart = val(Mid(dateStr, 6, 2))

            cell.Value = DateSerial(yearPart, monthPart, 1)
            cell.NumberFormat = "YYYY.MM"
        Else
            cell.ClearContents
        End If
    Next cell
End Sub

