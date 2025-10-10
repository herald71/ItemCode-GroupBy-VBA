Attribute VB_Name = "Module3"
Sub 열값을입력받아숫자로전환()
Attribute 열값을입력받아숫자로전환.VB_Description = "열값을 입력받아 숫자만 추출"
Attribute 열값을입력받아숫자로전환.VB_ProcData.VB_Invoke_Func = "t\n14"
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Dim cell As Range
    Dim arrCols As Variant
    Dim col As Variant
    Dim i As Long
    Dim cleanValue As String
    Dim colInput As String
    
    ' 사용자로부터 검사할 열을 입력 받습니다.
    colInput = InputBox("검사할 열을 입력하세요. 열은 쉼표(,)로 구분하세요. 예: D,E,J,M")
    If colInput = "" Then Exit Sub ' 사용자가 입력을 취소하면 코드를 종료합니다.
    
    ' 입력받은 열을 배열로 변환합니다.
    arrCols = Split(colInput, ",")
    
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

