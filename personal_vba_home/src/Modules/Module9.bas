Attribute VB_Name = "Module9"

Sub 진텍스상품바코드정리()
    Dim ws As Worksheet
    Set ws = Sheets("data")
    
    ' 1. "data" 시트의 1,2행을 삭제합니다.
    ws.Rows("1:2").Delete
    
' 2. E 열에 필드값이 없는 항목을 찾아서 그 행의 모든 데이터를 삭제합니다.
    Dim lastRow As Long, i As Long
    lastRow = ws.Cells(ws.Rows.count, "E").End(xlUp).Row

    For i = lastRow To 1 Step -1
            If Trim(ws.Cells(i, "E").Value) = "" Then
            ws.Rows(i).Delete
    End If
    Next i

    
    ' 3. J열부터 S열까지 모든 데이터 삭제합니다.
    ws.Columns("J:S").Delete Shift:=xlToLeft
    
    ' 4. G, F, D 열을 순서대로 삭제합니다.
    ws.Columns("G").Delete Shift:=xlToLeft
    ws.Columns("F").Delete Shift:=xlToLeft
    ws.Columns("D").Delete Shift:=xlToLeft
    
    ' 5. A, B 열의 데이터를 서로 바꿉니다. 임시 열을 사용하는 방식으로 수정합니다.
    ws.Columns("A").EntireColumn.Copy
    ws.Columns("XFD").EntireColumn.PasteSpecial Paste:=xlPasteValues
    ws.Columns("B").EntireColumn.Copy
    ws.Columns("A").EntireColumn.PasteSpecial Paste:=xlPasteValues
    ws.Columns("XFD").EntireColumn.Copy
    ws.Columns("B").EntireColumn.PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    ' 임시로 사용된 XFD 열을 삭제합니다.
        ws.Columns("XFD").Delete
    
    ' 6. A열 앞에 빈 열을 삽입합니다.
     ws.Columns("A").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
End Sub


