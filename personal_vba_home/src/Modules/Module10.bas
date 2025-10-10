Attribute VB_Name = "Module10"
Sub 볼드체처리하기()
    Dim ws As Worksheet
    Dim cell As Range
    Dim startPos As Integer
    Dim textLength As Integer
    Dim searchText As String
    Dim boldText As String

    ' 현재 활성화된 시트 선택
    Set ws = ActiveSheet

    ' 찾을 텍스트와 볼드 처리할 텍스트 설정
    searchText = "<b>"
    boldText = "</b>"

    ' 시트의 모든 셀을 순회하며 검색
    For Each cell In ws.UsedRange
        If InStr(cell.Value, searchText) > 0 And InStr(cell.Value, boldText) > 0 Then
            ' 텍스트 시작과 끝 위치 계산
            startPos = InStr(cell.Value, searchText) + Len(searchText)
            textLength = InStr(cell.Value, boldText) - startPos
            
            ' 텍스트에서 <b>와 </b> 제거
            cell.Value = Replace(cell.Value, searchText, "")
            cell.Value = Replace(cell.Value, boldText, "")

            ' 텍스트를 볼드 처리
            With cell.Characters(Start:=startPos - Len(searchText), Length:=textLength)
                .Font.Bold = True
            End With
        End If
    Next cell
End Sub

