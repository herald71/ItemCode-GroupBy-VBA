Attribute VB_Name = "Module11"
Sub 단어검색후색상으로강조하기()
    Dim rng As Range
    Dim cell As Range
    Dim searchText As String
    Dim colorText As String
    Dim startPos As Integer
    Dim count As Long ' 검색된 단어의 수를 세기 위한 변수
    Dim searchTextLength As Long
    Dim fontColor As Long

    ' 사용자로부터 검색할 단어와 색상을 입력받음
    searchText = InputBox("색상 굵은체로 표시할 단어를 입력하세요:", "검색어 입력")
    colorText = InputBox("글자 색상을 입력하세요 (예: Red 또는 255,0,0):", "색상 입력")
    
    ' 입력값이 비어있지 않은 경우에만 실행
    If searchText <> "" And colorText <> "" Then
        searchText = LCase(searchText) ' 사용자 입력을 소문자로 변환
        searchTextLength = Len(searchText) ' 검색 텍스트의 길이 계산
        fontColor = ColorToRGB(colorText) ' 입력받은 색상을 RGB 값으로 변환
        Set rng = ActiveSheet.UsedRange ' 활성 시트의 사용된 범위를 설정
        count = 0 ' 카운터 초기화
        
        For Each cell In rng
            If Not IsError(cell.Value) And Not IsEmpty(cell.Value) Then
                If VarType(cell.Value) = vbString Then ' 셀 값이 문자열인 경우에만 처리
                    If InStr(1, LCase(cell.Value), searchText, vbTextCompare) > 0 Then
                        startPos = InStr(1, LCase(cell.Value), searchText, vbTextCompare)
                        While startPos > 0
                            With cell.Characters(startPos, searchTextLength).Font
                                .Bold = True
                                .Color = fontColor ' 사용자가 지정한 색상으로 설정
                            End With
                            startPos = InStr(startPos + searchTextLength, LCase(cell.Value), searchText, vbTextCompare)
                            count = count + 1
                        Wend
                    End If
                End If
            End If
        Next cell
        
        ' 결과 메시지 표시
        If count > 0 Then
            MsgBox count & "개의 셀에서 '" & searchText & "'(이)가 검색되어 지정한 색상으로 굵게 표시되었습니다.", vbInformation, "검색 결과"
        Else
            MsgBox "'" & searchText & "'(을)를 포함하는 셀이 없습니다.", vbInformation, "검색 결과"
        End If
    Else
        MsgBox "검색할 단어 또는 색상이 입력되지 않았습니다.", vbExclamation, "입력 오류"
    End If
End Sub

Function ColorToRGB(colorText As String) As Long
    Dim colorParts() As String
    If InStr(colorText, ",") > 0 Then
        colorParts = Split(colorText, ",")
        If UBound(colorParts) = 2 Then
            ColorToRGB = RGB(val(colorParts(0)), val(colorParts(1)), val(colorParts(2)))
        Else
            ColorToRGB = RGB(255, 0, 0) ' 기본값은 빨간색
        End If
    Else
        Select Case LCase(colorText)
            Case "red"
                ColorToRGB = RGB(255, 0, 0)
            Case "green"
                ColorToRGB = RGB(0, 128, 0)
            Case "blue"
                ColorToRGB = RGB(0, 0, 255)
            Case Else
                ColorToRGB = RGB(0, 0, 0) ' 기본값은 검정색
        End Select
    End If
End Function

