Attribute VB_Name = "Module32"
Sub ColorRowsByColumnGroup()
    ' ---------------------------------------------------------
    ' 프로그램명 : ColorRowsByColumnGroup
    ' 한글설명   : 지정된 열의 값을 기준으로 행 색상을 그룹별로 구분
    ' 작성일자   : 2025-06-28
    ' 설명       : 현재 활성화된 워크시트에서 사용자가 입력한 열의 값을 기준으로
    '              같은 값을 가진 행들에 연한 파스텔톤 색상을 순차적으로 적용하여
    '              시각적으로 그룹을 구분할 수 있도록 도와주는 매크로입니다.
    ' ---------------------------------------------------------

    Dim ws As Worksheet
    Set ws = ActiveSheet  ' 현재 활성화된 시트 기준으로 실행

    Dim colInput As Variant
    Dim colNumber As Long

    ' 사용자에게 열 문자 입력 요청 (예: F)
    colInput = Application.InputBox( _
        prompt:="기준이 될 열을 입력하세요 (예: F)", _
        Title:="색상 구분 기준 열 입력", Type:=2)

    If colInput = False Then Exit Sub ' 사용자가 취소했을 경우 종료

    ' 열 문자를 열 번호로 변환
    On Error GoTo InvalidColumn
    colNumber = Range(colInput & "1").column
    On Error GoTo 0

    ' 지정 열의 마지막 데이터 행 번호 찾기
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, colNumber).End(xlUp).Row

    ' Dictionary 객체 생성 (값별 색상 저장용)
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim i As Long, key As String
    Dim colorList As Variant
    Dim colorIndex As Long

    ' 연한 파스텔톤 RGB 색상 리스트
    colorList = Array( _
        RGB(255, 235, 238), RGB(232, 245, 233), RGB(232, 234, 246), _
        RGB(255, 249, 196), RGB(224, 242, 241), RGB(241, 248, 233), _
        RGB(248, 237, 227), RGB(237, 231, 246), RGB(255, 236, 179), _
        RGB(225, 245, 254), RGB(240, 244, 195), RGB(255, 245, 238) _
    )

    colorIndex = 0

    ' 2행부터 마지막 행까지 반복하며 색상 적용
    For i = 2 To lastRow
        key = ws.Cells(i, colNumber).Value

        If Not dict.exists(key) Then
            dict.Add key, colorList(colorIndex Mod UBound(colorList) + 1)
            colorIndex = colorIndex + 1
        End If

        ws.Rows(i).Interior.Color = dict(key)
    Next i

    Exit Sub

InvalidColumn:
    MsgBox "입력한 열이 유효하지 않습니다. 다시 확인해 주세요.", vbExclamation
End Sub


