Attribute VB_Name = "Module26"
Sub 필터링_및_복사()
    Dim ws As Worksheet
    Dim newWs As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim filterRange As Range
    Dim copyRange As Range
    Dim wsName As String
    Dim i As Integer
    Dim userResponse As Variant
    Dim searchMonths As String
    Dim monthArray() As String
    Dim minVal As Double, maxVal As Double
    Dim filterSummary As String
    Dim recentSearchRange As String
    Dim coupangPriceRange As String
    Dim coupangReviewRange As String
    Dim coupangRocketRatio As String
    Dim coupangSellerRocketRatio As String
    
    ' 현재 선택된 시트
    Set ws = ActiveSheet

    ' 마지막 행과 열 찾기
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).column

    ' 데이터가 없는 경우 종료
    If lastRow < 2 Then
        MsgBox "데이터가 없습니다. 실행을 종료합니다.", vbExclamation
        Exit Sub
    End If

    ' 기존 AutoFilter 제거
    If ws.AutoFilterMode Then ws.AutoFilterMode = False

    ' 필터를 적용할 전체 범위 설정
    Set filterRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
    
    ' AutoFilter 활성화
    filterRange.AutoFilter

    ' 필터링 요약 초기화
    filterSummary = "필터링 적용 요약:" & vbLf

    ' 브랜드 키워드(O) 제외 여부 확인
    userResponse = MsgBox("브랜드 키워드가 'O'인 항목을 제외하시겠습니까?", vbYesNo, "필터링 옵션")
    If userResponse = vbYes Then
        filterRange.AutoFilter Field:=4, Criteria1:="<>O"
        filterSummary = filterSummary & "- 브랜드 키워드(O) 제외" & vbLf
    End If

    ' 쇼핑성 키워드(X) 제외
    filterRange.AutoFilter Field:=5, Criteria1:="<>X"
    filterSummary = filterSummary & "- 쇼핑성 키워드(X) 제외" & vbLf

    ' 최근 1개월 검색량 필터링
    recentSearchRange = InputBox("최근 1개월 검색량 범위를 입력하세요 (예: 1000~100000)")
    If recentSearchRange <> "" And IsValidRange(recentSearchRange) Then
        minVal = Split(recentSearchRange, "~")(0)
        maxVal = Split(recentSearchRange, "~")(1)
        filterRange.AutoFilter Field:=7, Criteria1:=">=" & minVal, Criteria2:="<=" & maxVal
        filterSummary = filterSummary & "- 최근 1개월 검색량: " & recentSearchRange & vbLf
    End If

    ' 작년 최대 검색 월 필터링
    searchMonths = InputBox("작년 최대 검색 월을 입력하세요 (예: 4,5,6,7,8)")
    If searchMonths <> "" And IsValidMonthList(searchMonths) Then
        monthArray = Split(searchMonths, ",")
        filterRange.AutoFilter Field:=14, Criteria1:=monthArray, Operator:=xlFilterValues
        filterSummary = filterSummary & "- 작년 최대 검색 월: " & searchMonths & vbLf
    End If

    ' 쿠팡 평균가 필터링
    coupangPriceRange = InputBox("쿠팡 평균가 범위를 입력하세요 (예: 9800~29999)")
    If coupangPriceRange <> "" And IsValidRange(coupangPriceRange) Then
        minVal = Split(coupangPriceRange, "~")(0)
        maxVal = Split(coupangPriceRange, "~")(1)
        filterRange.AutoFilter Field:=26, Criteria1:=">=" & minVal, Criteria2:="<=" & maxVal
        filterSummary = filterSummary & "- 쿠팡 평균가: " & coupangPriceRange & vbLf
    End If

    ' 쿠팡 평균리뷰수 필터링
    coupangReviewRange = InputBox("쿠팡 평균리뷰수 범위를 입력하세요 (예: 0~200)")
    If coupangReviewRange <> "" And IsValidRange(coupangReviewRange) Then
        minVal = Split(coupangReviewRange, "~")(0)
        maxVal = Split(coupangReviewRange, "~")(1)
        filterRange.AutoFilter Field:=29, Criteria1:=">=" & minVal, Criteria2:="<=" & maxVal
        filterSummary = filterSummary & "- 쿠팡 평균리뷰수: " & coupangReviewRange & vbLf
    End If

    ' 쿠팡 로켓배송비율 필터링
    coupangRocketRatio = InputBox("쿠팡 로켓배송비율 범위를 입력하세요 (예: 0~50)")
    If coupangRocketRatio <> "" And IsValidRange(coupangRocketRatio) Then
        minVal = CDbl(Split(coupangRocketRatio, "~")(0)) / 100
        maxVal = CDbl(Split(coupangRocketRatio, "~")(1)) / 100
        filterRange.AutoFilter Field:=30, Criteria1:=">=" & minVal, Criteria2:="<=" & maxVal
        filterSummary = filterSummary & "- 쿠팡 로켓배송비율: " & coupangRocketRatio & vbLf
    End If

    ' 쿠팡 판매자로켓 배송비율 필터링
    coupangSellerRocketRatio = InputBox("쿠팡 판매자로켓 배송비율 범위를 입력하세요 (예: 0~50)")
    If coupangSellerRocketRatio <> "" And IsValidRange(coupangSellerRocketRatio) Then
        minVal = CDbl(Split(coupangSellerRocketRatio, "~")(0)) / 100
        maxVal = CDbl(Split(coupangSellerRocketRatio, "~")(1)) / 100
        filterRange.AutoFilter Field:=31, Criteria1:=">=" & minVal, Criteria2:="<=" & maxVal
        filterSummary = filterSummary & "- 쿠팡 판매자로켓 배송비율: " & coupangSellerRocketRatio & vbLf
    End If

    ' 필터링된 데이터를 복사하여 새로운 시트에 저장
    i = 1
    wsName = "공략키워드"
    Do While SheetExists(wsName & i)
        i = i + 1
    Loop
    wsName = wsName & i
    
    Set newWs = ActiveWorkbook.Sheets.Add
    newWs.Name = wsName
    
    On Error Resume Next
    Set copyRange = filterRange.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0

    If copyRange Is Nothing Then
        MsgBox "필터링된 결과가 없습니다. 실행을 종료합니다.", vbExclamation
        ws.AutoFilterMode = False
        Exit Sub
    End If
    
    copyRange.Copy
    newWs.Range("A1").PasteSpecial Paste:=xlPasteValues

    ' 필터링 적용 요약을 마지막 행의 B열에 추가
    lastRow = newWs.Cells(newWs.Rows.count, "A").End(xlUp).Row
    newWs.Cells(lastRow + 1, 2).Value = filterSummary
    newWs.Cells(lastRow + 1, 2).WrapText = True
    
    ' 필터링 요약을 B1 셀의 노트(메모)로 추가
    With newWs.Range("B1")
        .ClearComments ' 기존 노트 제거
        .AddComment.text text:=filterSummary
    End With

    ws.AutoFilterMode = False

    MsgBox "필터링 및 데이터 복사가 완료되었습니다. 결과는 '" & wsName & "' 시트에 저장되었습니다.", vbInformation

End Sub

Function SheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

Function IsValidRange(ByVal userInput As String) As Boolean
    IsValidRange = (InStr(userInput, "~") > 0)
End Function

Function IsValidMonthList(ByVal userInput As String) As Boolean
    IsValidMonthList = (userInput <> "")
End Function

