Attribute VB_Name = "Module7"
Sub 쿠팡상품정보시트_자동화()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim cell As Range
    Dim img As Picture
    Dim imageUrl As String
    Dim deliveryType As String
    Dim i As Long
    Dim rocketCount As Long
    Dim sellerRocketCount As Long
    Dim normalDeliveryCount As Long
    Dim outputRow As Long

    ' 활성 시트 설정
    Set ws = ActiveSheet

    ' 마지막 행 찾기 (H열 기준)
    lastRow = ws.Cells(ws.Rows.count, "H").End(xlUp).Row

    ' F열에 데이터가 없는 행 삭제
    For i = lastRow To 2 Step -1
        If ws.Cells(i, "F").Value = "" Then
            ws.Rows(i).Delete
        End If
    Next i

    ' 마지막 행 다시 계산 (행이 삭제되었기 때문에 다시 계산해야 함)
    lastRow = ws.Cells(ws.Rows.count, "H").End(xlUp).Row

    ' 1. H열의 상품URL을 K열에 하이퍼링크로 생성
    For Each cell In ws.Range("H2:H" & lastRow)
        If cell.Value Like "http*://*" Then
            ws.Hyperlinks.Add Anchor:=ws.Cells(cell.Row, "K"), Address:=cell.Value, TextToDisplay:="바로가기"
        End If
    Next cell

    ' 2. I열의 이미지URL을 참고하여 L열에 이미지 삽입
    For Each cell In ws.Range("I2:I" & lastRow)
        imageUrl = cell.Value
        If imageUrl <> "" Then
            ' 이미지 삽입
            On Error Resume Next
            Set img = ws.Pictures.Insert(imageUrl)
            If Not img Is Nothing Then
                With img
                    ' 행 높이를 좀 더 크게 설정하여 이미지가 잘 보이도록
                    ws.Cells(cell.Row, "L").RowHeight = 50
                    .Top = ws.Cells(cell.Row, "L").Top
                    .Left = ws.Cells(cell.Row, "L").Left
                    .Width = ws.Cells(cell.Row, "L").Width
                    .Height = ws.Cells(cell.Row, "L").Height
                    
                    ' 비율에 맞춰 이미지 크기 조정 (너비를 기준으로)
                    Dim ratio As Double
                    ratio = .Width / .Height
                    If .Width > ws.Cells(cell.Row, "L").Width Then
                        .Width = ws.Cells(cell.Row, "L").Width
                        .Height = .Width / ratio
                    End If
                    
                    ' 이미지가 셀 중앙에 위치하도록 조정
                    .Top = ws.Cells(cell.Row, "L").Top + (ws.Cells(cell.Row, "L").Height - .Height) / 2
                    .Left = ws.Cells(cell.Row, "L").Left + (ws.Cells(cell.Row, "L").Width - .Width) / 2
                End With
            End If
            On Error GoTo 0
        End If
    Next cell

    ' 3. J열의 배송형태에 따라 셀 색상 변경 및 갯수 세기
    rocketCount = 0
    sellerRocketCount = 0
    normalDeliveryCount = 0

    For Each cell In ws.Range("J2:J" & lastRow)
        deliveryType = cell.Value
        Select Case deliveryType
            Case "판매자로켓"
                cell.Interior.Color = RGB(255, 200, 150) ' 연한 오렌지색
                sellerRocketCount = sellerRocketCount + 1
            Case "로켓배송"
                cell.Interior.Color = RGB(150, 200, 255) ' 연한 파란색
                rocketCount = rocketCount + 1
            Case "일반배송"
                cell.Interior.Color = RGB(255, 150, 150) ' 연한 빨강색
                normalDeliveryCount = normalDeliveryCount + 1
        End Select
    Next cell

    ' 4. A열, H열, I열 숨기기
    ws.Columns("A:A").Hidden = True
    ws.Columns("H:I").Hidden = True

    ' 5. 테두리 및 정렬 설정
    With ws.Range("A1:L" & lastRow)
        .Borders.LineStyle = xlContinuous ' 테두리 설정
        .HorizontalAlignment = xlCenter ' 텍스트 가운데 정렬
        .VerticalAlignment = xlCenter ' 내용 가운데 정렬
    End With

    ' C열의 내용 왼쪽 정렬, C1은 가운데 정렬
    ws.Range("C2:C" & lastRow).HorizontalAlignment = xlLeft
    ws.Range("C1").HorizontalAlignment = xlCenter

    ' 6. K1 셀과 L1 셀에 "바로가기"와 "이미지" 텍스트 추가
    ws.Range("K1").Value = "바로가기"
    ws.Range("L1").Value = "이미지"

    ' 7. 배송 형태 갯수를 C열 마지막 행 아래 두번째에 출력
    outputRow = lastRow + 2
    ws.Cells(outputRow, "C").Value = "배송형태 통계: "
    ws.Cells(outputRow + 1, "C").Value = "로켓배송: " & rocketCount & "건"
    ws.Cells(outputRow + 2, "C").Value = "판매자로켓: " & sellerRocketCount & "건"
    ws.Cells(outputRow + 3, "C").Value = "일반배송: " & normalDeliveryCount & "건"

    ' 두 번째 기능 시작 - 가격대 분석

    Dim priceRange As Range
    Dim qtyRange As Range

    ' 마지막 행 찾기 (F열 기준)
    lastRow = ws.Cells(ws.Rows.count, "F").End(xlUp).Row

    ' 판매가와 판매갯수 범위 설정
    Set priceRange = ws.Range("F2:F" & lastRow) ' F열에 판매가
    Set qtyRange = ws.Range("G2:G" & lastRow)   ' G열에 판매갯수

    ' 가격대 간격 설정 - 사용자 입력
    Dim interval As Double
    interval = Application.InputBox("가격대 간격을 입력하세요:", "가격대 간격 입력", Type:=1)
    
    ' 입력값이 0 이하일 경우 종료
    If interval <= 0 Then
        MsgBox "올바른 가격대 간격을 입력하세요."
        Exit Sub
    End If

    ' 가격대별 판매갯수를 저장할 Dictionary 생성
    Dim priceGroups As Object
    Set priceGroups = CreateObject("Scripting.Dictionary")

    For i = 2 To lastRow
        Dim price As Double
        Dim qty As Double
        price = ws.Cells(i, "F").Value
        qty = ws.Cells(i, "G").Value

        ' 유효한 숫자인지 확인
        If IsNumeric(price) And IsNumeric(qty) Then
            ' 가격대 계산
            Dim groupStart As Double
            groupStart = Int(price / interval) * interval
            groupKey = Format(groupStart, "#,##0") & " - " & Format(groupStart + interval - 1, "#,##0") ' 쉼표 추가

            ' Dictionary에 누적
            If priceGroups.exists(groupKey) Then
                priceGroups(groupKey) = priceGroups(groupKey) + qty
            Else
                priceGroups.Add groupKey, qty
            End If
        End If
    Next i

    ' 가장 많이 팔린 가격대 찾기
    Dim maxQty As Double
    Dim maxGroup As String
    Dim secondMaxQty As Double
    Dim secondMaxGroup As String
    Dim thirdMaxQty As Double
    Dim thirdMaxGroup As String
    maxQty = 0
    secondMaxQty = 0
    thirdMaxQty = 0

    For Each key In priceGroups.Keys
        If priceGroups(key) > maxQty Then
            ' 이전 최대값들을 한 단계씩 밀기
            thirdMaxQty = secondMaxQty
            thirdMaxGroup = secondMaxGroup
            secondMaxQty = maxQty
            secondMaxGroup = maxGroup
            maxQty = priceGroups(key)
            maxGroup = key
        ElseIf priceGroups(key) > secondMaxQty Then
            ' 이전 두 번째 최대값을 세 번째로 이동
            thirdMaxQty = secondMaxQty
            thirdMaxGroup = secondMaxGroup
            secondMaxQty = priceGroups(key)
            secondMaxGroup = key
        ElseIf priceGroups(key) > thirdMaxQty Then
            ' 세 번째 최대값 갱신
            thirdMaxQty = priceGroups(key)
            thirdMaxGroup = key
        End If
    Next key

    ' 결과 메시지 생성
    Dim message As String
    message = ""
    If maxGroup <> "" Then
        message = "가장 많이 팔린 가격대는 " & maxGroup & "이며, 총 판매갯수는 " & Format(maxQty, "#,##0") & "입니다." & vbCrLf
        If secondMaxGroup <> "" Then
            message = message & "두 번째로 많이 팔린 가격대는 " & secondMaxGroup & "이며, 총 판매갯수는 " & Format(secondMaxQty, "#,##0") & "입니다." & vbCrLf
        End If
        If thirdMaxGroup <> "" Then
            message = message & "세 번째로 많이 팔린 가격대는 " & thirdMaxGroup & "이며, 총 판매갯수는 " & Format(thirdMaxQty, "#,##0") & "입니다."
        End If
        MsgBox message
    Else
        MsgBox "데이터가 없거나 유효한 값이 없습니다."
        Exit Sub
    End If

    ' 결과값을 C열의 일반배송 통계값 아래에 출력
    outputRow = outputRow + 5 ' 배송 통계 이후 5행 아래에 결과 출력
    ws.Cells(outputRow, "C").Value = "가격대 분석 결과:"
    ws.Cells(outputRow + 1, "C").Value = "가장 많이 팔린 가격대는 " & maxGroup & "이며, 총 판매갯수는 " & Format(maxQty, "#,##0") & "입니다."

    If secondMaxGroup <> "" Then
        ws.Cells(outputRow + 2, "C").Value = "두 번째로 많이 팔린 가격대는 " & secondMaxGroup & "이며, 총 판매갯수는 " & Format(secondMaxQty, "#,##0") & "입니다."
    End If

    If thirdMaxGroup <> "" Then
        ws.Cells(outputRow + 3, "C").Value = "세 번째로 많이 팔린 가격대는 " & thirdMaxGroup & "이며, 총 판매갯수는 " & Format(thirdMaxQty, "#,##0") & "입니다."
    End If

End Sub


