Attribute VB_Name = "Module31"
Sub 반복작업()
'*******************************************************************************
' 매크로명: 반복작업
' 작성일자: 2025-04-21
' 기능설명:
' 1. 데이터가 있는 범위에 테두리 설정
' 2. 첫 번째 행 서식 설정
'    - 가운데 정렬
'    - 글꼴 크기 12
'    - 행 높이 30
'    - 배경색 노란색
'    - 굵은 글꼴
' 3. 자동 필터 적용
' 4. C열 기준 내림차순 정렬
' 5. A열에 순차번호 부여
'
' 바로 가기 키: Ctrl+Shift+K
'
' 수정이력:
' 2025-04-21: 최초 작성
' - 기본 서식 설정 및 정렬 기능 구현
' - 필터 자동 해제 및 재적용 기능 추가
' - 오류 처리 추가
' 2025-04-21: 기능 추가
' - 데이터 범위 자동 감지 추가
' - A열 순번 자동 부여 추가
' - 오류 처리 강화
'*******************************************************************************

    On Error GoTo ErrorHandler
    
    '-------------------------------------------
    ' 초기 설정
    '-------------------------------------------
    ' 현재 활성화된 워크시트 참조 설정
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' 변수 선언
    Dim lastRow As Long
    Dim i As Long
    Dim dataRange As Range
    Dim headerRange As Range
    
    ' 마지막 행 찾기 (B열 기준)
    lastRow = ws.Cells(ws.Rows.count, "B").End(xlUp).Row
    
    ' 데이터가 없는 경우 처리
    If lastRow < 2 Then
        MsgBox "처리할 데이터가 없습니다.", vbExclamation
        Exit Sub
    End If
    
    ' 범위 설정
    Set dataRange = ws.Range("A1:D" & lastRow)
    Set headerRange = ws.Range("A1:D1")
    
    ' 화면 업데이트 중지 (처리 속도 향상)
    Application.ScreenUpdating = False
    
    ' 기존에 적용된 필터가 있다면 해제
    If ws.AutoFilterMode Then
        ws.AutoFilterMode = False
    End If
    
    '-------------------------------------------
    ' 테두리 설정
    '-------------------------------------------
    ' 데이터가 있는 범위에 테두리 적용
    With dataRange
        .Borders.LineStyle = xlContinuous
        .Borders.colorIndex = 0
        .Borders.TintAndShade = 0
        .Borders.Weight = xlThin
    End With
    
    '-------------------------------------------
    ' 헤더 행(1행) 서식 설정
    '-------------------------------------------
    With headerRange
        ' 정렬 설정
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        
        ' 글꼴 설정
        With .Font
            .Size = 12
            .Bold = True
        End With
        
        ' 배경색 설정
        With .Interior
            .Color = 65535  ' 노란색
            .TintAndShade = 0
        End With
    End With
    
    '-------------------------------------------
    ' 행 높이 및 필터 설정
    '-------------------------------------------
    ' 1행의 높이를 30으로 설정
    headerRange.RowHeight = 30
    
    '-------------------------------------------
    ' 데이터 정렬
    '-------------------------------------------
    ' C열 기준 내림차순 정렬
    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add2 key:=Range("C2:C" & lastRow), _
        SortOn:=xlSortOnValues, _
        Order:=xlDescending, _
        DataOption:=xlSortNormal
        
    With ws.Sort
        .SetRange dataRange
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ' 자동 필터 적용
    If Not ws.AutoFilterMode Then
        dataRange.AutoFilter
    End If
    
    '-------------------------------------------
    ' 순번 부여
    '-------------------------------------------
    ' A열에 순차번호 부여 (2행부터 마지막 행까지)
    With ws.Range("A2:A" & lastRow)
        ' 순번 부여
        For i = 2 To lastRow
            ws.Cells(i, 1).Value = i - 1
        Next i
        
        ' 가운데 정렬
        .HorizontalAlignment = xlCenter
    End With

ExitSub:
    ' 화면 업데이트 재개
    Application.ScreenUpdating = True
    
    ' 작업 완료 후 F1 셀로 이동
    ws.Range("F1").Select
    Exit Sub

ErrorHandler:
    ' 오류 발생 시 처리
    MsgBox "오류가 발생했습니다." & vbNewLine & _
           "오류 내용: " & Err.Description, vbCritical
    
    ' 화면 업데이트 재개
    Application.ScreenUpdating = True
    Exit Sub
End Sub


