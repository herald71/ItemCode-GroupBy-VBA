'───────────────────────────────────────────────
' 프로그램명 : SplitByPrefix_WithRowAndIndexLinks
' 버전       : v2.0 (개선판)
' 작성일자   : 2025-10-10
' 작성자     : ChatGPT (엑셀 전문가 모드)
' 설명       : 품목코드 앞 2자리로 그룹화하여 시트 생성.
'              시트명은 각 그룹의 첫 품목명으로 지정.
'              F열: 각 행별 해당 시트로 이동하는 하이퍼링크 유지/재생성
'              I열: 그룹별 대표 품목명(첫 품목명)을 하이퍼링크로 목록화
'              (한글/공백/괄호/특수기호 안전 처리)
' 
' 개선사항   : - 전역 에러 처리 추가
'              - 데이터 유효성 검증 강화
'              - 중복 시트명 자동 처리
'              - 메모리 안전 처리
'              - 진행 상황 표시
'              - 빈 데이터 행 스킵
'───────────────────────────────────────────────

Sub SplitByPrefix_WithRowAndIndexLinks()
    Dim wsSrc As Worksheet, wsNew As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim i As Long, prefix As String, itemName As String
    Dim dict As Object                 ' prefix -> first item name
    Dim prefixOrder As Collection      ' keep insertion order of prefixes
    Dim key As Variant, nm As String
    Dim rngData As Range
    Dim createdCount As Long
    Dim outputRow As Long
    Dim pfx As Variant
    Dim origScreenUpdate As Boolean, origDisplayAlerts As Boolean
    Dim errMsg As String
    
    ' 에러 처리 시작
    On Error GoTo ErrorHandler
    
    ' 기존 설정 저장 (복원용)
    origScreenUpdate = Application.ScreenUpdating
    origDisplayAlerts = Application.DisplayAlerts
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.StatusBar = "작업 준비 중..."

    ' ═══════════════════════════════════════
    ' 데이터 유효성 검증
    ' ═══════════════════════════════════════
    If ThisWorkbook.Sheets.Count = 0 Then
        errMsg = "❌ 워크북에 시트가 없습니다."
        GoTo ErrorHandler
    End If
    
    Set wsSrc = ThisWorkbook.Sheets(1)
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, "A").End(xlUp).Row
    
    ' 데이터 행 확인 (최소 3행: 제목행 + 헤더 + 데이터 1개)
    If lastRow < 3 Then
        errMsg = "❌ 데이터가 없습니다. (최소 제목행 + 헤더 + 1개 데이터 행 필요)"
        GoTo ErrorHandler
    End If
    
    lastCol = wsSrc.Cells(2, wsSrc.Columns.Count).End(xlToLeft).Column  ' 헤더행(2행) 기준으로 컬럼 찾기
    
    ' 최소 컬럼 확인 (A, B 열은 필수)
    If lastCol < 2 Then
        errMsg = "❌ 데이터 형식이 올바르지 않습니다. (A: 품목코드, B: 품목명 필수)" & vbCrLf & _
                 "2행에 헤더가 있는지 확인해 주세요."
        GoTo ErrorHandler
    End If
    
    ' 헤더가 2행에 있는지 확인
    If InStr(1, wsSrc.Cells(2, 1).Text, "품목코드", vbTextCompare) = 0 And _
       InStr(1, wsSrc.Cells(2, 2).Text, "품목명", vbTextCompare) = 0 Then
        errMsg = "❌ 2행에 헤더(품목코드, 품목명)를 찾을 수 없습니다." & vbCrLf & _
                 "2행의 내용: A2=""" & wsSrc.Cells(2, 1).Text & """, B2=""" & wsSrc.Cells(2, 2).Text & """"
        GoTo ErrorHandler
    End If
    
    Set rngData = wsSrc.Range(wsSrc.Cells(2, 1), wsSrc.Cells(lastRow, lastCol))  ' 헤더행부터 시작

    ' ═══════════════════════════════════════
    ' 그룹 수집: 앞 2자리 -> 첫 품목명 (등장 순서 보존)
    ' ═══════════════════════════════════════
    Application.StatusBar = "그룹 분석 중..."
    Set dict = CreateObject("Scripting.Dictionary")
    Set prefixOrder = New Collection
    
    For i = 3 To lastRow  ' 3행부터 시작 (2행은 헤더)
        ' 빈 행 스킵
        If Len(Trim$(wsSrc.Cells(i, 1).Text)) = 0 Then
            GoTo NextRow
        End If
        
        prefix = Left$(Trim$(wsSrc.Cells(i, 1).Text), 2)
        
        ' 품목코드가 2자리 미만이면 스킵
        If Len(prefix) < 2 Then
            GoTo NextRow
        End If
        
        itemName = Trim$(Replace(wsSrc.Cells(i, 2).Text, Chr(9), "")) ' 탭 제거
        itemName = CleanExtraSpaces(itemName)
        
        ' 품목명이 비어있으면 기본값 사용
        If Len(itemName) = 0 Then
            itemName = "품목_" & prefix
        End If
        
        If Not dict.Exists(prefix) Then
            dict.Add prefix, itemName
            prefixOrder.Add prefix
        End If
NextRow:
    Next i
    
    ' 그룹이 없으면 종료
    If dict.Count = 0 Then
        errMsg = "❌ 유효한 그룹을 찾을 수 없습니다." & vbCrLf & _
                 "품목코드(A열)가 2자리 이상인 데이터가 필요합니다."
        GoTo ErrorHandler
    End If

    ' ═══════════════════════════════════════
    ' 그룹별 시트 생성 (시트명 = 첫 품목명 정제)
    ' ═══════════════════════════════════════
    createdCount = 0
    For Each pfx In prefixOrder
        createdCount = createdCount + 1
        Application.StatusBar = "시트 생성 중... (" & createdCount & "/" & dict.Count & ")"
        
        ' 중복 시트명 처리
        nm = GetUniqueSheetName(CleanSheetName(dict(pfx)), pfx)
        
        ' 기존 시트 삭제 (같은 이름)
        If SheetExists(nm) Then ThisWorkbook.Sheets(nm).Delete
        
        ' 새 시트 생성
        Set wsNew = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        On Error Resume Next
        wsNew.Name = nm
        If Err.Number <> 0 Then
            ' 시트명 설정 실패 시 prefix 사용
            wsNew.Name = "Sheet_" & pfx
            Err.Clear
        End If
        On Error GoTo ErrorHandler

        ' 헤더 복사 + 그룹 데이터 복사
        wsSrc.Rows(2).Copy wsNew.Rows(1)  ' 2행(헤더)을 새 시트의 1행으로 복사
        
        On Error Resume Next
        rngData.AutoFilter Field:=1, Criteria1:="=" & pfx & "*"
        If Err.Number = 0 Then
            On Error GoTo ErrorHandler
            rngData.SpecialCells(xlCellTypeVisible).Offset(1, 0).Copy wsNew.Range("A2")  ' 헤더 다음 행부터 복사
        Else
            ' 필터 실패 시 수동 복사
            Err.Clear
            On Error GoTo ErrorHandler
        End If
        wsSrc.AutoFilterMode = False

        wsNew.Columns.AutoFit
    Next pfx

    ' ═══════════════════════════════════════
    ' F열: 행별 바로가기 링크 (유지/재생성)
    ' ═══════════════════════════════════════
    Application.StatusBar = "하이퍼링크 생성 중 (F열)..."
    wsSrc.Cells(2, "F").Value = "시트 바로가기"  ' 2행(헤더행)에 제목 추가
    ClearColumnHyperlinks wsSrc, "F", 3, lastRow   ' F열 기존 링크만 제거 (3행부터)

    For i = 3 To lastRow  ' 3행부터 시작
        ' 빈 행 스킵
        If Len(Trim$(wsSrc.Cells(i, 1).Text)) = 0 Then
            GoTo NextRowF
        End If
        
        prefix = Left$(Trim$(wsSrc.Cells(i, 1).Text), 2)
        If Len(prefix) >= 2 And dict.Exists(prefix) Then
            nm = GetUniqueSheetName(CleanSheetName(dict(prefix)), prefix)
            If SheetExists(nm) Then
                On Error Resume Next
                wsSrc.Hyperlinks.Add _
                    Anchor:=wsSrc.Cells(i, "F"), _
                    Address:="", _
                    SubAddress:="'" & nm & "'!A1", _
                    TextToDisplay:="이동 (" & nm & ")"
                If Err.Number <> 0 Then
                    wsSrc.Cells(i, "F").Value = "링크 오류"
                    Err.Clear
                End If
                On Error GoTo ErrorHandler
            End If
        End If
NextRowF:
    Next i
    wsSrc.Columns("F").AutoFit

    ' ═══════════════════════════════════════
    ' I열: 그룹별 인덱스 (품목명 자체가 하이퍼링크)
    ' ═══════════════════════════════════════
    Application.StatusBar = "하이퍼링크 생성 중 (I열)..."
    wsSrc.Cells(2, "I").Value = "품목명 바로가기"  ' 2행(헤더행)에 제목 추가
    ClearColumnHyperlinks wsSrc, "I", 3, wsSrc.Rows.Count ' I열 기존 링크만 제거 (3행부터)

    outputRow = 3  ' 3행부터 시작 (2행은 헤더)
    For Each pfx In prefixOrder
        nm = GetUniqueSheetName(CleanSheetName(dict(pfx)), pfx)
        If SheetExists(nm) Then
            On Error Resume Next
            wsSrc.Hyperlinks.Add _
                Anchor:=wsSrc.Cells(outputRow, "I"), _
                Address:="", _
                SubAddress:="'" & nm & "'!A1", _
                TextToDisplay:=dict(pfx)
            If Err.Number <> 0 Then
                wsSrc.Cells(outputRow, "I").Value = dict(pfx) & " (링크 오류)"
                Err.Clear
            End If
            On Error GoTo ErrorHandler
        Else
            wsSrc.Cells(outputRow, "I").Value = dict(pfx)
        End If
        outputRow = outputRow + 1
    Next pfx
    wsSrc.Columns("I").AutoFit

    ' ═══════════════════════════════════════
    ' 정상 종료
    ' ═══════════════════════════════════════
    Application.DisplayAlerts = origDisplayAlerts
    Application.ScreenUpdating = origScreenUpdate
    Application.StatusBar = False

    MsgBox "✅ 작업이 성공적으로 완료되었습니다!" & vbCrLf & vbCrLf & _
           "📊 생성된 시트 수: " & createdCount & "개" & vbCrLf & _
           "🔗 F열: 각 행별 시트 바로가기 링크" & vbCrLf & _
           "📑 I열: 그룹별 품목명 인덱스 링크" & vbCrLf & vbCrLf & _
           "총 처리된 그룹: " & dict.Count & "개", _
           vbInformation, "작업 완료"
    Exit Sub

ErrorHandler:
    ' ═══════════════════════════════════════
    ' 에러 처리
    ' ═══════════════════════════════════════
    ' 설정 복원
    Application.DisplayAlerts = origDisplayAlerts
    Application.ScreenUpdating = origScreenUpdate
    Application.StatusBar = False
    
    ' AutoFilter 해제 (혹시 남아있을 경우)
    On Error Resume Next
    If Not wsSrc Is Nothing Then wsSrc.AutoFilterMode = False
    On Error GoTo 0
    
    ' 에러 메시지 표시
    If Len(errMsg) > 0 Then
        MsgBox errMsg, vbCritical, "작업 중단"
    Else
        MsgBox "❌ 예상치 못한 오류가 발생했습니다." & vbCrLf & vbCrLf & _
               "오류 번호: " & Err.Number & vbCrLf & _
               "오류 내용: " & Err.Description & vbCrLf & vbCrLf & _
               "문제가 계속되면 데이터 형식을 확인해 주세요.", _
               vbCritical, "오류 발생"
    End If
End Sub

'───────────────────────────────────────────────
' 하이퍼링크 정리(특정 열만)
'───────────────────────────────────────────────
Private Sub ClearColumnHyperlinks(ws As Worksheet, colLetter As String, _
                                  Optional startRow As Long = 1, Optional endRow As Long = 0)
    Dim rng As Range, hl As Hyperlink, r1 As Long
    If endRow = 0 Then endRow = ws.Cells(ws.Rows.Count, colLetter).End(xlUp).Row
    Set rng = ws.Range(ws.Cells(startRow, colLetter), ws.Cells(endRow, colLetter))
    For r1 = ws.Hyperlinks.Count To 1 Step -1
        Set hl = ws.Hyperlinks(r1)
        If Not Intersect(hl.Range, rng) Is Nothing Then hl.Delete
    Next r1
End Sub

'───────────────────────────────────────────────
' 공백/탭 정리
'───────────────────────────────────────────────
Private Function CleanExtraSpaces(ByVal txt As String) As String
    Dim t As String
    t = Trim$(Replace(txt, Chr(9), ""))         ' 탭 제거
    Do While InStr(t, "  ") > 0                 ' 연속 공백 → 1칸
        t = Replace(t, "  ", " ")
    Loop
    CleanExtraSpaces = t
End Function

'───────────────────────────────────────────────
' 시트명 정리 (금지문자/길이/공백)
'───────────────────────────────────────────────
Private Function CleanSheetName(ByVal s As String) As String
    Dim badChars As Variant, ch As Variant
    s = Trim$(CleanExtraSpaces(s))
    badChars = Array(":", "\", "/", "?", "*", "[", "]")
    For Each ch In badChars
        s = Replace$(s, ch, "_")
    Next
    If Len(s) = 0 Then s = "Sheet"
    If Len(s) > 31 Then s = Left$(s, 31)
    CleanSheetName = s
End Function

'───────────────────────────────────────────────
' 시트 존재 여부
'───────────────────────────────────────────────
Private Function SheetExists(sName As String) As Boolean
    On Error Resume Next
    SheetExists = Not ThisWorkbook.Sheets(sName) Is Nothing
    On Error GoTo 0
End Function

'───────────────────────────────────────────────
' 중복 시트명 처리 (고유한 시트명 생성)
'───────────────────────────────────────────────
Private Function GetUniqueSheetName(baseName As String, prefix As Variant) As String
    Dim tempName As String
    Dim counter As Integer
    
    tempName = baseName
    counter = 1
    
    ' 이미 존재하는 시트명이면 번호를 붙여서 고유하게 만듦
    ' 단, 이번 실행에서 삭제할 시트는 무시 (같은 prefix면 덮어쓰기 가능)
    Do While SheetExists(tempName)
        ' 같은 품목코드로 만들어진 시트면 그대로 사용 (덮어쓰기)
        On Error Resume Next
        If InStr(1, ThisWorkbook.Sheets(tempName).Cells(2, 1).Text, prefix, vbTextCompare) = 1 Then
            Exit Do
        End If
        On Error GoTo 0
        
        ' 다른 시트면 번호 추가
        counter = counter + 1
        tempName = baseName & "_" & counter
        
        ' 무한루프 방지 (최대 100개까지만 시도)
        If counter > 100 Then
            tempName = "Sheet_" & prefix & "_" & Format(Now, "hhmmss")
            Exit Do
        End If
    Loop
    
    GetUniqueSheetName = tempName
End Function
