Attribute VB_Name = "Module22"
' ==============================================
' 프로그램명 : URL생성하고하이퍼링크만들기
' 설명 :
'   - B열(2번째 열)에 있는 검색어를 기반으로 네이버, 네이버쇼핑, 쿠팡, 유튜브, 틱톡, 쇼핑라이브, 검색어트렌드, 도매꾹의 검색 URL을 자동 생성합니다.
'   - 각 URL로 이동할 수 있는 하이퍼링크를 새로운 열에 생성합니다.
'   - URL이 들어간 열은 숨기고, 전체 데이터에 테두리를 추가합니다.
'   - 기존 데이터에 영향을 주지 않도록 첫 번째 빈 열부터 결과를 작성합니다.
'   - URL 인코딩을 통해 한글/특수문자 검색어도 정상 처리합니다.
'   - 작업 진행 상황을 상태표시줄에 표시하고, 완료 시 메시지 박스를 띄웁니다.
' 사용법 :
'   1. B열(2번째 열)에 검색어를 입력합니다(2행부터).
'   2. 시트를 활성화한 상태에서 이 매크로를 실행합니다.
'   3. 자동으로 URL 및 하이퍼링크가 생성됩니다.
' ==============================================
'
' ====== [EUC-KR 인코딩 함수 추가] ======
#If VBA7 Then
    Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" ( _
        ByVal CodePage As Long, ByVal dwFlags As Long, _
        ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long, _
        ByVal lpMultiByteStr As LongPtr, ByVal cbMultiByte As Long, _
        ByVal lpDefaultChar As LongPtr, ByVal lpUsedDefaultChar As LongPtr) As Long
#Else
    Private Declare Function WideCharToMultiByte Lib "kernel32" ( _
        ByVal CodePage As Long, ByVal dwFlags As Long, _
        ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, _
        ByVal lpMultiByteStr As Long, ByVal cbMultiByte As Long, _
        ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
#End If

Function EncodeEUC_KR(str As String) As String
    Dim l As Long, i As Long
    Dim arr() As Byte
    Dim result As String
    Dim b As Byte
    Dim tempStr As String

    l = WideCharToMultiByte(949, 0, StrPtr(str), -1, 0, 0, 0, 0)
    If l > 1 Then
        ReDim arr(l - 2)
        WideCharToMultiByte 949, 0, StrPtr(str), -1, VarPtr(arr(0)), l - 1, 0, 0
        For i = 0 To UBound(arr)
            b = arr(i)
            If (b >= &H30 And b <= &H39) Or (b >= &H41 And b <= &H5A) Or (b >= &H61 And b <= &H7A) Then
                tempStr = Chr(b)
            Else
                tempStr = "%" & Right("0" & Hex(b), 2)
            End If
            result = result & tempStr
        Next i
    End If
    EncodeEUC_KR = result
End Function

Sub URL생성하고하이퍼링크만들기()
    Dim ws As Worksheet ' 작업할 시트 객체
    Dim lastRow As Long ' B열의 마지막 데이터 행 번호
    Dim startCol As Long ' 첫 번째 빈 열 번호
    Dim startColURL As Long ' URL을 쓸 시작 열 번호
    Dim startColLink As Long ' 하이퍼링크를 쓸 시작 열 번호
    Dim i As Long ' 반복문 인덱스
    Dim dataRange As Range ' 테두리 적용 범위
    
    ' 현재 활성화된 시트 선택
    On Error Resume Next
    Set ws = ActiveSheet
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "활성화된 시트가 없습니다. 먼저 작업할 시트를 활성화하세요!", vbExclamation
        Exit Sub
    End If
    
    ' B열 데이터가 있는 마지막 행 찾기 (2행부터 데이터가 있다고 가정)
    lastRow = ws.Cells(ws.Rows.count, "B").End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "B열에 데이터가 없습니다. 확인해주세요!", vbExclamation
        Exit Sub
    End If
    
    ' 상태 표시줄 초기화
    Application.StatusBar = "작업 시작: URL 생성 중..."
    
    ' 첫 번째 빈 열 찾기 (기존 데이터에 영향 없도록)
    startCol = 1
    Do While Application.WorksheetFunction.CountA(ws.Columns(startCol)) > 0
        startCol = startCol + 1
    Loop
    
    ' URL 및 하이퍼링크 추가할 열 위치 설정
    startColURL = startCol ' URL 열 시작
    startColLink = startColURL + 8 ' 하이퍼링크 열 시작 (URL 열 8개 뒤)
    
    ' URL 열 제목 추가
    ws.Cells(1, startColURL).Value = "네이버검색 URL"
    ws.Cells(1, startColURL + 1).Value = "네이버쇼핑검색 URL"
    ws.Cells(1, startColURL + 2).Value = "쿠팡URL"
    ws.Cells(1, startColURL + 3).Value = "유튜브URL"
    ws.Cells(1, startColURL + 4).Value = "틱톡URL"
    ws.Cells(1, startColURL + 5).Value = "쇼핑라이브URL"
    ws.Cells(1, startColURL + 6).Value = "검색어트렌드 URL"
    ws.Cells(1, startColURL + 7).Value = "도매꾹검색 URL"
    
    ' 각 행마다 검색어로 URL 생성 및 추가
    For i = 2 To lastRow
        Application.StatusBar = "URL 생성 중... (" & i - 1 & "/" & lastRow - 1 & ")"
        
        If ws.Cells(i, 2).Value <> "" Then ' B열에 검색어가 있을 때만
            ' 네이버 검색 URL
            ws.Cells(i, startColURL).Value = "https://search.naver.com/search.naver?where=nexearch&sm=top_hty&fbm=0&ie=utf8&query=" & WorksheetFunction.EncodeURL(ws.Cells(i, 2).Value)
            ' 네이버쇼핑 검색 URL
            ws.Cells(i, startColURL + 1).Value = "https://search.shopping.naver.com/ns/search?query=" & WorksheetFunction.EncodeURL(ws.Cells(i, 2).Value)
            ' 쿠팡 검색 URL
            ws.Cells(i, startColURL + 2).Value = "https://www.coupang.com/np/search?q=" & WorksheetFunction.EncodeURL(ws.Cells(i, 2).Value)
            ' 유튜브 검색 URL
            ws.Cells(i, startColURL + 3).Value = "https://www.youtube.com/results?search_query=" & WorksheetFunction.EncodeURL(ws.Cells(i, 2).Value)
            ' 틱톡 검색 URL
            ws.Cells(i, startColURL + 4).Value = "https://www.tiktok.com/search?q=" & WorksheetFunction.EncodeURL(ws.Cells(i, 2).Value)
            ' 쇼핑라이브 검색 URL
            ws.Cells(i, startColURL + 5).Value = "https://shoppinglive.naver.com/search/lives?query=" & WorksheetFunction.EncodeURL(ws.Cells(i, 2).Value)
            ' 검색어트렌드 URL (고정)
            ws.Cells(i, startColURL + 6).Value = "https://datalab.naver.com/keyword/trendSearch.naver"
            ' 도매꾹 검색 URL (EUC-KR 인코딩 적용)
            ws.Cells(i, startColURL + 7).Value = "https://domeggook.com/main/item/itemList.php?sfc=ttl&sf=ttl&sw=" & EncodeEUC_KR(ws.Cells(i, 2).Value)
        End If
    Next i
    
    ' 하이퍼링크 열 제목 추가
    ws.Cells(1, startColLink).Value = "네이버검색"
    ws.Cells(1, startColLink + 1).Value = "네이버쇼핑검색"
    ws.Cells(1, startColLink + 2).Value = "쿠팡"
    ws.Cells(1, startColLink + 3).Value = "유튜브"
    ws.Cells(1, startColLink + 4).Value = "틱톡"
    ws.Cells(1, startColLink + 5).Value = "쇼핑라이브"
    ws.Cells(1, startColLink + 6).Value = "검색어트렌드"
    ws.Cells(1, startColLink + 7).Value = "도매꾹검색"
    
    ' 각 행마다 하이퍼링크 추가
    For i = 2 To lastRow
        Application.StatusBar = "하이퍼링크 추가 중... (" & i - 1 & "/" & lastRow - 1 & ")"
        
        ' 각 URL이 있을 때만 하이퍼링크 생성
        If ws.Cells(i, startColURL).Value <> "" Then
            ws.Hyperlinks.Add ws.Cells(i, startColLink), ws.Cells(i, startColURL).Value, , , "바로가기"
        End If
        If ws.Cells(i, startColURL + 1).Value <> "" Then
            ws.Hyperlinks.Add ws.Cells(i, startColLink + 1), ws.Cells(i, startColURL + 1).Value, , , "바로가기"
        End If
        If ws.Cells(i, startColURL + 2).Value <> "" Then
            ws.Hyperlinks.Add ws.Cells(i, startColLink + 2), ws.Cells(i, startColURL + 2).Value, , , "바로가기"
        End If
        If ws.Cells(i, startColURL + 3).Value <> "" Then
            ws.Hyperlinks.Add ws.Cells(i, startColLink + 3), ws.Cells(i, startColURL + 3).Value, , , "바로가기"
        End If
        If ws.Cells(i, startColURL + 4).Value <> "" Then
            ws.Hyperlinks.Add ws.Cells(i, startColLink + 4), ws.Cells(i, startColURL + 4).Value, , , "바로가기"
        End If
        If ws.Cells(i, startColURL + 5).Value <> "" Then
            ws.Hyperlinks.Add ws.Cells(i, startColLink + 5), ws.Cells(i, startColURL + 5).Value, , , "바로가기"
        End If
        ' 검색어트렌드는 고정 URL로 하이퍼링크 생성
        ws.Hyperlinks.Add ws.Cells(i, startColLink + 6), "https://datalab.naver.com/keyword/trendSearch.naver", , , "바로가기"
        ' 도매꾹 하이퍼링크 생성
        If ws.Cells(i, startColURL + 7).Value <> "" Then
            ws.Hyperlinks.Add ws.Cells(i, startColLink + 7), ws.Cells(i, startColURL + 7).Value, , , "바로가기"
        End If
    Next i
    
    ' URL 열 숨기기 (사용자에게 URL이 보이지 않도록)
    ws.Range(ws.Columns(startColURL), ws.Columns(startColURL + 7)).EntireColumn.Hidden = True
    
    ' 전체 데이터에 테두리 추가
    Set dataRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, startColLink + 7))
    With dataRange.Borders
        .LineStyle = xlContinuous
        .Color = RGB(0, 0, 0)
        .Weight = xlThin
    End With
    
    ' 완료 메시지 및 상태표시줄 초기화
    Application.StatusBar = "작업 완료!"
    Application.StatusBar = False
    MsgBox "URL과 하이퍼링크가 생성되고, URL 열이 숨겨졌으며 테두리가 추가되었습니다!", vbInformation
End Sub




