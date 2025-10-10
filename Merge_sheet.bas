'==============================================================================
' 프로시저명: 엑셀화일시트별병합
' 기능: 선택한 폴더 내의 모든 엑셀 파일들을 하나의 워크북으로 병합
'       각 파일의 모든 시트를 개별 시트로 복사하여 통합
' 작성일: 
' 수정일: 2025-10-10 (임시파일 필터링, 시트명 중복 해결, 오류 처리 강화)
'==============================================================================
Sub 엑셀화일시트별병합()
    '--------------------------------------------------------------------------
    ' 변수 선언부
    '--------------------------------------------------------------------------
    Dim FolderPath As String        ' 선택된 폴더의 경로를 저장
    Dim FileName As String           ' 현재 처리 중인 파일명
    Dim wbSource As Workbook         ' 원본 엑셀 파일 (읽어올 파일)
    Dim wsSource As Worksheet        ' 원본 엑셀 파일의 각 시트
    Dim wbDest As Workbook           ' 대상 엑셀 파일 (병합 결과를 저장할 새 파일)
    Dim wsDest As Worksheet          ' 대상 엑셀 파일의 새 시트
    Dim filePath As String           ' 전체 파일 경로 (폴더경로 + 파일명)
    Dim FileTitle As String          ' 파일명에서 확장자를 제거한 이름
    Dim SheetName As String          ' 새로 생성할 시트의 이름
    Dim fd As FileDialog             ' 폴더 선택 대화상자 객체
    Dim FileCount As Integer         ' 처리된 파일 개수
    Dim SheetCount As Integer        ' 복사된 시트 개수
    Dim DuplicateCounter As Integer  ' 중복 시트명 처리용 카운터
    
    '--------------------------------------------------------------------------
    ' 1단계: 사용자로부터 폴더 선택받기
    '--------------------------------------------------------------------------
    ' 폴더 선택 대화상자 생성
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.Title = "병합할 파일들이 있는 폴더를 선택하세요."
    
    ' 대화상자를 표시하고 사용자의 선택 확인
    ' fd.Show = -1 이면 사용자가 '확인'을 클릭한 것
    If fd.Show = -1 Then
        ' 선택된 폴더 경로를 가져오고 끝에 백슬래시 추가
        FolderPath = fd.SelectedItems(1) & "\"
    Else
        ' 사용자가 '취소'를 클릭한 경우 프로시저 종료
        MsgBox "폴더를 선택하지 않았습니다. 작업을 취소합니다."
        Exit Sub
    End If
    
    '--------------------------------------------------------------------------
    ' 2단계: 병합된 데이터를 저장할 새 워크북 생성
    '--------------------------------------------------------------------------
    Set wbDest = Workbooks.Add      ' 새 워크북 생성 (기본 시트 1개 포함)
    
    ' 처리 카운터 초기화
    FileCount = 0
    SheetCount = 0
    
    '--------------------------------------------------------------------------
    ' 3단계: 선택한 폴더 내의 모든 엑셀 파일 검색 및 병합
    '--------------------------------------------------------------------------
    ' Dir 함수로 첫 번째 엑셀 파일 찾기
    ' *.xls* 패턴으로 .xls, .xlsx, .xlsm 등 모든 엑셀 형식 검색
    FileName = Dir(FolderPath & "*.xls*")
    
    ' 폴더 내의 모든 엑셀 파일에 대해 반복 처리
    ' FileName이 빈 문자열("")이 될 때까지 반복
    Do While FileName <> ""
        '----------------------------------------------------------------------
        ' 3-1: 임시 파일 및 현재 파일 제외
        '----------------------------------------------------------------------
        ' ~$ 로 시작하는 엑셀 임시 파일 건너뛰기
        If Left(FileName, 2) <> "~$" Then
            '----------------------------------------------------------------------
            ' 3-2: 원본 파일 열기 (오류 처리 포함)
            '----------------------------------------------------------------------
            On Error Resume Next    ' 오류 발생 시 다음 코드로 진행
            filePath = FolderPath & FileName    ' 전체 경로 생성
            Set wbSource = Workbooks.Open(filePath, ReadOnly:=True)  ' 읽기 전용으로 파일 열기
            
            ' 파일 열기 실패 시 다음 파일로 진행
            If Err.Number <> 0 Then
                Debug.Print "파일 열기 실패: " & FileName & " (오류: " & Err.Description & ")"
                Err.Clear
                Set wbSource = Nothing
            End If
            On Error GoTo 0     ' 오류 처리 모드 해제
            
            '----------------------------------------------------------------------
            ' 3-3: 파일이 정상적으로 열렸을 경우에만 처리
            '----------------------------------------------------------------------
            If Not wbSource Is Nothing Then
                ' 확장자 제거한 파일명 추출
                FileTitle = Left(FileName, InStrRev(FileName, ".") - 1)
                
                ' 원본 파일의 각 시트에 대해 반복
                For Each wsSource In wbSource.Sheets
                    '--------------------------------------------------------------
                    ' 3-3-1: 대상 워크북에 새 시트 추가
                    '--------------------------------------------------------------
                    Set wsDest = wbDest.Sheets.Add(After:=wbDest.Sheets(wbDest.Sheets.count))
                    
                    '--------------------------------------------------------------
                    ' 3-3-2: 원본 시트의 내용을 새 시트로 복사
                    '--------------------------------------------------------------
                    ' 원본 시트의 사용된 범위(데이터가 있는 모든 셀)를
                    ' 새 시트의 A1 셀부터 복사
                    On Error Resume Next
                    wsSource.UsedRange.Copy wsDest.Cells(1, 1)
                    Application.CutCopyMode = False  ' 복사 모드 해제
                    On Error GoTo 0
                    
                    '--------------------------------------------------------------
                    ' 3-3-3: 시트 이름 설정 (파일명_시트명 형식)
                    '--------------------------------------------------------------
                    ' 시트명 길이 제한: Excel은 최대 31자까지 허용
                    ' 형식: "파일명_시트명"
                    SheetName = FileTitle & "_" & wsSource.Name
                    If Len(SheetName) > 31 Then
                        ' 31자 초과 시 자르기
                        SheetName = Left(SheetName, 31)
                    End If
                    
                    ' 시트 이름 중복 처리
                    DuplicateCounter = 1
                    On Error Resume Next
                    Do While Not IsError(Evaluate("'" & SheetName & "'!A1"))
                        ' 중복 시트명이 존재하면 번호 추가
                        SheetName = Left(FileTitle & "_" & wsSource.Name & "_" & DuplicateCounter, 31)
                        DuplicateCounter = DuplicateCounter + 1
                    Loop
                    On Error GoTo 0
                    
                    ' 시트 이름 설정
                    On Error Resume Next
                    wsDest.Name = SheetName
                    If Err.Number <> 0 Then
                        ' 그래도 실패하면 타임스탬프 추가
                        wsDest.Name = "Sheet_" & Format(Now, "HHmmss")
                        Err.Clear
                    End If
                    On Error GoTo 0
                    
                    SheetCount = SheetCount + 1  ' 복사된 시트 카운트 증가
                Next wsSource
                
                '----------------------------------------------------------------------
                ' 3-4: 원본 파일 닫기
                '----------------------------------------------------------------------
                ' False 매개변수: 변경사항을 저장하지 않고 닫기
                wbSource.Close False
                Set wbSource = Nothing
                FileCount = FileCount + 1    ' 처리된 파일 카운트 증가
            End If
        End If
        
        '----------------------------------------------------------------------
        ' 3-5: 다음 파일로 이동
        '----------------------------------------------------------------------
        ' Dir 함수를 매개변수 없이 호출하면 다음 파일명 반환
        ' 더 이상 파일이 없으면 빈 문자열("") 반환
        FileName = Dir
    Loop
    
    '--------------------------------------------------------------------------
    ' 4단계: 정리 작업
    '--------------------------------------------------------------------------
    ' 처리된 파일이 있는 경우에만 기본 시트 삭제
    If SheetCount > 0 Then
        ' 워크북 생성 시 기본으로 추가되는 빈 시트(Sheet1) 삭제
        Application.DisplayAlerts = False   ' 삭제 확인 메시지 표시 안 함
        On Error Resume Next
        wbDest.Sheets(1).Delete             ' 첫 번째 시트 삭제
        On Error GoTo 0
        Application.DisplayAlerts = True    ' 경고 메시지 표시 다시 활성화
        
        '----------------------------------------------------------------------
        ' 5단계: 작업 완료 메시지
        '----------------------------------------------------------------------
        MsgBox "병합 완료!" & vbCrLf & vbCrLf & _
               "처리된 파일: " & FileCount & "개" & vbCrLf & _
               "복사된 시트: " & SheetCount & "개", vbInformation, "작업 완료"
    Else
        '----------------------------------------------------------------------
        ' 처리된 파일이 없는 경우
        '----------------------------------------------------------------------
        MsgBox "선택한 폴더에 병합할 엑셀 파일이 없습니다.", vbExclamation, "알림"
        ' 빈 워크북 닫기
        wbDest.Close False
    End If
End Sub

