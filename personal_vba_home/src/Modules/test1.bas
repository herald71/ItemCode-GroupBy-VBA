Attribute VB_Name = "Module1"
Sub 엑셀파일시트병합()
    Dim folderPath As String
    Dim FileName As String
    Dim wbSource As Workbook
    Dim wsSource As Worksheet
    Dim wbDest As Workbook
    Dim wsDest As Worksheet
    Dim filePath As String
    Dim lastRow As Long
    Dim FileTitle As String
    Dim fd As FileDialog
    
    ' 사용자에게 폴더 선택 대화상자 표시
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.Title = "병합할 파일들이 있는 폴더를 선택하세요."
    
    ' 사용자가 폴더를 선택한 경우
    If fd.Show = -1 Then
        folderPath = fd.SelectedItems(1) & "\" ' 선택된 폴더 경로를 저장함
    Else
        MsgBox "폴더가 선택되지 않았습니다. 작업을 중단합니다."
        Exit Sub
    End If
    
    ' 병합된 데이터를 저장할 새 워크북 생성
    Set wbDest = Workbooks.Add
    
    ' 폴더 내 첫 번째 엑셀 파일 찾기
    FileName = Dir(folderPath & "*.xls*") ' 엑셀 파일(.xls, .xlsx, .xlsm) 확장자
    
    ' 폴더 내의 모든 파일에 대해 반복
    Do While FileName <> ""
        ' 소스 파일 열기
        filePath = folderPath & FileName
        Set wbSource = Workbooks.Open(filePath)
        
        ' 각 파일의 첫 번째 시트 가져오기
        For Each wsSource In wbSource.Sheets
            ' 새 시트 추가
            Set wsDest = wbDest.Sheets.Add(After:=wbDest.Sheets(wbDest.Sheets.count))
            ' 소스 시트의 모든 데이터를 복사
            wsSource.UsedRange.Copy wsDest.Cells(1, 1)
            
            ' 시트 이름을 파일 이름으로 설정 (확장자 제외)
            FileTitle = Left(FileName, InStrRev(FileName, ".") - 1)
            On Error Resume Next ' 시트 이름이 중복되면 오류 무시
            wsDest.Name = FileTitle
            On Error GoTo 0
        Next wsSource
        
        ' 소스 파일 닫기
        wbSource.Close False
        
        ' 다음 파일로 이동
        FileName = Dir
    Loop
    
    ' 병합된 파일 정리
    Application.DisplayAlerts = False
    wbDest.Sheets(1).Delete ' 기본으로 생성된 빈 시트 삭제
    Application.DisplayAlerts = True
    
    MsgBox "모든 파일이 성공적으로 병합되었습니다!"
End Sub

