Attribute VB_Name = "Module17"
Sub MergeExcelFiles_Onesheet()
    Dim MyPath As String, FilesInPath As String
    Dim MyFiles() As String
    Dim SourceRcount As Long, FNum As Long
    Dim mybook As Workbook, BaseWks As Worksheet
    Dim sourceRange As Range, destrange As Range
    Dim rnum As Long, CalcMode As Long
    Dim fd As FileDialog ' 파일 대화상자 변수 선언

    ' 폴더 선택 대화상자를 사용하여 사용자에게 폴더 선택을 요청
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.Title = "폴더를 선택하세요"

    ' 사용자가 폴더를 선택하면 경로를 MyPath에 저장
    If fd.Show = -1 Then
        MyPath = fd.SelectedItems(1) ' 선택된 폴더 경로
    Else
        MsgBox "폴더가 선택되지 않았습니다. 매크로를 종료합니다."
        Exit Sub
    End If

    ' 폴더 경로 끝이 "\"로 끝나지 않으면 "\" 추가
    If Right(MyPath, 1) <> "\" Then MyPath = MyPath & "\"

    ' 파일 경로와 일치하는 파일 목록을 가져옵니다.
    FilesInPath = Dir(MyPath & "*.xls*")
    If FilesInPath = "" Then
        MsgBox "해당 폴더에는 합칠 Excel 파일이 없습니다."
        Exit Sub
    End If

    ' 파일 배열에 파일 이름을 저장
    FNum = 0
    Do While FilesInPath <> ""
        FNum = FNum + 1
        ReDim Preserve MyFiles(1 To FNum)
        MyFiles(FNum) = FilesInPath
        FilesInPath = Dir()
    Loop

    ' 활성 워크북의 첫 번째 시트를 사용
    Set BaseWks = ActiveWorkbook.Sheets(1)
    rnum = 1

    ' 각 파일을 열고 복사한 다음 합칩니다.
    For FNum = 1 To UBound(MyFiles)
        Set mybook = Workbooks.Open(MyPath & MyFiles(FNum))

        ' 첫 번째 파일에서는 모든 데이터를 복사하고, 그 다음 파일에서는 첫 번째 행을 제외합니다.
        With mybook.Sheets(1)
            Set sourceRange = .UsedRange
            If FNum > 1 Then
                Set sourceRange = sourceRange.Offset(1, 0).Resize(sourceRange.Rows.count - 1, sourceRange.Columns.count)
            End If

            ' 데이터를 기본 워크시트로 복사
            If rnum + sourceRange.Rows.count > BaseWks.Rows.count Then
                MsgBox "결과가 Excel 시트의 최대 행 수를 초과합니다."
                mybook.Close SaveChanges:=False
                GoTo ExitTheSub
            Else
                Set destrange = BaseWks.Cells(rnum, "A")
                sourceRange.Copy Destination:=destrange
                rnum = rnum + sourceRange.Rows.count
            End If
        End With

        mybook.Close SaveChanges:=False
    Next FNum

ExitTheSub:
End Sub

