Attribute VB_Name = "Module18"
Sub 열기준으로데이타분리하기()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim colToSplit As Long
    Dim rng As Range
    Dim uniqueValues As Collection
    Dim cell As Range
    Dim category As Variant
    Dim newWB As Workbook
    Dim filePrefix As String
    Dim outputFileName As String
    Dim dataRng As Range
    Dim filteredRng As Range
    Dim savePath As String
    Dim targetWS As Worksheet
    Dim firstDataRow As Range
    Dim fd As FileDialog
    
    ' Set active sheet
    Set ws = ActiveSheet
    
    ' Find last row and column
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).column
    
    ' Get column number to split
    colToSplit = Application.InputBox("분리할 열 번호를 입력하세요 (숫자, 예: 3 for Column C):", Type:=1)
    If colToSplit = 0 Then Exit Sub  ' User clicked Cancel
    
    ' Validate column number
    If colToSplit < 1 Or colToSplit > lastCol Then
        MsgBox "유효하지 않은 열 번호입니다.", vbExclamation
        Exit Sub
    End If
    
    ' Get file prefix
    filePrefix = InputBox("파일 이름 접두사를 입력하세요:")
    If filePrefix = "" Then Exit Sub  ' User clicked Cancel
    
    ' Show folder picker dialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .Title = "결과 파일을 저장할 폴더를 선택하세요"
        .AllowMultiSelect = False
        If .Show = 0 Then
            MsgBox "폴더를 선택하지 않았습니다. 매크로를 종료합니다.", vbExclamation
            Exit Sub
        End If
        savePath = .SelectedItems(1) & "\"
    End With
    
    ' Initialize unique values collection
    Set uniqueValues = New Collection
    
    ' Collect unique values
    On Error Resume Next
    For Each cell In ws.Range(ws.Cells(2, colToSplit), ws.Cells(lastRow, colToSplit))
        If Not IsEmpty(cell) And cell.Value <> "" Then
            uniqueValues.Add cell.Value, CStr(cell.Value)
        End If
    Next cell
    On Error GoTo 0
    
    ' Check if any unique values were found
    If uniqueValues.count = 0 Then
        MsgBox "선택한 열에서 데이터를 찾을 수 없습니다.", vbExclamation
        Exit Sub
    End If
    
    ' Debug message - show number of unique values found
    Debug.Print "Unique values found: " & uniqueValues.count
    Debug.Print "Save path: " & savePath
    
    ' Set reference to data range
    Set dataRng = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
    
    ' Turn off screen updating for better performance
    Application.ScreenUpdating = False
    
    ' Process each category
    Dim fileCount As Long
    fileCount = 0
    
    For Each category In uniqueValues
        ' Debug message - processing category
        Debug.Print "Processing category: " & category
        
        ' Create new workbook
        Set newWB = Workbooks.Add
        Set targetWS = newWB.Sheets(1)
        
        ' Copy header
        ws.Rows(1).Copy targetWS.Rows(1)
        
        ' Apply filter and copy data
        With dataRng
            .AutoFilter
            .AutoFilter Field:=colToSplit, Criteria1:=category
        End With
        
        ' Find and copy filtered data
        On Error Resume Next
        Set filteredRng = dataRng.SpecialCells(xlCellTypeVisible)
        
        If Not filteredRng Is Nothing Then
            filteredRng.Copy
            targetWS.Cells(1, 1).PasteSpecial xlPasteValues
            Application.CutCopyMode = False
            
            ' Create safe filename
            outputFileName = savePath & filePrefix & "_" & _
                Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace( _
                CStr(category), "\", ""), "/", ""), ":", ""), "*", ""), "?", ""), """", ""), _
                "<", ""), ">", ""), "|", ""), " ", "_") & ".xlsx"
            
            ' Debug message - saving file
            Debug.Print "Attempting to save file: " & outputFileName
            
            ' Save and close new workbook
            On Error Resume Next
            newWB.SaveAs FileName:=outputFileName
            If Err.Number = 0 Then
                fileCount = fileCount + 1
            Else
                Debug.Print "Error saving file: " & Err.Description
            End If
            On Error GoTo 0
            
            newWB.Close SaveChanges:=False
        End If
        
        ' Remove filter
        ws.AutoFilterMode = False
    Next category
    
    ' Turn screen updating back on
    Application.ScreenUpdating = True
    
    ' Show completion message with file count
    If fileCount > 0 Then
        MsgBox fileCount & "개의 파일이 다음 경로에 저장되었습니다:" & vbNewLine & savePath, vbInformation
    Else
        MsgBox "파일이 저장되지 않았습니다. 자세한 내용은 디버그 창을 확인하세요.", vbCritical
    End If
End Sub



