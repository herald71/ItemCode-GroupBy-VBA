Attribute VB_Name = "Module12"
Sub 썸네일삽입()
    Dim urlColumn As String
    Dim insertColumn As String
    Dim cell As Range
    Dim img As Picture
    Dim imageUrl As String
    Dim regex As Object

    ' 정규식을 사용하여 영문자만 허용하도록 설정
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = False
    regex.Pattern = "^[A-Za-z]+$"

    ' 사용자로부터 이미지 URL이 있는 열과 이미지를 삽입할 열 정보를 입력받습니다.
    urlColumn = InputBox("이미지 URL이 포함된 열을 입력하세요. 예: G")
    insertColumn = InputBox("이미지를 삽입할 열을 입력하세요. 예: J")

    ' 영문자만 입력되었는지 확인합니다.
    If Not regex.Test(urlColumn) Or Not regex.Test(insertColumn) Then
        MsgBox "열 이름은 반드시 영문자만 입력해야 합니다.", vbExclamation
        Exit Sub
    End If

    On Error GoTo ErrorHandler ' 오류 발생 시 점프할 위치 설정
    
    ' 입력받은 열을 기반으로 URL을 검토합니다.
    For Each cell In Range(urlColumn & "2:" & urlColumn & Cells(Rows.count, urlColumn).End(xlUp).Row)
        imageUrl = cell.Value
        
        ' 이미지 URL이 빈 문자열이 아닌 경우에만 이미지 삽입을 시도합니다.
        If imageUrl <> "" Then
            ' 이미지를 워크시트에 삽입합니다.
            Set img = ActiveSheet.Pictures.Insert(imageUrl)
            
            With img
                ' 이미지를 사용자가 지정한 열에 위치시킵니다.
                .Top = Cells(cell.Row, insertColumn).Top
                .Left = Cells(cell.Row, insertColumn).Left
                
                ' 셀의 높이를 50으로 설정합니다.
                Cells(cell.Row, insertColumn).RowHeight = 50
                
                ' 이미지의 원래 가로세로 비율을 계산합니다.
                Dim origRatio As Double
                origRatio = .Width / .Height
                
                ' 삽입 열 셀의 가로세로 비율을 계산합니다.
                Dim cellRatio As Double
                cellRatio = Cells(cell.Row, insertColumn).Width / Cells(cell.Row, insertColumn).Height
                
                ' 셀 비율에 따라 이미지의 크기를 조정합니다.
                If origRatio > cellRatio Then
                    .Width = Cells(cell.Row, insertColumn).Width
                    .Height = .Width / origRatio
                Else
                    .Height = Cells(cell.Row, insertColumn).Height
                    .Width = .Height * origRatio
                End If
                
                ' 이미지를 셀의 중앙에 배치합니다.
                .Top = Cells(cell.Row, insertColumn).Top + (Cells(cell.Row, insertColumn).Height - .Height) / 2
                .Left = Cells(cell.Row, insertColumn).Left + (Cells(cell.Row, insertColumn).Width - .Width) / 2
            End With
        End If
    Next cell
    
    Exit Sub ' 정상 종료

ErrorHandler:
    MsgBox "오류가 발생했습니다: " & Err.Description, vbCritical
End Sub

