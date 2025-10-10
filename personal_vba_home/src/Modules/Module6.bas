Attribute VB_Name = "Module6"
Sub 모든이미지삭제()
    Dim pic As Picture
    ' 현재 활성화된 시트의 모든 이미지 삭제
    For Each pic In ActiveSheet.Pictures
        pic.Delete
    Next pic
End Sub

