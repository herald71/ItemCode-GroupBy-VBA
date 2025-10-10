Attribute VB_Name = "Module40"
'───────────────────────────────────────────────
' 프로그램명 : PersonalVBA_GitSync
' 버전       : v1.0
' 작성일자   : 2025-10-11
' 작성자     : ChatGPT (VBA 관리 도우미)
' 설명       : PERSONAL.XLSB의 모듈/클래스/폼을
'              파일로 내보내고, 다시 가져오는 자동화.
'              Git 관리가 쉬운 src 폴더 구조로 정리.
'───────────────────────────────────────────────
Option Explicit

'=== 사용자 설정: Git 저장소 경로(필수 변경) ======================
Private Const BASE_PATH As String = "C:\Users\owner\Documents\source\Excel_macro\personal_vba"
'================================================================

' 내부 예약 모듈명(실행 중 제거 방지용): 이 모듈 이름과 동일하게 유지
Private Const BOOTSTRAP_MODULE As String = "modPersonalVBA_ExportImport"

' VBComponent.Type 상수(늦은 바인딩용)
Private Const vbext_ct_StdModule As Long = 1
Private Const vbext_ct_ClassModule As Long = 2
Private Const vbext_ct_MSForm As Long = 3
Private Const vbext_ct_Document As Long = 100

' ========== 공개 매크로(매크로 창에서 실행) ==========
Public Sub Export_Personal_ToSrc()
    On Error GoTo EH
    Dim wb As Workbook, vbProj As Object
    Dim vbComp As Object, fso As Object
    Dim outPath As String

    Set wb = Workbooks("PERSONAL.XLSB")
    Set vbProj = wb.VBProject
    If Not CanAccessVBProject(vbProj) Then
        MsgBox "VBA 프로젝트에 접근할 수 없습니다." & vbCrLf & _
               "보안 설정에서 'VBA 개체 모델에 대한 신뢰'를 켜세요.", vbExclamation
        Exit Sub
    End If

    EnsureFolders

    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim cnt As Long: cnt = 0

    For Each vbComp In vbProj.VBComponents
        Select Case vbComp.Type
            Case vbext_ct_StdModule
                outPath = BASE_PATH & "\src\Modules\" & vbComp.Name & ".bas"
            Case vbext_ct_ClassModule
                outPath = BASE_PATH & "\src\Classes\" & vbComp.Name & ".cls"
            Case vbext_ct_MSForm
                outPath = BASE_PATH & "\src\Forms\" & vbComp.Name & ".frm"
            Case Else
                outPath = ""
        End Select

        If Len(outPath) > 0 Then
            vbComp.Export outPath
            cnt = cnt + 1
        End If
    Next

    MsgBox "내보내기 완료: " & cnt & "개 컴포넌트" & vbCrLf & _
           BASE_PATH & "\src\* 에 저장되었습니다.", vbInformation
    Exit Sub
EH:
    MsgBox "내보내기 오류: " & Err.Number & " - " & Err.Description, vbCritical
End Sub

Public Sub Import_Src_To_Personal()
    On Error GoTo EH
    Dim wb As Workbook, vbProj As Object, fso As Object
    Dim imported As Long

    Set wb = Workbooks("PERSONAL.XLSB")
    Set vbProj = wb.VBProject
    If Not CanAccessVBProject(vbProj) Then
        MsgBox "VBA 프로젝트에 접근할 수 없습니다." & vbCrLf & _
               "보안 설정에서 'VBA 개체 모델에 대한 신뢰'를 켜세요.", vbExclamation
        Exit Sub
    End If

    EnsureFolders
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' 1) 기존 코드 정리(문서 모듈 제외, 부트스트랩 모듈 제외)
    RemoveAllRemovable vbProj, Array(BOOTSTRAP_MODULE)

    ' 2) 가져오기(모듈/클래스/폼)
    imported = 0
    imported = imported + ImportFolder(vbProj, BASE_PATH & "\src\Modules\", Array("bas"))
    imported = imported + ImportFolder(vbProj, BASE_PATH & "\src\Classes\", Array("cls"))
    imported = imported + ImportFolder(vbProj, BASE_PATH & "\src\Forms\", Array("frm"))

    MsgBox "가져오기 완료: " & imported & "개 컴포넌트 반영", vbInformation
    Exit Sub
EH:
    MsgBox "가져오기 오류: " & Err.Number & " - " & Err.Description, vbCritical
End Sub

' ========== 보조 루틴 ==========
Private Sub EnsureFolders()
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim paths As Variant, p As Variant
    paths = Array( _
        BASE_PATH, _
        BASE_PATH & "\src", _
        BASE_PATH & "\src\Modules", _
        BASE_PATH & "\src\Classes", _
        BASE_PATH & "\src\Forms" _
    )
    For Each p In paths
        If Not fso.FolderExists(p) Then fso.CreateFolder p
    Next
End Sub

Private Function CanAccessVBProject(vbProj As Object) As Boolean
    On Error Resume Next
    Dim n As Long: n = vbProj.VBComponents.count
    CanAccessVBProject = (Err.Number = 0)
    Err.Clear
End Function

Private Sub RemoveAllRemovable(vbProj As Object, ByVal SkipNames As Variant)
    On Error Resume Next
    Dim comp As Object
    For Each comp In vbProj.VBComponents
        Select Case comp.Type
            Case vbext_ct_StdModule, vbext_ct_ClassModule, vbext_ct_MSForm
                If Not IsInArray(comp.Name, SkipNames) Then
                    vbProj.VBComponents.Remove comp
                End If
            Case Else
                ' 문서 모듈(시트/ThisWorkbook)은 건드리지 않음
        End Select
    Next
End Sub

Private Function ImportFolder(vbProj As Object, ByVal folderPath As String, ByVal exts As Variant) As Long
    On Error GoTo EH
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then Exit Function

    Dim file As Object, ext As String, nameOnly As String
    Dim cnt As Long: cnt = 0

    For Each file In fso.GetFolder(folderPath).Files
        ext = LCase(fso.GetExtensionName(file.Name))
        If IsInArray(ext, exts) Then
            nameOnly = fso.GetBaseName(file.Name)
            ' 동일 이름이 있으면 제거(부트스트랩 모듈 제외)
            If nameOnly <> BOOTSTRAP_MODULE Then
                On Error Resume Next
                vbProj.VBComponents.Remove vbProj.VBComponents(nameOnly)
                On Error GoTo EH
            End If
            vbProj.VBComponents.Import file.Path
            cnt = cnt + 1
        End If
    Next
    ImportFolder = cnt
    Exit Function
EH:
    MsgBox "Import 오류: " & folderPath & vbCrLf & Err.Number & " - " & Err.Description, vbCritical
End Function

Private Function IsInArray(val As String, arr As Variant) As Boolean
    Dim v As Variant
    For Each v In arr
        If StrComp(CStr(v), CStr(val), vbTextCompare) = 0 Then
            IsInArray = True
            Exit Function
        End If
    Next
End Function


