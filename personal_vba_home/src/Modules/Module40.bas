Attribute VB_Name = "Module40"
'����������������������������������������������������������������������������������������������
' ���α׷��� : PersonalVBA_GitSync
' ����       : v1.0
' �ۼ�����   : 2025-10-11
' �ۼ���     : ChatGPT (VBA ���� �����)
' ����       : PERSONAL.XLSB�� ���/Ŭ����/����
'              ���Ϸ� ��������, �ٽ� �������� �ڵ�ȭ.
'              Git ������ ���� src ���� ������ ����.
'����������������������������������������������������������������������������������������������
Option Explicit

'=== ����� ����: Git ����� ���(�ʼ� ����) ======================
Private Const BASE_PATH As String = "C:\Users\owner\Documents\source\Excel_macro\personal_vba"
'================================================================

' ���� ���� ����(���� �� ���� ������): �� ��� �̸��� �����ϰ� ����
Private Const BOOTSTRAP_MODULE As String = "modPersonalVBA_ExportImport"

' VBComponent.Type ���(���� ���ε���)
Private Const vbext_ct_StdModule As Long = 1
Private Const vbext_ct_ClassModule As Long = 2
Private Const vbext_ct_MSForm As Long = 3
Private Const vbext_ct_Document As Long = 100

' ========== ���� ��ũ��(��ũ�� â���� ����) ==========
Public Sub Export_Personal_ToSrc()
    On Error GoTo EH
    Dim wb As Workbook, vbProj As Object
    Dim vbComp As Object, fso As Object
    Dim outPath As String

    Set wb = Workbooks("PERSONAL.XLSB")
    Set vbProj = wb.VBProject
    If Not CanAccessVBProject(vbProj) Then
        MsgBox "VBA ������Ʈ�� ������ �� �����ϴ�." & vbCrLf & _
               "���� �������� 'VBA ��ü �𵨿� ���� �ŷ�'�� �Ѽ���.", vbExclamation
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

    MsgBox "�������� �Ϸ�: " & cnt & "�� ������Ʈ" & vbCrLf & _
           BASE_PATH & "\src\* �� ����Ǿ����ϴ�.", vbInformation
    Exit Sub
EH:
    MsgBox "�������� ����: " & Err.Number & " - " & Err.Description, vbCritical
End Sub

Public Sub Import_Src_To_Personal()
    On Error GoTo EH
    Dim wb As Workbook, vbProj As Object, fso As Object
    Dim imported As Long

    Set wb = Workbooks("PERSONAL.XLSB")
    Set vbProj = wb.VBProject
    If Not CanAccessVBProject(vbProj) Then
        MsgBox "VBA ������Ʈ�� ������ �� �����ϴ�." & vbCrLf & _
               "���� �������� 'VBA ��ü �𵨿� ���� �ŷ�'�� �Ѽ���.", vbExclamation
        Exit Sub
    End If

    EnsureFolders
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' 1) ���� �ڵ� ����(���� ��� ����, ��Ʈ��Ʈ�� ��� ����)
    RemoveAllRemovable vbProj, Array(BOOTSTRAP_MODULE)

    ' 2) ��������(���/Ŭ����/��)
    imported = 0
    imported = imported + ImportFolder(vbProj, BASE_PATH & "\src\Modules\", Array("bas"))
    imported = imported + ImportFolder(vbProj, BASE_PATH & "\src\Classes\", Array("cls"))
    imported = imported + ImportFolder(vbProj, BASE_PATH & "\src\Forms\", Array("frm"))

    MsgBox "�������� �Ϸ�: " & imported & "�� ������Ʈ �ݿ�", vbInformation
    Exit Sub
EH:
    MsgBox "�������� ����: " & Err.Number & " - " & Err.Description, vbCritical
End Sub

' ========== ���� ��ƾ ==========
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
                ' ���� ���(��Ʈ/ThisWorkbook)�� �ǵ帮�� ����
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
            ' ���� �̸��� ������ ����(��Ʈ��Ʈ�� ��� ����)
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
    MsgBox "Import ����: " & folderPath & vbCrLf & Err.Number & " - " & Err.Description, vbCritical
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


