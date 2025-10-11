'===============================================================================
' Program Name: ExportImportManager
' Version     : v2.1
' Created     : 2025-10-11
' Modified    : 2025-10-11
' Author      : ChatGPT (User Request Based)
' Description : Export/Import Excel VBA project modules/classes/forms to/from files
'               - Enhanced error handling
'               - Duplicate module handling logic
'               - Detailed progress and result reporting
'               - Automatic VBA project access permission check
'
' Main Functions:
'   1. ExportAllModules    : Export all VBA components to files
'   2. ImportAllModules    : Import VBA components from files
'   3. VBA_Project_Trust_Info : Display trust settings guide
'
' Required Setup Before Use:
'   [File] > [Options] > [Trust Center] > [Trust Center Settings] > [Macro Settings]
'   -> Check "Trust access to the VBA project object model"
'===============================================================================

Option Explicit

'===============================================================================
' Procedure Name: Open_ExportImportManager
' Function     : Main entry point for VBA Export/Import menu
' Parameters   : None
' Returns      : None
' Description  : Display menu for user to select export/import operation
'                and execute corresponding function based on selection
'===============================================================================
Sub Open_ExportImportManager()
    On Error GoTo ErrorHandler
    
    Dim answer As VbMsgBoxResult
    
    ' Check VBA project access permission
    If Not CheckVBAProjectAccess() Then Exit Sub
    
    ' Request user to select operation
    answer = MsgBox("What would you like to do?" & vbCrLf & vbCrLf & _
                    "Yes(Y)     -> Export (Save current VBA to files)" & vbCrLf & _
                    "No(N)      -> Import (Load VBA from files)" & vbCrLf & _
                    "Cancel     -> Exit", _
                    vbYesNoCancel + vbQuestion, "VBA Export/Import Manager v2.1")
    
    ' Execute operation based on user selection
    Select Case answer
        Case vbYes
            Call ExportAllModules
        Case vbNo
            Call ImportAllModules
        Case vbCancel
            MsgBox "Operation cancelled.", vbInformation, "Notice"
    End Select
    
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred." & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Error Description: " & Err.Description, vbCritical, "Error"
End Sub

'===============================================================================
' Procedure Name: ExportAllModules
' Function     : Export all VBA components from current workbook to files
' Parameters   : None
' Returns      : None
' Description  : Export all standard modules, class modules, forms, document modules
'                from VBA project to VBA_Export folder as individual files.
'                File extensions are automatically determined by component type.
'===============================================================================
Sub ExportAllModules()
    On Error GoTo ErrorHandler
    
    Dim vbComp As Object
    Dim exportPath As String
    Dim fileExt As String
    Dim exportCount As Long
    Dim skipCount As Long
    Dim resultMsg As String
    
    ' Initialize counters
    exportCount = 0
    skipCount = 0
    
    ' Set export path (VBA_Export subfolder in same directory as workbook)
    exportPath = ThisWorkbook.Path & "\VBA_Export\"
    
    ' Create folder if it doesn't exist
    If Dir(exportPath, vbDirectory) = "" Then
        MkDir exportPath
    End If
    
    ' Loop through all VBA components
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        ' Determine file extension and export based on component type
        Select Case vbComp.Type
            Case 1  ' vbext_ct_StdModule: Standard Module
                fileExt = ".bas"
                vbComp.Export exportPath & vbComp.Name & fileExt
                exportCount = exportCount + 1
                
            Case 2  ' vbext_ct_ClassModule: Class Module
                fileExt = ".cls"
                vbComp.Export exportPath & vbComp.Name & fileExt
                exportCount = exportCount + 1
                
            Case 3  ' vbext_ct_MSForm: User Form
                fileExt = ".frm"
                vbComp.Export exportPath & vbComp.Name & fileExt
                exportCount = exportCount + 1
                
            Case 100 ' vbext_ct_Document: Document Module (Sheet, ThisWorkbook, etc.)
                fileExt = ".bas"
                vbComp.Export exportPath & vbComp.Name & fileExt
                exportCount = exportCount + 1
                
            Case Else
                ' Skip other types
                skipCount = skipCount + 1
        End Select
    Next vbComp
    
    ' Build result message
    resultMsg = "Export completed successfully!" & vbCrLf & vbCrLf & _
                "Export Results:" & vbCrLf & _
                "   - Success: " & exportCount & " files" & vbCrLf
    
    If skipCount > 0 Then
        resultMsg = resultMsg & "   - Skipped: " & skipCount & " files" & vbCrLf
    End If
    
    resultMsg = resultMsg & vbCrLf & "Save Location:" & vbCrLf & "   " & exportPath
    
    MsgBox resultMsg, vbInformation, "Export Complete"
    
    Exit Sub

ErrorHandler:
    MsgBox "Error occurred during export." & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Error Description: " & Err.Description & vbCrLf & vbCrLf & _
           "Files exported so far: " & exportCount, _
           vbCritical, "Export Error"
End Sub

'===============================================================================
' Procedure Name: ImportAllModules
' Function     : Import all VBA files from VBA_Export folder to current project
' Parameters   : None
' Returns      : None
' Description  : Find .bas, .cls, .frm files and add them to VBA project.
'                If module with same name already exists, ask user for confirmation
'                before removing existing module and importing new one.
'===============================================================================
Sub ImportAllModules()
    On Error GoTo ErrorHandler
    
    Dim importPath As String
    Dim fileName As String
    Dim importCount As Long
    Dim skipCount As Long
    Dim replaceCount As Long
    Dim resultMsg As String
    Dim overwriteAll As Boolean
    Dim answer As VbMsgBoxResult
    
    ' Initialize counters
    importCount = 0
    skipCount = 0
    replaceCount = 0
    overwriteAll = False
    
    ' Set import path
    importPath = ThisWorkbook.Path & "\VBA_Export\"
    
    ' Check if folder exists
    If Dir(importPath, vbDirectory) = "" Then
        MsgBox "VBA_Export folder does not exist." & vbCrLf & vbCrLf & _
               "Please run 'Export' first or create folder at:" & vbCrLf & _
               importPath, vbExclamation, "Folder Not Found"
        Exit Sub
    End If
    
    ' Check if there are files to import
    If Dir(importPath & "*.bas") = "" And _
       Dir(importPath & "*.cls") = "" And _
       Dir(importPath & "*.frm") = "" Then
        MsgBox "No VBA files found to import." & vbCrLf & vbCrLf & _
               "Path: " & importPath, vbExclamation, "No Files Found"
        Exit Sub
    End If
    
    ' 1. Import .bas files (Standard Modules)
    fileName = Dir(importPath & "*.bas")
    Do While fileName <> ""
        If ImportSingleModule(importPath & fileName, overwriteAll, replaceCount) Then
            importCount = importCount + 1
        Else
            skipCount = skipCount + 1
        End If
        fileName = Dir
    Loop
    
    ' 2. Import .cls files (Class Modules)
    fileName = Dir(importPath & "*.cls")
    Do While fileName <> ""
        If ImportSingleModule(importPath & fileName, overwriteAll, replaceCount) Then
            importCount = importCount + 1
        Else
            skipCount = skipCount + 1
        End If
        fileName = Dir
    Loop
    
    ' 3. Import .frm files (User Forms)
    fileName = Dir(importPath & "*.frm")
    Do While fileName <> ""
        If ImportSingleModule(importPath & fileName, overwriteAll, replaceCount) Then
            importCount = importCount + 1
        Else
            skipCount = skipCount + 1
        End If
        fileName = Dir
    Loop
    
    ' Build result message
    resultMsg = "Import completed successfully!" & vbCrLf & vbCrLf & _
                "Import Results:" & vbCrLf & _
                "   - Added: " & importCount & " files" & vbCrLf
    
    If replaceCount > 0 Then
        resultMsg = resultMsg & "   - Replaced: " & replaceCount & " files" & vbCrLf
    End If
    
    If skipCount > 0 Then
        resultMsg = resultMsg & "   - Skipped: " & skipCount & " files" & vbCrLf
    End If
    
    MsgBox resultMsg, vbInformation, "Import Complete"
    
    Exit Sub

ErrorHandler:
    MsgBox "Error occurred during import." & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Error Description: " & Err.Description & vbCrLf & vbCrLf & _
           "Files imported so far: " & importCount, _
           vbCritical, "Import Error"
End Sub

'===============================================================================
' Function Name: ImportSingleModule
' Function     : Import single module file (with duplicate handling)
' Parameters   : filePath (String) - Full path of file to import
'                overwriteAll (Boolean) - Overwrite all flag (ByRef)
'                replaceCount (Long) - Count of replaced modules (ByRef)
' Returns      : Boolean - True if successful, False if failed
' Description  : Check if module with same name exists before importing file
'                and ask user for overwrite confirmation if necessary.
'===============================================================================
Private Function ImportSingleModule(ByVal filePath As String, _
                                     ByRef overwriteAll As Boolean, _
                                     ByRef replaceCount As Long) As Boolean
    On Error GoTo ErrorHandler
    
    Dim moduleName As String
    Dim vbComp As Object
    Dim answer As VbMsgBoxResult
    Dim moduleExists As Boolean
    
    ' Extract module name from filename (remove extension)
    moduleName = Mid(Dir(filePath), 1, InStrRev(Dir(filePath), ".") - 1)
    
    ' Check if module with same name already exists
    moduleExists = False
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        If vbComp.Name = moduleName Then
            moduleExists = True
            Exit For
        End If
    Next vbComp
    
    ' If module already exists
    If moduleExists Then
        ' If "overwrite all" flag is not set, ask user for confirmation
        If Not overwriteAll Then
            answer = MsgBox("Module '" & moduleName & "' already exists." & vbCrLf & vbCrLf & _
                           "Do you want to overwrite it?" & vbCrLf & vbCrLf & _
                           "Yes(Y)     -> Overwrite this module only" & vbCrLf & _
                           "No(N)      -> Skip this module" & vbCrLf & _
                           "Cancel     -> Overwrite all remaining modules", _
                           vbYesNoCancel + vbQuestion, "Duplicate Module Found")
            
            Select Case answer
                Case vbYes
                    ' Remove existing module
                    ThisWorkbook.VBProject.VBComponents.Remove vbComp
                    replaceCount = replaceCount + 1
                    
                Case vbNo
                    ' Skip this module
                    ImportSingleModule = False
                    Exit Function
                    
                Case vbCancel
                    ' Enable "overwrite all" mode
                    overwriteAll = True
                    ThisWorkbook.VBProject.VBComponents.Remove vbComp
                    replaceCount = replaceCount + 1
            End Select
        Else
            ' In "overwrite all" mode, automatically remove existing module
            ThisWorkbook.VBProject.VBComponents.Remove vbComp
            replaceCount = replaceCount + 1
        End If
    End If
    
    ' Import module from file
    ThisWorkbook.VBProject.VBComponents.Import filePath
    ImportSingleModule = True
    
    Exit Function

ErrorHandler:
    MsgBox "Error occurred while importing module" & vbCrLf & _
           "File: " & filePath & vbCrLf & _
           "Error: " & Err.Description, vbExclamation
    ImportSingleModule = False
End Function

'===============================================================================
' Function Name: CheckVBAProjectAccess
' Function     : Check VBA project access permission
' Parameters   : None
' Returns      : Boolean - True if accessible, False if not
' Description  : Check if VBA project object model is accessible.
'                Display setup instructions if access is not available.
'===============================================================================
Private Function CheckVBAProjectAccess() As Boolean
    On Error Resume Next
    
    Dim testAccess As Object
    Dim accessGranted As Boolean
    
    ' Try to access VBA project
    Set testAccess = ThisWorkbook.VBProject.VBComponents
    accessGranted = (Err.Number = 0)
    
    On Error GoTo 0
    
    ' If access is not granted, display instruction message
    If Not accessGranted Then
        MsgBox "Cannot access VBA project." & vbCrLf & vbCrLf & _
               "Please enable the following setting:" & vbCrLf & vbCrLf & _
               "1. [File] -> [Options] -> [Trust Center]" & vbCrLf & _
               "2. Click [Trust Center Settings] button" & vbCrLf & _
               "3. Select [Macro Settings] menu" & vbCrLf & _
               "4. Check 'Trust access to the VBA project object model'" & vbCrLf & _
               "5. Restart Excel", _
               vbExclamation, "Access Permission Required"
        CheckVBAProjectAccess = False
    Else
        CheckVBAProjectAccess = True
    End If
End Function

'===============================================================================
' Procedure Name: VBA_Project_Trust_Info
' Function     : Display VBA project trust settings guide
' Parameters   : None
' Returns      : None
' Description  : Display instruction message for VBA project object model
'                access trust settings.
'===============================================================================
Sub VBA_Project_Trust_Info()
    MsgBox "VBA Project Trust Settings Guide" & vbCrLf & vbCrLf & _
           "To use this macro, the following setting is required:" & vbCrLf & vbCrLf & _
           "Setup Steps:" & vbCrLf & _
           "1. [File] -> [Options] -> [Trust Center]" & vbCrLf & _
           "2. Click [Trust Center Settings] button" & vbCrLf & _
           "3. Select [Macro Settings] menu" & vbCrLf & _
           "4. Check 'Trust access to the VBA project object model'" & vbCrLf & _
           "5. Click [OK] and restart Excel" & vbCrLf & vbCrLf & _
           "This setting is required only once.", _
           vbInformation, "Trust Settings Guide"
End Sub