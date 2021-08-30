Attribute VB_Name = "Source_Control"
Option Explicit

Sub ImportCode()

    'PURPOSE: Import .bas and .frm VB Objects into test_file from Code folder
    'SOURCE: http://www.vbaexpress.com/forum/showthread.php?36969-Solved-Automatically-import-a-bas-module
    'SOURCE: https://exceloffthegrid.com/vba-code-loop-files-folder-sub-folders/

    Dim VBProj As Object 'VBIDE.VBProject
    Dim codeFolder As String, codeFiles As String, frmFiles As String
    Dim code_file As String, filename As String
    
    Call ActivateReferenceLibrary
    Call RemoveCode
    
    'Close VB Editor
    Application.VBE.MainWindow.Visible = False
    
    codeFolder = CombinePaths(GetWorkbookPath, "Code") & "\"
    code_file = Dir(codeFolder)
    
    Set VBProj = Nothing
    
    On Error Resume Next
    
    Set VBProj = ThisWorkbook.VBProject
    
    On Error GoTo 0
    
    'Check if trust settings allow macros
    If VBProj Is Nothing Then
        MsgBox "Can't continue--I'm not trusted!"
        Exit Sub
    End If
    
    'Check if vb project already has other modules
    If ThisWorkbook.VBProject.VBComponents.Count > 10 Then
        MsgBox "Can't continue--already have code imported!"
        Exit Sub
    End If
    
    
    'Loop through .bas files
    While code_file <> ""
        
        If Not (code_file = "Source_Control.bas" _
            Or EndsWith(code_file, ".cls") Or EndsWith(code_file, ".frx")) Then
            filename = codeFolder & code_file
            VBProj.VBComponents.Import filename
        End If
        
        code_file = Dir
    Wend
    
    'Open VB Editor
    Application.VBE.MainWindow.Visible = True
    
End Sub

Sub RemoveCode()
    'PURPOSE: Removes all VBA Project objects apart from "source_control" module
    'Source: https://stackoverflow.com/questions/18518493/remove-all-vba-modules-from-excel-file
    On Error Resume Next
    
    Dim Element As Object
    
    For Each Element In ActiveWorkbook.VBProject.VBComponents
        If Not (Element.Name = "Source_Control") Then
            ActiveWorkbook.VBProject.VBComponents.Remove Element
        End If
    Next

End Sub


Sub ExportCode()
    
    'PURPOSE: Export out all VBA Poject components from workbook into "Code" folder
    'SOURCE: https://stackoverflow.com/questions/49724/programmatically-extract-macro-vba-code-from-word-2007-docs/49796#49796
    
'    If Not CanAccessVBOM Then Exit Sub ' Exit if access to VB object model is not allowed
'    If (ThisWorkbook.VBProject.VBE.ActiveWindow Is Nothing) Then
'        Exit Sub ' Exit if VBA window is not open
'    End If

    Dim comp As VBComponent
    Dim codeFolder As String
    Dim myForm As UserForm
    
    codeFolder = CombinePaths(GetWorkbookPath, "Code")
    On Error Resume Next
    MkDir codeFolder
    On Error GoTo 0
    Dim filename As String

    For Each comp In ThisWorkbook.VBProject.VBComponents
        Select Case comp.Type
            Case vbext_ct_ClassModule
                filename = CombinePaths(codeFolder, comp.Name & ".cls")
                DeleteFile filename
                comp.Export filename
                DoEvents
                
            Case vbext_ct_StdModule
                filename = CombinePaths(codeFolder, comp.Name & ".bas")
                DeleteFile filename
                comp.Export filename
                DoEvents
                
            Case vbext_ct_MSForm
                filename = CombinePaths(codeFolder, comp.Name & ".frm")
                DeleteFile filename
                comp.Export filename
                DoEvents
                
            Case vbext_ct_Document
                filename = CombinePaths(codeFolder, comp.Name & ".cls")
                DeleteFile filename
                comp.Export filename
                DoEvents
        End Select
    Next
    
    'Unload all user forms
    For Each myForm In UserForms
        Unload myForm
    Next
    
    'Save backup file
    If ThisWorkbook.Name <> "Backup_File.xlsm" Then
        Call Save_Backup
    End If
    
End Sub

Sub Save_Backup()
    'PURPOSE: Save a copy of workbook
    
    Dim BackupFile As String
    
    On Error GoTo ErrHandler
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    BackupFile = Source_Control.GetWorkbookPath & "Backup_File.xlsm"
    
    If Len(Dir(BackupFile)) <> 0 Then Kill (BackupFile)
        ThisWorkbook.SaveCopyAs BackupFile
        DoEvents
    
ErrHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
End Sub

Function CanAccessVBOM() As Boolean
    
    'PURPOSE: Check resgistry to see if we can access the VB object model
    'SOURCE: https://stackoverflow.com/questions/49724/programmatically-extract-macro-vba-code-from-word-2007-docs/49796#49796
    
    Dim wsh As Object
    Dim str1 As String
    Dim AccessVBOM As Long

    Set wsh = CreateObject("WScript.Shell")
    str1 = "HKEY_CURRENT_USER\Software\Microsoft\Office\" & _
        Application.Version & "\Excel\Security\AccessVBOM"
    On Error Resume Next
    AccessVBOM = wsh.RegRead(str1)
    Set wsh = Nothing
    CanAccessVBOM = (AccessVBOM = 1)
    
End Function


Sub DeleteFile(filename As String)
    
    'PURPOSE: Delete file
    'SOURCE: https://stackoverflow.com/questions/49724/programmatically-extract-macro-vba-code-from-word-2007-docs/49796#49796
    
    On Error Resume Next
    Kill filename
End Sub

Function GetWorkbookPath() As String

    'PURPOSE: Extract the file directory path of current workbook
    'SOURCE: https://stackoverflow.com/questions/49724/programmatically-extract-macro-vba-code-from-word-2007-docs/49796#49796
    
    Dim fullName As String
    Dim wrkbookName As String
    Dim pos As Long

    wrkbookName = ThisWorkbook.Name
    fullName = ThisWorkbook.fullName

    pos = InStr(1, fullName, wrkbookName, vbTextCompare)

    GetWorkbookPath = Left$(fullName, pos - 1)
    
End Function

Function CombinePaths(ByVal Path1 As String, ByVal Path2 As String) As String
    
    'PURPOSE: Combine current workbook file directory path with other folder path
    'SOURCE: https://stackoverflow.com/questions/49724/programmatically-extract-macro-vba-code-from-word-2007-docs/49796#49796
    
    If Not EndsWith(Path1, "\") Then
        Path1 = Path1 & "\"
    End If
    CombinePaths = Path1 & Path2
End Function


Public Function EndsWith(str As String, ending As String) As Boolean

    'PURPOSE: Check string ends with substring
    'SOURCE: https://stackoverflow.com/questions/49724/programmatically-extract-macro-vba-code-from-word-2007-docs/49796#49796
    
     Dim endingLen As Integer
     endingLen = Len(ending)
     EndsWith = (Right(Trim(UCase(str)), endingLen) = UCase(ending))
     
End Function

Public Function StartsWith(str As String, start As String) As Boolean
     
    'PURPOSE: Check string starts with substring
    'SOURCE: https://stackoverflow.com/questions/49724/programmatically-extract-macro-vba-code-from-word-2007-docs/49796#49796
    
     Dim startLen As Integer
     startLen = Len(start)
     StartsWith = (Left(Trim(UCase(str)), startLen) = UCase(start))
End Function

Public Function GetWorkbook(ByVal sFullName As String) As Workbook
    
    'PURPOSE: Open file if full file path provided
    'SOURCE: https://stackoverflow.com/questions/9373082/detect-whether-excel-workbook-is-already-open
    Dim sFile As String
    Dim wbReturn As Workbook
    
    sFile = Dir(sFullName)

    On Error Resume Next
        Set wbReturn = Workbooks(sFile)

        If wbReturn Is Nothing Then
            Set wbReturn = Workbooks.Open(sFullName, ReadOnly:=False, _
                            UpdateLinks:=3, IgnoreReadOnlyRecommended:=True)
        End If
    On Error GoTo 0

    Set GetWorkbook = wbReturn

End Function

Sub ActivateReferenceLibrary()

'PURPOSE: Show How To Activate Specific Object Libraries
'SOURCE: https://www.thespreadsheetguru.com/the-code-vault/2014/5/18/activate-object-library-references-in-visual-basic-editor

'Error Handler in Case Reference is Already Activated
  On Error Resume Next
    
    'Activate Microsoft Scripting Runtime Library (version 1.0)
        ThisWorkbook.VBProject.References.AddFromGuid _
          GUID:="{420B2830-E718-11CF-893D-00A0C9054228}", _
          Major:=0, Minor:=0
    
    'Activate Visual Basic for Applications Extensibility Library (version 5.3)
      ThisWorkbook.VBProject.References.AddFromGuid _
        GUID:="{0002E157-0000-0000-C000-000000000046}", _
        Major:=0, Minor:=0 'Use zeroes to default to latest version

'Reset Error Handler
  On Error GoTo 0

End Sub
