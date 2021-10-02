VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)

    'PURPOSE: When saving file a backup file is created in the Backup_Folder that is created
    'Up to 7 of the earliest backups are retained

    Dim fso As Object
    Dim fil As Object
    Dim oldfile As Object
    Dim wb As Workbook, wb2 As Workbook
    Dim BackUpFolder As String, BackupFile As String
    
    On Error GoTo ErrHandler
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    BackUpFolder = GetWorkbookPath & "Backup_Folder"
    BackupFile = BackUpFolder & "\Back-up Study Tracker " & Format(Now, "yyyy-mm-dd hhmm") & ".xlsm"
    
    'Create backup folder if doesn't exist
    'SOURCE: https://www.techonthenet.com/excel/formulas/mkdir.php
    If Len(Dir(BackUpFolder, vbDirectory)) = 0 Then
        MkDir BackUpFolder
    End If
    
    'Save a new copy of backup file
    ThisWorkbook.SaveCopyAs BackupFile
    
    'PURPOSE: Delete files in backupfolder
    'SOURCE: https://stackoverflow.com/questions/27550992/excel-vba-leave-5-newest-backups-and-delete-the-rest
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Do Until fso.GetFolder(BackUpFolder).Files.count <= 7
        For Each fil In fso.GetFolder(BackUpFolder).Files
            'Checks to see if this file is older than the oldest file thus far
            If oldfile Is Nothing Then Set oldfile = fil
            If oldfile.DateLastModified > fil.DateLastModified Then Set oldfile = fil
        Next fil
        fso.DeleteFile oldfile, True
        Set oldfile = Nothing
    Loop
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    Exit Sub
    
ErrHandler:
    
    MsgBox ("Failed to save backup file")
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

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
