VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
    'PURPOSE: Call subroutine to activate reference libraries when opening worksheet
    
    On Error Resume Next
    'Ensure reference libraries are activated
    Call Source_Control.ActivateReferenceLibrary
    
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    
    'PURPOSE: When saving file trigger code extraction from workbook into a "Code" folder and creation of test file
    
    Dim OrigName As String, TestFile As String
    Dim wb As Workbook, wb2 As Workbook
    
    On Error GoTo ErrHandler
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    If ThisWorkbook.Name <> "Test_File.xlsm" Then
        'Extract out Modules, sheet code and forms into code folder
        Call Source_Control.ExportCode
        DoEvents
        
        'Create copy of file in same folder
        Set wb = ThisWorkbook
    
        OrigName = ThisWorkbook.fullName
        TestFile = Source_Control.GetWorkbookPath & "Test_File.xlsm"
        
        If Len(Dir(TestFile)) <> 0 Then Kill (TestFile)
        ActiveWorkbook.SaveCopyAs TestFile
        DoEvents
        
        'Open Test_File
        Set wb2 = Source_Control.GetWorkbook(TestFile)
        
        If wb2 Is Nothing Then
            MsgBox ("Failed to open 'Test_File.xlsm'")
            Exit Sub
        End If
        
        'Get rid of code in Test_File
        wb2.Activate
        Source_Control.RemoveCode
        
        'Close and save changes to Test_File
        wb2.Close True
        Set wb2 = Nothing
        wb.Activate
    End If
    
ErrHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
End Sub


