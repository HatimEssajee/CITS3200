Attribute VB_Name = "OpenForm"
Option Explicit

Sub OpenForm()
    'PURPOSE: Determines dimensions of register table and loads first userform
    
    'Reference register table
    Set RegTable = ThisWorkbook.Sheets("Register").ListObjects("Register")
    
    'Store current username in memory
    'Source: https://www.excelsirji.com/vba-code-to-get-logged-in-user-name/
    Username = Application.Username
    
    'Source: https://officetricks.com/excel-vba-get-username-windows-system/
    'Username = ThisWorkbook.BuiltinDocumentProperties("Author")
    
    
    'Force default starting rowIndex for empty form and tickbox checked
    RowIndex = -1
    Tick = True
    
    'Set initial location
    UserFormTopPos = Application.Top + 25
    UserFormLeftPos = Application.Left + 25
    
    'Display Project Form UserForm
    'Source: https://www.contextures.com/xlUserForm02.html
    form00_Nav.Show False
    
End Sub

