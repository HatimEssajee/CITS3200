Attribute VB_Name = "OpenForm"
Option Explicit

Sub OpenForm()
    'PURPOSE: Determines dimensions of register table and loads first userform
    
    'Reference register table
    Set RegTable = ThisWorkbook.Sheets("Register").ListObjects("Register")
    
    'Store current username in memory
    'Source: https://officetricks.com/excel-vba-get-username-windows-system/
    Username = ThisWorkbook.BuiltinDocumentProperties("Author")
    
    'Force default starting rowIndex for empty form and tickbox checked
    RowIndex = -1
    Tick = True
    
    'Set initial location
    UserFormLeftPos = Application.Top + 25
    UserFormTopPos = Application.Left + 25
    
    'Display Project Form UserForm
    'Source: https://www.contextures.com/xlUserForm02.html
    form00_Nav.Show False
    
End Sub

