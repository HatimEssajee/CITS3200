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
    FC_Tick = True
    SAG_Tick = True
    
    'Set initial location
    UserFormTopPos = Application.Top + 25
    UserFormLeftPos = Application.Left + 25
    
    'Correct array used to guide what test to apply for each register field
    '0 if skip, 1 has to be filled, 2 if has to be text, 3 if has to be date
    Correct = Array(2, 1, 1, 1, 1, 1, 0, 0, 0, _
                    3, 3, 3, 3, 3, 0, 0, 0, _
                    3, 3, 2, 0, 0, 0, _
                    3, 2, 3, 2, 3, 0, 0, 0, _
                    3, 0, 0, 0, _
                    3, 3, 3, 3, 0, 2, 3, 3, 0, 3, 3, 0, 3, 3, 0, 2, 3, 3, 0, 0, 0, _
                    3, 3, 3, 0, 3, 3, 3, 0, 3, 3, 3, 0, 3, 3, 3, 0, 3, 3, 3, 0, 3, 3, 3, 0, 2, 3, 3, 3, 0, 0, 0, _
                    3, 3, 3, 0, 3, 0, 3, 3, 0, 0, 0, _
                    3, 3, 3, 0, 0, 0, _
                    3, 3, 3, 3, 3, 3, 3, 0, 0, 0, _
                    3, 0, 0, 0, _
                    3)
                    
    'Display Project Form UserForm
    'Source: https://www.contextures.com/xlUserForm02.html
    form00_Nav.Show False
    
End Sub

