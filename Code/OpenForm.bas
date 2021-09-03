Attribute VB_Name = "OpenForm"
Option Explicit

Sub OpenForm()
    'PURPOSE: Determines dimensions of register table and loads first userform
        
'    'Find first and last used Row in register table
'    'Source: https://www.thespreadsheetguru.com/blog/2014/6/20/the-vba-guide-to-listobject-excel-tables
'    HeaderRow = Sheets("Register").Range("Register[[#Headers],[Study UID]]").Row
'    TopRow = HeaderRow + 1
'    BtmRow = RegTable.ListRows.Count
    
    'Reference register table
    Set RegTable = ThisWorkbook.Sheets("Register").ListObjects("Register")
    
    'Store current username in memory
    'Source: https://officetricks.com/excel-vba-get-username-windows-system/
    Username = ThisWorkbook.BuiltinDocumentProperties("Author")
    
    'Force default starting rowIndex for empty form and tickbox checked
    RowIndex = -1
    Tick = True
    
    'Display Project Form UserForm
    'Source: https://www.contextures.com/xlUserForm02.html
    form00_Nav.Show False
    
End Sub

