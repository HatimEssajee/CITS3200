Attribute VB_Name = "OpenForm"
Option Explicit

Sub OpenForm()
    
'    'Create combined team supporters list
'    Dim ws As Worksheet
'    Dim lGI As Long
'    Dim lTS As Long
'    Dim i As Integer
'    Dim j As Integer
'    Set ws = Sheets("Lookup Lists")
'
'    lGI = ws.Range("GI_Team").Rows.Count
'    lTS = ws.Range("TS_Team").Rows.Count
'
'    'Clear Team Supporters list
'    ws.Unprotect
'
'    ws.Range("K3").End(xlDown).ClearContents
'
'    'Copy ranges and alphabetize
'    ws.Range("K3").Resize(lGI, 1).Value = ws.Range("GI_Team").Value
'    ws.Range("K3").Offset(lGI, 0).Resize(lTS, 1).Value = ws.Range("TS_Team").Value
'
'    ws.Range("K3").Resize(lGI + lTS, 1).Sort key1:=ws.Range("K3"), order1:=xlAscending
'
'    ws.Protect
    
    'Display Project Form UserForm
    'Source: https://www.contextures.com/xlUserForm02.html
    form00_Nav.Show False
    
End Sub

