VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet_Report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    Dim RIndex As Long
    Dim Report As ListObject
    Dim i As Long
    
    Set Report = ListObjects("Report")
    
    'Clicking in table
    If Not Intersect(Target, Report.DataBodyRange) Is Nothing Then
        
        'Close all open user forms
        'SOURCE: https://stackoverflow.com/questions/48562941/closing-any-open-userform
        If UserForms.count <> 0 Then
            For i = VBA.UserForms.count - 1 To 0 Step -1
                Unload VBA.UserForms(i)
            Next
        End If
        
        RIndex = Selection.Row - 11
        RIndex = Report.DataBodyRange.Cells(RIndex, Report.ListColumns.count).Value
        
        'Redirect foxus away from selected cell
        ThisWorkbook.Sheets("Overview Report").Range("C1").Activate
        
        Overview_Report.OpenFromTable (RIndex)
        
    End If
    
End Sub

