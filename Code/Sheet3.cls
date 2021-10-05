VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    Dim RIndex As Long
    Dim Report As ListObject
    
    Set Report = ListObjects("Report")
    
    'Clicking in table
    If Not Intersect(Target, Report.DataBodyRange) Is Nothing Then
        
        RIndex = Selection.Row - 10
        RIndex = Report.DataBodyRange.Cells(RIndex, 35).Value
        
        'Redirect foxus away from selected cell
        ThisWorkbook.Sheets("Overview Report").Range("C1").Activate
        
        Overview_Report.OpenFromTable (RIndex)
        
    End If
    
End Sub

