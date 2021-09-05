Attribute VB_Name = "Overview_Report"
Option Explicit

'Sub Report_Builder()
'    Dim ws As Worksheet
'    Dim lGI As Long
'    Dim lTS As Long
'    Dim rGI As Range
'    Dim rTS As Range
'    Dim myTable As ListObject
'    Dim xCell As Range
'    Dim items(1 To 3) As Variant
'    Dim i As Integer
'    Dim fColor As Long
'    Dim sColor As Long
'
'    'Turn of update and calcs to speed up macro
'    Application.EnableEvents = False
'    Application.ScreenUpdating = False
'    Application.DisplayAlerts = False
'    Application.Calculation = xlCalculationManual
'
'    Set ws = Sheets("Work Alloc Report")
'    Set myTable = Sheets("Register").ListObjects("Register")
'    Set rGI = Sheets("Lookup Lists").Range("GI_Team")
'    Set rTS = Sheets("Lookup Lists").Range("TS_Team")
'
'    'Set color index for bands
'    fColor = xlNone
'    sColor = RGB(217, 217, 217)
'
'
'    'Clear sheet
'    ws.Range("A:G").Clear
'
'    'Add headers
'    With ws.Range("B4")
'        .Value = "Do not use columns A:G as they are cleared regularly"
'        .Font.Bold = True
'        .HorizontalAlignment = xlLeft
'        .VerticalAlignment = xlCenter
'        .Font.Color = vbRed
'    End With
'
'    With ws.Range("B6")
'        .Value = "Team Member"
'        .Font.Bold = True
'        .HorizontalAlignment = xlCenter
'        .VerticalAlignment = xlCenter
'    End With
'
'    With ws.Range("C6")
'        .Value = "Champion"
'        .Font.Bold = True
'        .HorizontalAlignment = xlCenter
'        .VerticalAlignment = xlCenter
'    End With
'
'    With ws.Range("D6")
'        .Value = "Supporter"
'        .Font.Bold = True
'        .HorizontalAlignment = xlCenter
'        .VerticalAlignment = xlCenter
'    End With
'
'    With ws.Range("C5:D5")
'        .Merge
'        .Value = "No. Ongoing Projects"
'        .Font.Bold = True
'        .HorizontalAlignment = xlCenter
'        .VerticalAlignment = xlCenter
'    End With
'
'    With ws.Range("E5:E6")
'        .Merge
'        .Value = "Overdue Ongoing Project"
'        .Font.Bold = True
'        .WrapText = True
'        .HorizontalAlignment = xlCenter
'        .VerticalAlignment = xlCenter
'    End With
'
'    'Determine length of name lists
'    lGI = Sheets("Lookup Lists").Range("GI_Team").Rows.Count
'    lTS = Sheets("Lookup Lists").Range("TS_Team").Rows.Count
'
'    'Add first title row
'    With ws.Range("B8")
'        .Value = "Technical & Support"
'        .Font.Color = vbBlue
'    End With
'
'    'Add border around title
'    ws.Range("B8").Resize(1, 4).BorderAround Weight:=xlThick
'
'    'Add named range list for team personnel
'    ws.Range("B9").Resize(lTS, 1).Value = rTS.Value
'
'    Call ColorBarMe(ws.Range("B9").Resize(lTS, 4), fColor, sColor)
'
'    'Run countifs on register table to get number of ongoing projects
'    For Each xCell In ws.Range("B9").Resize(lTS, 1)
'
'        items(1) = Application.WorksheetFunction.CountIfs(myTable.DataBodyRange.Columns(24), xCell.Value, myTable.DataBodyRange.Columns(4), "<>Complete")
'        items(2) = Application.WorksheetFunction.CountIfs(myTable.DataBodyRange.Columns(25), "*" & xCell.Value & "*", myTable.DataBodyRange.Columns(4), "<>Complete")
'        items(3) = Application.WorksheetFunction.CountIfs(myTable.DataBodyRange.Columns(24), "*" & xCell.Value & "*", myTable.DataBodyRange.Columns(4), "<>Complete", _
'                                    myTable.DataBodyRange.Columns(27), "Yes")
'
'        'Replace zeros with "" for cleaner look
'        For i = 1 To 3
'            If items(i) = 0 Then items(i) = ""
'        Next i
'
'        xCell.Offset(0, 1).Resize(1, 3).Value = items
'
'    Next xCell
'
'    'Tally total for team
'    items(1) = Application.WorksheetFunction.Sum(ws.Range("B9").Offset(0, 1).Resize(lTS, 1))
'    items(2) = Application.WorksheetFunction.Sum(ws.Range("B9").Offset(0, 2).Resize(lTS, 1))
'    items(3) = Application.WorksheetFunction.Sum(ws.Range("B9").Offset(0, 3).Resize(lTS, 1))
'
'    With ws.Range("B8").Offset(0, 1).Resize(1, 3)
'        .Value = items
'        .Font.Bold = True
'        .Font.Color = vbBlue
'    End With
'
'    'Add second title two rows below last name
'    With ws.Range("B9").Offset(lTS + 2, 0)
'        .Value = "Governance & Improvement"
'        .Font.Color = vbBlue
'    End With
'
'    'Add border around title
'    ws.Range("B9").Offset(lTS + 2, 0).Resize(1, 4).BorderAround Weight:=xlThick
'
'
'    'Add second named range list for team personnel
'    ws.Range("B9").Offset(lTS + 3, 0).Resize(lGI, 1).Value = rGI.Value
'
'    Call ColorBarMe(ws.Range("B9").Offset(lTS + 3, 0).Resize(lGI, 4), fColor, sColor)
'
'    For Each xCell In ws.Range("B9").Offset(lTS + 3, 0).Resize(lGI, 1)
'        items(1) = Application.WorksheetFunction.CountIfs(myTable.DataBodyRange.Columns(24), xCell.Value, myTable.DataBodyRange.Columns(4), "<>Complete")
'        items(2) = Application.WorksheetFunction.CountIfs(myTable.DataBodyRange.Columns(25), "*" & xCell.Value & "*", myTable.DataBodyRange.Columns(4), "<>Complete")
'        items(3) = Application.WorksheetFunction.CountIfs(myTable.DataBodyRange.Columns(24), "*" & xCell.Value & "*", myTable.DataBodyRange.Columns(4), "<>Complete", _
'                                    myTable.DataBodyRange.Columns(27), "Yes")
'
'        'Replace zeros with "" for cleaner look
'        For i = 1 To 3
'            If items(i) = 0 Then items(i) = ""
'        Next i
'
'        xCell.Offset(0, 1).Resize(1, 3).Value = items
'
'    Next xCell
'
'    'Tally total for team
'    items(1) = Application.WorksheetFunction.Sum(ws.Range("B9").Offset(lTS + 3, 1).Resize(lTS, 1))
'    items(2) = Application.WorksheetFunction.Sum(ws.Range("B9").Offset(lTS + 3, 1).Resize(lTS, 1))
'    items(3) = Application.WorksheetFunction.Sum(ws.Range("B9").Offset(lTS + 3, 1).Resize(lTS, 1))
'
'    With ws.Range("B9").Offset(lTS + 2, 1).Resize(1, 3)
'        .Value = items
'        .Font.Bold = True
'        .Font.Color = vbBlue
'    End With
'
'    'Refresh pivot
'    ActiveWorkbook.RefreshAll
'
'    'revert to default settings
'    Application.EnableEvents = True
'    Application.ScreenUpdating = True
'    Application.DisplayAlerts = True
'    Application.Calculation = xlCalculationAutomatic
'
'
'
'End Sub
'
'Sub ColorBarMe(rng As Range, firstColor As Long, secondColor As Long)
'    'Add colour banding
'    'Source: https://stackoverflow.com/questions/4629100/alternate-row-colors-in-range
'    Dim FirstRow As Long
'    Dim nCols As Long
'    Dim nRows As Long
'    Dim xCell As Range
'
'    FirstRow = rng.item(1).Row
'    nCols = rng.Columns.Count
'    nRows = rng.Rows.Count
'
'    Set rng = rng.Resize(nRows, 1)
'
'    For Each xCell In rng
'        If (xCell.Row - FirstRow) Mod 2 = 0 Then
'            xCell.Resize(1, nCols).Interior.Color = firstColor
'        Else
'            xCell.Resize(1, nCols).Interior.Color = secondColor
'        End If
'    Next xCell
'
'End Sub
