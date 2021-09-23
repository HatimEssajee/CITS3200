Attribute VB_Name = "Overview_Report"
Option Explicit

Sub Bring_Data()
    
    'PURPOSE: Determines dimensions of register table and loads first userform
    
    Dim Rpt As ListObject
    Dim err As Range
    Dim ReadRow As Variant
    Dim i As Integer, j As Integer, k As Integer
    Dim cRed As Long, cGreen As Long, cWhite As Long
    
    'Speed up
    'source: https://www.automateexcel.com/vba/turn-off-screen-updating/
    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayStatusBar = False
    Application.EnableEvents = False

    
    'Set color
    cRed = RGB(246, 176, 176)
    cGreen = RGB(146, 208, 80)
    cWhite = xlNone
    
    'Reference register table
    Set RegTable = ThisWorkbook.Sheets("Register").ListObjects("Register")
    Set Rpt = ThisWorkbook.Sheets("Overview Report").ListObjects("Report")
    Set err = ThisWorkbook.Sheets("Overview Report").Range("D1")
    
    j = 1
    
    'Reset error message
    err.Value = vbNullString
    
    'Clear report data
     If Not Rpt.DataBodyRange Is Nothing Then
        With Rpt.DataBodyRange
            .Interior.Color = cWhite
            .Font.Bold = False
            .Delete
        End With
    End If
    
    'Exit out if regiuster has no data
    If RegTable.DataBodyRange Is Nothing Then
        err.Value = "Register table has no data"
        Exit Sub
    End If
    
    'Exit out if all data is deleted
    
    
    'Apply filter to remove deleted
    ReadRow = RegTable.DataBodyRange
       
    For i = 1 To UBound(ReadRow)
        
        If ReadRow(i, 8) <> "DELETED" Then
            'Add row if greater than 1
            Rpt.ListRows.Add
                       
            'Copy data across
            With Rpt.DataBodyRange
                .Cells(j, 1) = ReadRow(i, 8)
                .Cells(j, 2) = Format(ReadRow(i, 2), "dd-mmm-yyyy")
                .Cells(j, 3) = Format(ReadRow(i, 6), "dd-mmm-yyyy hh:mm")
                
                'Study Details
                .Cells(j, 4) = ReadRow(i, 10)
                .Cells(j, 5) = ReadRow(i, 9)
                .Cells(j, 6) = ReadRow(i, 11)
                .Cells(j, 7) = ReadRow(i, 12)
                .Cells(j, 8) = ReadRow(i, 13)
                
                For k = 5 To 8
                    If .Cells(j, k).Value = vbNullString Then
                        Call AddFormat(.Cells(j, k), cRed)
                    End If
                Next k
                
                'CDA
                If ReadRow(i, 117) <> vbNullString Then
                    .Cells(j, 9).Value = "Date Recv. Sponsor = " & Format(ReadRow(i, 17), "dd-mmm-yy") & Chr(10) & _
                                         "Date Sent Contracts = " & Format(ReadRow(i, 18), "dd-mmm-yy") & Chr(10) & _
                                         "Date Recv. Contracts = " & Format(ReadRow(i, 19), "dd-mmm-yy") & Chr(10) & _
                                         "Date Sent Sponsor = " & Format(ReadRow(i, 20), "dd-mmm-yy") & Chr(10) & _
                                         "Date Finalised = " & Format(ReadRow(i, 21), "dd-mmm-yy") & Chr(10)
                                       
                    If ReadRow(i, 117) Then
                        Call AddFormat(.Cells(j, 9), cGreen)
                    Else
                        Call AddFormat(.Cells(j, 9), cRed)
                    End If
                    
                End If
                
                'FS
                If ReadRow(i, 118) <> vbNullString Then
                    .Cells(j, 10).Value = "Date Recv. = " & Format(ReadRow(i, 22), "dd-mmm-yy") & Chr(10) & _
                                         "Date Completed = " & Format(ReadRow(i, 23), "dd-mmm-yy") & "; " & ReadRow(i, 24)
                    If ReadRow(i, 118) Then
                        Call AddFormat(.Cells(j, 10), cGreen)
                    Else
                        Call AddFormat(.Cells(j, 10), cRed)
                    End If
                End If
                
                'Site Selection
                If ReadRow(i, 119) <> vbNullString Then
                    .Cells(j, 11).Value = "Pre-study visit = " & Format(ReadRow(i, 28), "dd-mmm-yy") & "; " & ReadRow(i, 29) & Chr(10) & _
                                         "Valid. visit = " & Format(ReadRow(i, 30), "dd-mmm-yy") & "; " & ReadRow(i, 31) & Chr(10) & _
                                         "Date Site Selected = " & Format(ReadRow(i, 32), "dd-mmm-yy")
                    If ReadRow(i, 119) Then
                        Call AddFormat(.Cells(j, 11), cGreen)
                    Else
                        Call AddFormat(.Cells(j, 11), cRed)
                    End If
                End If
                
                'Recruitment
                If ReadRow(i, 120) <> vbNullString Then
                    .Cells(j, 12).Value = "Plan. Meeting = " & Format(ReadRow(i, 36), "dd-mmm-yy") & Chr(10) & _
                                         "Status = " & ReadRow(i, 37)
                    If ReadRow(i, 120) Then
                        Call AddFormat(.Cells(j, 12), cGreen)
                    Else
                        Call AddFormat(.Cells(j, 12), cRed)
                    End If
                End If
                
            End With
            
            j = j + 1
        End If
    Next i
    
    'Speed up
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayStatusBar = True
    Application.EnableEvents = True

    
    'loop through each row
    'SOURCE: https://stackoverflow.com/questions/20097234/copying-rows-from-one-listobject-to-another-using-excel-vba
    'For i = 1 To RegTable.DataBodyRange.Rows.Count
End Sub

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

Private Sub AddFormat(rng As Range, Optional Col As Long)
    'PURPOSE: add borders as required
    
    With rng
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Interior.Color = Col
        .WrapText = True
        .HorizontalAlignment = xlHAlignLeft
        .VerticalAlignment = xlVAlignTop
        .Font.Bold = False
    End With
    
    
End Sub
