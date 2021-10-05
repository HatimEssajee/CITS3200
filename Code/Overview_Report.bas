Attribute VB_Name = "Overview_Report"
Option Explicit

Sub Bring_Data()
    
    'PURPOSE: Brings in all all data from register table in a summarized manner
    
    Dim Rpt As ListObject
    Dim err As Range, Header As Range
    Dim ReadRow As Variant
    Dim i As Integer, j As Integer, k As Integer
    Dim cRed As Long, cGreen As Long, cWhite As Long
    Dim tempStr As String
    
    'Turn off Settings to speed up
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
    Set err = ThisWorkbook.Sheets("Overview Report").Range("C1")
    Set Header = RegTable.HeaderRowRange
    j = 1
    
    'Reset error message
    With err
        .Value = vbNullString
        .Font.Color = vbRed
        .Font.Size = 11
    End With
    
    'Clear report data
     If Not Rpt.DataBodyRange Is Nothing Then
        With Rpt.DataBodyRange
            .Interior.Color = cWhite
            .Font.Bold = False
            .Borders.LineStyle = xlNone
            .Delete
        End With
    End If
    
    'Exit out if register has no data
    If RegTable.DataBodyRange Is Nothing Then
        err.Value = "Register table has no data"
        Exit Sub
    End If
    
    'Exit out if all data is deleted
    If Application.WorksheetFunction.CountIf(RegTable.DataBodyRange.Columns(7), "DELETED") = RegTable.DataBodyRange.Rows.count Then
        err.Value = "Register table only has deleted values"
        Exit Sub
    End If
        
    'Apply filter to remove deleted
    ReadRow = RegTable.DataBodyRange
       
    For i = 1 To UBound(ReadRow)
        
        If ReadRow(i, 7) <> "DELETED" Then
            'Add row
            Rpt.ListRows.Add
                       
            'Copy data across
            With Rpt.DataBodyRange
                
                'Study Details
                .Cells(j, 1) = ReadRow(i, 7)
                .Cells(j, 2) = Format(ReadRow(i, 1), "dd-mmm-yyyy")
                .Cells(j, 3) = ReadRow(i, 9)
                .Cells(j, 4) = Format(ReadRow(i, 5), "dd-mmm-yyyy hh:mm")
                .Cells(j, 5) = ReadRow(i, 8)
                .Cells(j, 6) = ReadRow(i, 10)
                .Cells(j, 7) = ReadRow(i, 11)
                .Cells(j, 8) = ReadRow(i, 12)
                .Cells(j, 35) = i
                
                For k = 5 To 8
                    If .Cells(j, k).Value = vbNullString Then
                        Call AddFormat(.Cells(j, k), cRed)
                    End If
                Next k
                
                'Study details colour
                If ReadRow(i, 156) Then
                        Call AddFormat(.Cells(j, 1), cGreen)
                    Else
                        Call AddFormat(.Cells(j, 1), cRed)
                End If
                
                'Add reminder in age column
                If ReadRow(i, 13) <> vbNullString And Not (ReadRow(i, 129)) Then
                        .Cells(j, 8).Value = .Cells(j, 8).Value & Chr(10) & Chr(10) & "Reminder:" & Chr(10) & _
                                                ReadRow(i, 13)
                End If
                
                'CDA
                If ReadRow(i, 130) <> vbNullString Then
                    .Cells(j, 9).Value = "Date Recv. Sponsor = " & Format(ReadRow(i, 16), "dd-mmm-yy") & Chr(10) & _
                                         "Date Sent Contracts = " & Format(ReadRow(i, 17), "dd-mmm-yy") & Chr(10) & _
                                         "Date Recv. Contracts = " & Format(ReadRow(i, 18), "dd-mmm-yy") & Chr(10) & _
                                         "Date Sent Sponsor = " & Format(ReadRow(i, 19), "dd-mmm-yy") & Chr(10) & _
                                         "Date Finalised = " & Format(ReadRow(i, 20), "dd-mmm-yy") & Chr(10)
                    
                    If ReadRow(i, 130) Then
                        Call AddFormat(.Cells(j, 9), cGreen)
                    Else
                        Call AddFormat(.Cells(j, 9), cRed)
                        
                        'Add reminder if incomplete
                        If ReadRow(i, 21) <> vbNullString Then
                            .Cells(j, 9).Value = .Cells(j, 9).Value & Chr(10) & Chr(10) & "Reminder:" & Chr(10) & _
                                                ReadRow(i, 21)
                        End If
                        
                    End If
                    
                End If
                
                'Feasibility
                If ReadRow(i, 131) <> vbNullString Then
                    .Cells(j, 10).Value = "Date Recv. = " & Format(ReadRow(i, 24), "dd-mmm-yy") & Chr(10) & _
                                         "Date Completed = " & Format(ReadRow(i, 25), "dd-mmm-yy") & "; " & ReadRow(i, 26)
                
                    If ReadRow(i, 131) Then
                        Call AddFormat(.Cells(j, 10), cGreen)
                    Else
                        Call AddFormat(.Cells(j, 10), cRed)
                        
                        'Add reminder if incomplete
                        If ReadRow(i, 27) <> vbNullString Then
                            .Cells(j, 10).Value = .Cells(j, 10).Value & Chr(10) & Chr(10) & "Reminder:" & Chr(10) & _
                                                ReadRow(i, 27)
                        End If
                        
                    End If
                End If
                
                'Site Selection
                If ReadRow(i, 132) <> vbNullString Then
                    .Cells(j, 11).Value = "Pre-study visit = " & Format(ReadRow(i, 30), "dd-mmm-yy") & "; " & ReadRow(i, 31) & Chr(10) & _
                                         "Valid. visit = " & Format(ReadRow(i, 32), "dd-mmm-yy") & "; " & ReadRow(i, 33) & Chr(10) & _
                                         "Date Site Selected = " & Format(ReadRow(i, 34), "dd-mmm-yy")
                    If ReadRow(i, 132) Then
                        Call AddFormat(.Cells(j, 11), cGreen)
                    Else
                        Call AddFormat(.Cells(j, 11), cRed)
                        
                        'Add reminder if incomplete
                        If ReadRow(i, 35) <> vbNullString Then
                            .Cells(j, 11).Value = .Cells(j, 11).Value & Chr(10) & Chr(10) & "Reminder:" & Chr(10) & _
                                                ReadRow(i, 35)
                        End If
                        
                    End If
                End If
                
                'Recruitment
                If ReadRow(i, 133) <> vbNullString Then
                    .Cells(j, 12).Value = "Plan. Meeting = " & Format(ReadRow(i, 38), "dd-mmm-yy")
                    
                    If ReadRow(i, 133) Then
                        Call AddFormat(.Cells(j, 12), cGreen)
                    Else
                        Call AddFormat(.Cells(j, 12), cRed)
                        
                        'Add reminder if incomplete
                        If ReadRow(i, 39) <> vbNullString Then
                            .Cells(j, 12).Value = .Cells(j, 12).Value & Chr(10) & Chr(10) & "Reminder:" & Chr(10) & _
                                                ReadRow(i, 39)
                        End If
                        
                    End If
                End If
                
                'Overall Ethics
                If ReadRow(i, 153) = True Then
                    tempStr = ""
                    For k = 134 To 138
                        If ReadRow(i, k) Then
                            If tempStr <> "" Then
                                tempStr = tempStr + Chr(10)
                            End If
                            
                            tempStr = tempStr + Mid(Header.Cells(1, k), 10, Len(Header.Cells(1, k)))
                        End If
                    Next k
                    
                    .Cells(j, 13) = tempStr
                    Call AddFormat(.Cells(j, 13), cGreen)
                    
                ElseIf ReadRow(i, 153) <> vbNullString Then
                    tempStr = ""
                    For k = 134 To 138
                        If ReadRow(i, k) <> vbNullString And Not ReadRow(i, k) Then
                            If tempStr <> "" Then
                                tempStr = tempStr + Chr(10)
                            End If
                            
                            tempStr = tempStr + Mid(Header.Cells(1, k), 10, Len(Header.Cells(1, k)) - 18) + " Incomplete"
                        End If
                    Next k
                    .Cells(j, 13) = tempStr
                    Call AddFormat(.Cells(j, 13), cRed)
                End If
                
                
                'CAHS Ethics
                If ReadRow(i, 134) <> vbNullString Then
                    .Cells(j, 14).Value = "Date Submitted = " & Format(ReadRow(i, 42), "dd-mmm-yy") & Chr(10) & _
                                         "Date Responded = " & Format(ReadRow(i, 43), "dd-mmm-yy") & Chr(10) & _
                                         "Date Resubmitted = " & Format(ReadRow(i, 44), "dd-mmm-yy") & Chr(10) & _
                                         "Date Approved = " & Format(ReadRow(i, 45), "dd-mmm-yy")
                    
                    If ReadRow(i, 134) Then
                        Call AddFormat(.Cells(j, 14), cGreen)
                    Else
                        Call AddFormat(.Cells(j, 14), cRed)
                        
                        'Add reminder if incomplete
                        If ReadRow(i, 46) <> vbNullString Then
                            .Cells(j, 14).Value = .Cells(j, 14).Value & Chr(10) & Chr(10) & "Reminder:" & Chr(10) & _
                                                ReadRow(i, 46)
                        End If
                        
                    End If
                End If
                
                'NMA Ethics
                If ReadRow(i, 135) <> vbNullString Then
                    .Cells(j, 15).Value = "Ethics Commitee = " & ReadRow(i, 47) & Chr(10) & _
                                         "Date Submitted = " & Format(ReadRow(i, 48), "dd-mmm-yy") & Chr(10) & _
                                         "Date Approved = " & Format(ReadRow(i, 49), "dd-mmm-yy")
                    If ReadRow(i, 135) Then
                        Call AddFormat(.Cells(j, 15), cGreen)
                    Else
                        Call AddFormat(.Cells(j, 15), cRed)
                        
                        'Add reminder if incomplete
                        If ReadRow(i, 50) <> vbNullString Then
                            .Cells(j, 15).Value = .Cells(j, 15).Value & Chr(10) & Chr(10) & "Reminder:" & Chr(10) & _
                                                ReadRow(i, 50)
                        End If
                        
                    End If
                End If
                
                'WNHS Ethics
                If ReadRow(i, 136) <> vbNullString Then
                    .Cells(j, 16).Value = "Date Submitted = " & Format(ReadRow(i, 51), "dd-mmm-yy") & Chr(10) & _
                                         "Date Approved = " & Format(ReadRow(i, 52), "dd-mmm-yy")
                    If ReadRow(i, 136) Then
                        Call AddFormat(.Cells(j, 16), cGreen)
                    Else
                        Call AddFormat(.Cells(j, 16), cRed)
                        
                        'Add reminder if incomplete
                        If ReadRow(i, 53) <> vbNullString Then
                            .Cells(j, 16).Value = .Cells(j, 16).Value & Chr(10) & Chr(10) & "Reminder:" & Chr(10) & _
                                                ReadRow(i, 53)
                        End If
                        
                    End If
                End If
                
                'SJOG Ethics
                If ReadRow(i, 137) <> vbNullString Then
                    .Cells(j, 17).Value = "Date Submitted = " & Format(ReadRow(i, 54), "dd-mmm-yy") & Chr(10) & _
                                         "Date Approved = " & Format(ReadRow(i, 55), "dd-mmm-yy")
                    If ReadRow(i, 137) Then
                        Call AddFormat(.Cells(j, 17), cGreen)
                    Else
                        Call AddFormat(.Cells(j, 17), cRed)
                        
                        'Add reminder if incomplete
                        If ReadRow(i, 56) <> vbNullString Then
                            .Cells(j, 17).Value = .Cells(j, 17).Value & Chr(10) & Chr(10) & "Reminder:" & Chr(10) & _
                                                ReadRow(i, 56)
                        End If
                        
                    End If
                End If
                
                'Others Ethics
                If ReadRow(i, 138) <> vbNullString Then
                    .Cells(j, 18).Value = "Ethics Commitee = " & ReadRow(i, 57) & Chr(10) & _
                                         "Date Submitted = " & Format(ReadRow(i, 58), "dd-mmm-yy") & Chr(10) & _
                                         "Date Approved = " & Format(ReadRow(i, 59), "dd-mmm-yy")
                    If ReadRow(i, 138) Then
                        Call AddFormat(.Cells(j, 18), cGreen)
                    Else
                        Call AddFormat(.Cells(j, 18), cRed)
                        
                        'Add reminder if incomplete
                        If ReadRow(i, 60) <> vbNullString Then
                            .Cells(j, 18).Value = .Cells(j, 18).Value & Chr(10) & Chr(10) & "Reminder:" & Chr(10) & _
                                                ReadRow(i, 60)
                        End If
                        
                    End If
                End If
                
                'Overall Governance
                If ReadRow(i, 154) = True Then
                    tempStr = ""
                    For k = 139 To 145
                        If ReadRow(i, k) Then
                            If tempStr <> "" Then
                                tempStr = tempStr + Chr(10)
                            End If
                            
                            tempStr = tempStr + Mid(Header.Cells(1, k), 7, Len(Header.Cells(1, k)))
                        End If
                    Next k
                    
                    .Cells(j, 19) = tempStr
                    Call AddFormat(.Cells(j, 19), cGreen)
                
                ElseIf ReadRow(i, 154) <> vbNullString Then
                    tempStr = ""
                    For k = 139 To 145
                        If ReadRow(i, k) <> vbNullString And Not ReadRow(i, k) Then
                            If tempStr <> "" Then
                                tempStr = tempStr + Chr(10)
                            End If
                            
                            tempStr = tempStr + Mid(Header.Cells(1, k), 7, Len(Header.Cells(1, k)) - 15) + " Incomplete"
                        End If
                    Next k
                    .Cells(j, 19) = tempStr
                    Call AddFormat(.Cells(j, 19), cRed)
                End If
                
                'PCH Governance
                If ReadRow(i, 139) <> vbNullString Then
                    .Cells(j, 20).Value = "Date Submitted = " & Format(ReadRow(i, 63), "dd-mmm-yy") & Chr(10) & _
                                         "Date Responded = " & Format(ReadRow(i, 64), "dd-mmm-yy") & Chr(10) & _
                                         "Date Approved = " & Format(ReadRow(i, 65), "dd-mmm-yy")
                    If ReadRow(i, 139) Then
                        Call AddFormat(.Cells(j, 20), cGreen)
                    Else
                        Call AddFormat(.Cells(j, 20), cRed)
                        
                        'Add reminder if incomplete
                        If ReadRow(i, 66) <> vbNullString Then
                            .Cells(j, 20).Value = .Cells(j, 20).Value & Chr(10) & Chr(10) & "Reminder:" & Chr(10) & _
                                                ReadRow(i, 66)
                        End If
                        
                    End If
                End If
                
                'TKI Governance
                If ReadRow(i, 140) <> vbNullString Then
                    .Cells(j, 21).Value = "Date Submitted = " & Format(ReadRow(i, 67), "dd-mmm-yy") & Chr(10) & _
                                         "Date Responded = " & Format(ReadRow(i, 68), "dd-mmm-yy") & Chr(10) & _
                                         "Date Approved = " & Format(ReadRow(i, 69), "dd-mmm-yy")
                    If ReadRow(i, 140) Then
                        Call AddFormat(.Cells(j, 21), cGreen)
                    Else
                        Call AddFormat(.Cells(j, 21), cRed)
                        
                        'Add reminder if incomplete
                        If ReadRow(i, 70) <> vbNullString Then
                            .Cells(j, 21).Value = .Cells(j, 21).Value & Chr(10) & Chr(10) & "Reminder:" & Chr(10) & _
                                                ReadRow(i, 70)
                        End If
                        
                    End If
                End If
                
                'KEMH Governance
                If ReadRow(i, 141) <> vbNullString Then
                    .Cells(j, 22).Value = "Date Submitted = " & Format(ReadRow(i, 71), "dd-mmm-yy") & Chr(10) & _
                                         "Date Responded = " & Format(ReadRow(i, 72), "dd-mmm-yy") & Chr(10) & _
                                         "Date Approved = " & Format(ReadRow(i, 73), "dd-mmm-yy")
                    If ReadRow(i, 141) Then
                        Call AddFormat(.Cells(j, 22), cGreen)
                    Else
                        Call AddFormat(.Cells(j, 22), cRed)
                        
                        'Add reminder if incomplete
                        If ReadRow(i, 74) <> vbNullString Then
                            .Cells(j, 22).Value = .Cells(j, 22).Value & Chr(10) & Chr(10) & "Reminder:" & Chr(10) & _
                                                ReadRow(i, 74)
                        End If
                        
                    End If
                End If
                
                'SJOG Subiaco Governance
                If ReadRow(i, 142) <> vbNullString Then
                    .Cells(j, 23).Value = "Date Submitted = " & Format(ReadRow(i, 75), "dd-mmm-yy") & Chr(10) & _
                                         "Date Responded = " & Format(ReadRow(i, 76), "dd-mmm-yy") & Chr(10) & _
                                         "Date Approved = " & Format(ReadRow(i, 77), "dd-mmm-yy")
                    If ReadRow(i, 142) Then
                        Call AddFormat(.Cells(j, 23), cGreen)
                    Else
                        Call AddFormat(.Cells(j, 23), cRed)
                        
                        'Add reminder if incomplete
                        If ReadRow(i, 78) <> vbNullString Then
                            .Cells(j, 23).Value = .Cells(j, 23).Value & Chr(10) & Chr(10) & "Reminder:" & Chr(10) & _
                                                ReadRow(i, 78)
                        End If
                        
                    End If
                End If
                
                'SJOG Mt Lawley Governance
                If ReadRow(i, 143) <> vbNullString Then
                    .Cells(j, 24).Value = "Date Submitted = " & Format(ReadRow(i, 79), "dd-mmm-yy") & Chr(10) & _
                                         "Date Responded = " & Format(ReadRow(i, 80), "dd-mmm-yy") & Chr(10) & _
                                         "Date Approved = " & Format(ReadRow(i, 81), "dd-mmm-yy")
                    If ReadRow(i, 143) Then
                        Call AddFormat(.Cells(j, 24), cGreen)
                    Else
                        Call AddFormat(.Cells(j, 24), cRed)
                        
                        'Add reminder if incomplete
                        If ReadRow(i, 82) <> vbNullString Then
                            .Cells(j, 24).Value = .Cells(j, 24).Value & Chr(10) & Chr(10) & "Reminder:" & Chr(10) & _
                                                ReadRow(i, 82)
                        End If
                        
                    End If
                End If
                
                'SJOG Murdoch Governance
                If ReadRow(i, 144) <> vbNullString Then
                    .Cells(j, 25).Value = "Date Submitted = " & Format(ReadRow(i, 83), "dd-mmm-yy") & Chr(10) & _
                                         "Date Responded = " & Format(ReadRow(i, 84), "dd-mmm-yy") & Chr(10) & _
                                         "Date Approved = " & Format(ReadRow(i, 85), "dd-mmm-yy")
                    If ReadRow(i, 144) Then
                        Call AddFormat(.Cells(j, 25), cGreen)
                    Else
                        Call AddFormat(.Cells(j, 25), cRed)
                        
                        'Add reminder if incomplete
                        If ReadRow(i, 86) <> vbNullString Then
                            .Cells(j, 25).Value = .Cells(j, 25).Value & Chr(10) & Chr(10) & "Reminder:" & Chr(10) & _
                                                ReadRow(i, 86)
                        End If
                        
                    End If
                End If
                
                'Others Governance
                If ReadRow(i, 145) <> vbNullString Then
                    .Cells(j, 26).Value = "Gov. Body = " & Format(ReadRow(i, 87), "dd-mmm-yy") & Chr(10) & _
                                         "Date Submitted = " & Format(ReadRow(i, 88), "dd-mmm-yy") & Chr(10) & _
                                         "Date Responded = " & Format(ReadRow(i, 89), "dd-mmm-yy") & Chr(10) & _
                                         "Date Approved = " & Format(ReadRow(i, 90), "dd-mmm-yy")
                    If ReadRow(i, 145) Then
                        Call AddFormat(.Cells(j, 26), cGreen)
                    Else
                        Call AddFormat(.Cells(j, 26), cRed)
                        
                        'Add reminder if incomplete
                        If ReadRow(i, 91) <> vbNullString Then
                            .Cells(j, 26).Value = .Cells(j, 26).Value & Chr(10) & Chr(10) & "Reminder:" & Chr(10) & _
                                                ReadRow(i, 91)
                        End If
                        
                    End If
                End If
                
                  
                'Indemnity
                If ReadRow(i, 149) <> vbNullString Then
                    .Cells(j, 27).Value = "Date Received = " & Format(ReadRow(i, 105), "dd-mmm-yy") & Chr(10) & _
                                         "Date Sent Contracts = " & Format(ReadRow(i, 106), "dd-mmm-yy") & Chr(10) & _
                                         "Date Completed = " & Format(ReadRow(i, 107), "dd-mmm-yy")
                    If ReadRow(i, 149) Then
                        Call AddFormat(.Cells(j, 27), cGreen)
                    Else
                        Call AddFormat(.Cells(j, 27), cRed)
                        
                        'Add reminder if incomplete
                        If ReadRow(i, 108) <> vbNullString Then
                            .Cells(j, 27).Value = .Cells(j, 27).Value & Chr(10) & Chr(10) & "Reminder:" & Chr(10) & _
                                                ReadRow(i, 108)
                        End If
                        
                    End If
                End If
                
                'Overall Budget
                If ReadRow(i, 155) = True Then
                    tempStr = ""
                    For k = 146 To 148
                        If ReadRow(i, k) Then
                            If tempStr <> "" Then
                                tempStr = tempStr + Chr(10)
                            End If
                            
                            tempStr = tempStr + Mid(Header.Cells(1, k), 10, Len(Header.Cells(1, k)))
                        End If
                    Next k
                    
                    .Cells(j, 28) = tempStr
                    Call AddFormat(.Cells(j, 28), cGreen)
                    
                ElseIf ReadRow(i, 155) <> vbNullString Then
                    tempStr = ""
                    For k = 146 To 148
                        If ReadRow(i, k) <> vbNullString And Not ReadRow(i, k) Then
                            If tempStr <> "" Then
                                tempStr = tempStr + Chr(10)
                            End If
                            
                            tempStr = tempStr + Mid(Header.Cells(1, k), 10, Len(Header.Cells(1, k)) - 18) + " Incomplete"
                        End If
                    Next k
                    .Cells(j, 28) = tempStr
                    Call AddFormat(.Cells(j, 28), cRed)
                End If
                
                'VTG Budget
                If ReadRow(i, 146) <> vbNullString Then
                    .Cells(j, 29).Value = "Date Finalised = " & Format(ReadRow(i, 94), "dd-mmm-yy") & Chr(10) & _
                                         "Date Submitted Finance = " & Format(ReadRow(i, 95), "dd-mmm-yy") & Chr(10) & _
                                         "Date Approved = " & Format(ReadRow(i, 96), "dd-mmm-yy")
                    If ReadRow(i, 146) Then
                        Call AddFormat(.Cells(j, 29), cGreen)
                    Else
                        Call AddFormat(.Cells(j, 29), cRed)
                        
                        'Add reminder if incomplete
                        If ReadRow(i, 97) <> vbNullString Then
                            .Cells(j, 29).Value = .Cells(j, 29).Value & Chr(10) & Chr(10) & "Reminder:" & Chr(10) & _
                                                ReadRow(i, 97)
                        End If
                        
                    End If
                End If
                
                'TKI Budget
                If ReadRow(i, 147) <> vbNullString Then
                    .Cells(j, 30).Value = "Date Approved = " & Format(ReadRow(i, 98), "dd-mmm-yy")
                    
                    If ReadRow(i, 147) Then
                        Call AddFormat(.Cells(j, 30), cGreen)
                    Else
                        Call AddFormat(.Cells(j, 30), cRed)
                        
                        'Add reminder if incomplete
                        If ReadRow(i, 99) <> vbNullString Then
                            .Cells(j, 30).Value = .Cells(j, 30).Value & Chr(10) & Chr(10) & "Reminder:" & Chr(10) & _
                                                ReadRow(i, 99)
                        End If
                        
                    End If
                End If
                
                'Pharmacy Budget
                If ReadRow(i, 148) <> vbNullString Then
                    .Cells(j, 31).Value = "Date Quote Recv. = " & Format(ReadRow(i, 100), "dd-mmm-yy") & Chr(10) & _
                                         "Date PO Finalised = " & Format(ReadRow(i, 101), "dd-mmm-yy")
                    If ReadRow(i, 148) Then
                        Call AddFormat(.Cells(j, 31), cGreen)
                    Else
                        Call AddFormat(.Cells(j, 31), cRed)
                        
                        'Add reminder if incomplete
                        If ReadRow(i, 102) <> vbNullString Then
                            .Cells(j, 31).Value = .Cells(j, 31).Value & Chr(10) & Chr(10) & "Reminder:" & Chr(10) & _
                                                ReadRow(i, 102)
                        End If
                        
                    End If
                End If
                
                'CTRA
                If ReadRow(i, 150) <> vbNullString Then
                    .Cells(j, 32).Value = "Date Submitted R,G&C = " & Format(ReadRow(i, 111), "dd-mmm-yy") & Chr(10) & _
                                         "Date UWA Review = " & Format(ReadRow(i, 112), "dd-mmm-yy") & Chr(10) & _
                                         "Date Finance Review = " & Format(ReadRow(i, 113), "dd-mmm-yy") & Chr(10) & _
                                         "Date COO/CFO Sign-off = " & Format(ReadRow(i, 114), "dd-mmm-yy") & Chr(10) & _
                                         "Date VTG Sign-off = " & Format(ReadRow(i, 115), "dd-mmm-yy") & Chr(10) & _
                                         "Date Subm. Company = " & Format(ReadRow(i, 116), "dd-mmm-yy") & Chr(10) & _
                                         "Date Finalised = " & Format(ReadRow(i, 117), "dd-mmm-yy")
                    If ReadRow(i, 150) Then
                        Call AddFormat(.Cells(j, 32), cGreen)
                    Else
                        Call AddFormat(.Cells(j, 32), cRed)
                        
                        'Add reminder if incomplete
                        If ReadRow(i, 118) <> vbNullString Then
                            .Cells(j, 32).Value = .Cells(j, 32).Value & Chr(10) & Chr(10) & "Reminder:" & Chr(10) & _
                                                ReadRow(i, 118)
                        End If
                        
                    End If
                End If
                
                'Financial Disclosure
                If ReadRow(i, 151) <> vbNullString Then
                    .Cells(j, 33).Value = "Date Completed = " & Format(ReadRow(i, 121), "dd-mmm-yy")
                    If ReadRow(i, 151) Then
                        Call AddFormat(.Cells(j, 33), cGreen)
                    Else
                        Call AddFormat(.Cells(j, 33), cRed)
                        
                        'Add reminder if incomplete
                        If ReadRow(i, 122) <> vbNullString Then
                            .Cells(j, 33).Value = .Cells(j, 33).Value & Chr(10) & Chr(10) & "Reminder:" & Chr(10) & _
                                                ReadRow(i, 122)
                        End If
                        
                    End If
                End If
                
                'SIV
                If ReadRow(i, 152) <> vbNullString Then
                    .Cells(j, 34).Value = "Date of SIV = " & Format(ReadRow(i, 125), "dd-mmm-yy")
                    If ReadRow(i, 152) Then
                        Call AddFormat(.Cells(j, 34), cGreen)
                    Else
                        Call AddFormat(.Cells(j, 34), cRed)
                        
                        'Add reminder if incomplete
                        If ReadRow(i, 126) <> vbNullString Then
                            .Cells(j, 34).Value = .Cells(j, 34).Value & Chr(10) & Chr(10) & "Reminder:" & Chr(10) & _
                                                ReadRow(i, 126)
                        End If
                        
                    End If
                End If
                
            End With
            
            j = j + 1
        End If
    Next i
    
    'Reinstate Settings
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayStatusBar = True
    Application.EnableEvents = True
    
    MsgBox ("Data Succesfully Imported")
    err.Font.Color = vbBlack
    err.Value = "Data retrieved " & Format(Now, "dd-mmm-yyyy hh:mm AM/PM")
End Sub

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

Private Function ArrayCountIf(arr As Variant, RowIndex As Integer, lCol As Long, uCol As Long, Value As Variant) As Long
    'PURPOSE: run a countif over array between two column references - inclusive
    Dim i As Long
    Dim count As Long
    
    count = 0
    
    For i = lCol To uCol
        If arr(RowIndex, i) = Value Then
            count = count + 1
        End If
    Next i
    
    ArrayCountIf = count
End Function


Sub OpenFromTable(RIndex As Long)
    'PURPOSE: Determines dimensions of register table and loads first userform when report table is clicked
    
    'Reference register table
    Set RegTable = ThisWorkbook.Sheets("Register").ListObjects("Register")
    
    'Store current username in memory
    'Source: https://www.excelsirji.com/vba-code-to-get-logged-in-user-name/
    Username = Application.Username
    
    'Source: https://officetricks.com/excel-vba-get-username-windows-system/
    'Username = ThisWorkbook.BuiltinDocumentProperties("Author")
    
    
    'Force default starting rowIndex for empty form and tickbox checked
    RowIndex = RIndex
    Tick = True
    FC_Tick = True
    SAG_Tick = True
    
    'Set initial location
    UserFormTopPos = Application.Top + 25
    UserFormLeftPos = Application.Left + Application.Width / 3
    
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
    form00_Nav.show False
    
End Sub
