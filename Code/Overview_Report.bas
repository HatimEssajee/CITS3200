Attribute VB_Name = "Overview_Report"
Option Explicit

Sub Bring_Data()
    
    'PURPOSE: Brings in all all data from register table in a summarized manner
    
    Dim Rpt As ListObject
    Dim err As Range, Header As Range, GreenCells As Range, RedCells As Range
    Dim ReadRow As Variant
    Dim i As Integer, j As Integer, k As Integer
    Dim cRed As Long, cGreen As Long, cWhite As Long, cRows As Long
    Dim tempStr As String, WriteArr() As String, GreenRef As String, RedRef As String, GreenArr() As String, RedArr() As String
    Dim StartTime As Double
    Dim MinutesElapsed As String
    
    'Turn off Settings to speed up
    'source: https://www.automateexcel.com/vba/turn-off-screen-updating/
    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.EnableEvents = False
    
    'Set color
    cRed = RGB(246, 176, 176)
    cGreen = RGB(146, 208, 80)
    cWhite = xlNone
    
    'Reference register table
    Set RegTable = Sheet_Register.ListObjects("Register")
    Set Rpt = Sheet_Report.ListObjects("Report")
    Set err = Sheet_Report.Cells(1, 3)
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
    
    cRows = Application.WorksheetFunction.CountIf(RegTable.DataBodyRange.Columns(7), "<>DELETED")
    
    'Exit out if all data is deleted
    If cRows = 0 Then
        err.Value = "Register table only has deleted values"
        Exit Sub
    End If
        
    'Initialise Write Array
    ReDim WriteArr(1 To cRows, 1 To 35)
    ReDim GreenArr(1 To cRows)
    ReDim RedArr(1 To cRows)
    
    'Apply filter to remove deleted
    ReadRow = RegTable.DataBodyRange
    
    For i = 1 To UBound(ReadRow)
        
        If ReadRow(i, 7) <> "DELETED" Then
            Application.StatusBar = "Processing Row " & j & "out of " & cRows
            
            GreenRef = ""
            RedRef = ""
            
            'Study Details
            WriteArr(j, 1) = ReadRow(i, 7)
            WriteArr(j, 2) = Format(ReadRow(i, 1), "dd-mmm-yyyy")
            WriteArr(j, 3) = ReadRow(i, 9)
            WriteArr(j, 4) = Format(ReadRow(i, 5), "dd-mmm-yyyy hh:mm")
            WriteArr(j, 5) = ReadRow(i, 8)
            WriteArr(j, 6) = ReadRow(i, 10)
            WriteArr(j, 7) = ReadRow(i, 11)
            WriteArr(j, 8) = ReadRow(i, 12)
            WriteArr(j, 35) = i
            
            'Add reminder in age column
            If ReadRow(i, 13) <> vbNullString And ReadRow(i, 129) = False Then
                    WriteArr(j, 8) = WriteArr(j, 8) & Chr(10) & Chr(10) & "Reminder:" & Chr(10) & _
                                            ReadRow(i, 13)
            End If
                
            'CDA
            If ReadRow(i, 130) <> vbNullString Then
                WriteArr(j, 9) = "Date Recv. Sponsor = " & Format(ReadRow(i, 16), "dd-mmm-yy") & Chr(10) & _
                                     "Date Sent Contracts = " & Format(ReadRow(i, 17), "dd-mmm-yy") & Chr(10) & _
                                     "Date Recv. Contracts = " & Format(ReadRow(i, 18), "dd-mmm-yy") & Chr(10) & _
                                     "Date Sent Sponsor = " & Format(ReadRow(i, 19), "dd-mmm-yy") & Chr(10) & _
                                     "Date Finalised = " & Format(ReadRow(i, 20), "dd-mmm-yy") & Chr(10)
                
                'Add reminder if incomplete
                If ReadRow(i, 130) = False And ReadRow(i, 21) <> vbNullString Then
                        WriteArr(j, 9) = WriteArr(j, 9) & Chr(10) & Chr(10) & "Reminder:" & Chr(10) & _
                                            ReadRow(i, 21)
                End If
            End If
            
            'Feasibility
            If ReadRow(i, 131) <> vbNullString Then
                WriteArr(j, 10) = "Date Recv. = " & Format(ReadRow(i, 24), "dd-mmm-yy") & Chr(10) & _
                                     "Date Completed = " & Format(ReadRow(i, 25), "dd-mmm-yy") & "; " & ReadRow(i, 26)
            
                'Add reminder if incomplete
                If ReadRow(i, 131) = False And ReadRow(i, 27) <> vbNullString Then
                        WriteArr(j, 10) = WriteArr(j, 10) & Chr(10) & Chr(10) & "Reminder:" & Chr(10) & _
                                            ReadRow(i, 27)
                End If
            End If
            
            'Site Selection
            If ReadRow(i, 132) <> vbNullString Then
                WriteArr(j, 11) = "Pre-study visit = " & Format(ReadRow(i, 30), "dd-mmm-yy") & "; " & ReadRow(i, 31) & Chr(10) & _
                                     "Valid. visit = " & Format(ReadRow(i, 32), "dd-mmm-yy") & "; " & ReadRow(i, 33) & Chr(10) & _
                                     "Date Site Selected = " & Format(ReadRow(i, 34), "dd-mmm-yy")
                
                'Add reminder if incomplete
                If ReadRow(i, 132) = False And ReadRow(i, 35) <> vbNullString Then
                        WriteArr(j, 11) = WriteArr(j, 11) & Chr(10) & Chr(10) & "Reminder:" & Chr(10) & _
                                            ReadRow(i, 35)
                End If
            End If
            
            'Recruitment
            If ReadRow(i, 133) <> vbNullString Then
                WriteArr(j, 12) = "Plan. Meeting = " & Format(ReadRow(i, 38), "dd-mmm-yy")
                
                'Add reminder if incomplete
                If ReadRow(i, 133) = False And ReadRow(i, 39) <> vbNullString Then
                        WriteArr(j, 12) = WriteArr(j, 12) & Chr(10) & Chr(10) & "Reminder:" & Chr(10) & _
                                            ReadRow(i, 39)
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
                
                WriteArr(j, 13) = tempStr
                
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
                WriteArr(j, 13) = tempStr
            End If
            
            
            'CAHS Ethics
            If ReadRow(i, 134) <> vbNullString Then
                WriteArr(j, 14) = "Date Submitted = " & Format(ReadRow(i, 42), "dd-mmm-yy") & Chr(10) & _
                                     "Date Responded = " & Format(ReadRow(i, 43), "dd-mmm-yy") & Chr(10) & _
                                     "Date Resubmitted = " & Format(ReadRow(i, 44), "dd-mmm-yy") & Chr(10) & _
                                     "Date Approved = " & Format(ReadRow(i, 45), "dd-mmm-yy")
                
                'Add reminder if incomplete
                If ReadRow(i, 134) = False And ReadRow(i, 46) <> vbNullString Then
                        WriteArr(j, 14) = WriteArr(j, 14) & Chr(10) & Chr(10) & "Reminder:" & Chr(10) & _
                                            ReadRow(i, 46)
                End If
            End If
            
            'NMA Ethics
            If ReadRow(i, 135) <> vbNullString Then
                WriteArr(j, 15) = "Ethics Commitee = " & ReadRow(i, 47) & Chr(10) & _
                                     "Date Submitted = " & Format(ReadRow(i, 48), "dd-mmm-yy") & Chr(10) & _
                                     "Date Approved = " & Format(ReadRow(i, 49), "dd-mmm-yy")
                
                'Add reminder if incomplete
                If ReadRow(i, 135) = False And ReadRow(i, 50) <> vbNullString Then
                        WriteArr(j, 15) = WriteArr(j, 15) & Chr(10) & Chr(10) & "Reminder:" & Chr(10) & _
                                            ReadRow(i, 50)
                End If
            End If
            
            'WNHS Ethics
            If ReadRow(i, 136) <> vbNullString Then
                WriteArr(j, 16) = "Date Submitted = " & Format(ReadRow(i, 51), "dd-mmm-yy") & Chr(10) & _
                                     "Date Approved = " & Format(ReadRow(i, 52), "dd-mmm-yy")
                If ReadRow(i, 136) = False And ReadRow(i, 53) <> vbNullString Then
                        WriteArr(j, 16) = WriteArr(j, 16) & Chr(10) & Chr(10) & "Reminder:" & Chr(10) & _
                                            ReadRow(i, 53)
                End If
            End If
            
            'SJOG Ethics
            If ReadRow(i, 137) <> vbNullString Then
                WriteArr(j, 17) = "Date Submitted = " & Format(ReadRow(i, 54), "dd-mmm-yy") & Chr(10) & _
                                     "Date Approved = " & Format(ReadRow(i, 55), "dd-mmm-yy")
                
                'Add reminder if incomplete
                If ReadRow(i, 137) And ReadRow(i, 56) <> vbNullString Then
                        WriteArr(j, 17) = WriteArr(j, 17) & Chr(10) & Chr(10) & "Reminder:" & Chr(10) & _
                                            ReadRow(i, 56)
                End If
            End If
            
            'Others Ethics
            If ReadRow(i, 138) <> vbNullString Then
                WriteArr(j, 18) = "Ethics Commitee = " & ReadRow(i, 57) & Chr(10) & _
                                     "Date Submitted = " & Format(ReadRow(i, 58), "dd-mmm-yy") & Chr(10) & _
                                     "Date Approved = " & Format(ReadRow(i, 59), "dd-mmm-yy")
                
                'Add reminder if incomplete
                If ReadRow(i, 138) = False And ReadRow(i, 60) <> vbNullString Then
                        WriteArr(j, 18) = WriteArr(j, 18) & Chr(10) & Chr(10) & "Reminder:" & Chr(10) & _
                                            ReadRow(i, 60)
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
                
                WriteArr(j, 19) = tempStr
            
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
                WriteArr(j, 19) = tempStr
            End If
            
            'PCH Governance
            If ReadRow(i, 139) <> vbNullString Then
                WriteArr(j, 20) = "Date Submitted = " & Format(ReadRow(i, 63), "dd-mmm-yy") & Chr(10) & _
                                     "Date Responded = " & Format(ReadRow(i, 64), "dd-mmm-yy") & Chr(10) & _
                                     "Date Approved = " & Format(ReadRow(i, 65), "dd-mmm-yy")
                
                'Add reminder if incomplete
                If ReadRow(i, 139) = False And ReadRow(i, 66) <> vbNullString Then
                        WriteArr(j, 20) = WriteArr(j, 20) & Chr(10) & Chr(10) & "Reminder:" & Chr(10) & _
                                            ReadRow(i, 66)
                End If
            End If
            
            'TKI Governance
            If ReadRow(i, 140) <> vbNullString Then
                WriteArr(j, 21) = "Date Submitted = " & Format(ReadRow(i, 67), "dd-mmm-yy") & Chr(10) & _
                                     "Date Responded = " & Format(ReadRow(i, 68), "dd-mmm-yy") & Chr(10) & _
                                     "Date Approved = " & Format(ReadRow(i, 69), "dd-mmm-yy")
                
                'Add reminder if incomplete
                If ReadRow(i, 140) = False And ReadRow(i, 70) <> vbNullString Then
                        WriteArr(j, 21) = WriteArr(j, 21) & Chr(10) & Chr(10) & "Reminder:" & Chr(10) & _
                                            ReadRow(i, 70)
                End If
            End If
            
            'KEMH Governance
            If ReadRow(i, 141) <> vbNullString Then
                WriteArr(j, 22) = "Date Submitted = " & Format(ReadRow(i, 71), "dd-mmm-yy") & Chr(10) & _
                                     "Date Responded = " & Format(ReadRow(i, 72), "dd-mmm-yy") & Chr(10) & _
                                     "Date Approved = " & Format(ReadRow(i, 73), "dd-mmm-yy")
                
                'Add reminder if incomplete
                If ReadRow(i, 141) = False And ReadRow(i, 74) <> vbNullString Then
                        WriteArr(j, 22) = WriteArr(j, 22) & Chr(10) & Chr(10) & "Reminder:" & Chr(10) & _
                                            ReadRow(i, 74)
                End If
            End If
            
            'SJOG Subiaco Governance
            If ReadRow(i, 142) <> vbNullString Then
                WriteArr(j, 23) = "Date Submitted = " & Format(ReadRow(i, 75), "dd-mmm-yy") & Chr(10) & _
                                     "Date Responded = " & Format(ReadRow(i, 76), "dd-mmm-yy") & Chr(10) & _
                                     "Date Approved = " & Format(ReadRow(i, 77), "dd-mmm-yy")
                
                'Add reminder if incomplete
                If ReadRow(i, 142) = False And ReadRow(i, 78) <> vbNullString Then
                        WriteArr(j, 23) = WriteArr(j, 23) & Chr(10) & Chr(10) & "Reminder:" & Chr(10) & _
                                            ReadRow(i, 78)
                End If
            End If
            
            'SJOG Mt Lawley Governance
            If ReadRow(i, 143) <> vbNullString Then
                WriteArr(j, 24) = "Date Submitted = " & Format(ReadRow(i, 79), "dd-mmm-yy") & Chr(10) & _
                                     "Date Responded = " & Format(ReadRow(i, 80), "dd-mmm-yy") & Chr(10) & _
                                     "Date Approved = " & Format(ReadRow(i, 81), "dd-mmm-yy")
                
                'Add reminder if incomplete
                If ReadRow(i, 143) = False And ReadRow(i, 82) <> vbNullString Then
                        WriteArr(j, 24) = WriteArr(j, 24) & Chr(10) & Chr(10) & "Reminder:" & Chr(10) & _
                                            ReadRow(i, 82)
                End If
            End If
            
            'SJOG Murdoch Governance
            If ReadRow(i, 144) <> vbNullString Then
                WriteArr(j, 25) = "Date Submitted = " & Format(ReadRow(i, 83), "dd-mmm-yy") & Chr(10) & _
                                     "Date Responded = " & Format(ReadRow(i, 84), "dd-mmm-yy") & Chr(10) & _
                                     "Date Approved = " & Format(ReadRow(i, 85), "dd-mmm-yy")
                
                'Add reminder if incomplete
                If ReadRow(i, 144) = False And ReadRow(i, 86) <> vbNullString Then
                        WriteArr(j, 25) = WriteArr(j, 25) & Chr(10) & Chr(10) & "Reminder:" & Chr(10) & _
                                            ReadRow(i, 86)
                End If
            End If
            
            'Others Governance
            If ReadRow(i, 145) <> vbNullString Then
                WriteArr(j, 26) = "Gov. Body = " & Format(ReadRow(i, 87), "dd-mmm-yy") & Chr(10) & _
                                     "Date Submitted = " & Format(ReadRow(i, 88), "dd-mmm-yy") & Chr(10) & _
                                     "Date Responded = " & Format(ReadRow(i, 89), "dd-mmm-yy") & Chr(10) & _
                                     "Date Approved = " & Format(ReadRow(i, 90), "dd-mmm-yy")
                
                'Add reminder if incomplete
                If ReadRow(i, 145) = False And ReadRow(i, 91) <> vbNullString Then
                        WriteArr(j, 26) = WriteArr(j, 26) & Chr(10) & Chr(10) & "Reminder:" & Chr(10) & _
                                            ReadRow(i, 91)
                End If
            End If
            
              
            'Indemnity
            If ReadRow(i, 149) <> vbNullString Then
                WriteArr(j, 27) = "Date Received = " & Format(ReadRow(i, 105), "dd-mmm-yy") & Chr(10) & _
                                     "Date Sent Contracts = " & Format(ReadRow(i, 106), "dd-mmm-yy") & Chr(10) & _
                                     "Date Completed = " & Format(ReadRow(i, 107), "dd-mmm-yy")
                If ReadRow(i, 149) = False And ReadRow(i, 108) <> vbNullString Then
                        WriteArr(j, 27) = WriteArr(j, 27) & Chr(10) & Chr(10) & "Reminder:" & Chr(10) & _
                                            ReadRow(i, 108)
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
                
                WriteArr(j, 28) = tempStr
                
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
                WriteArr(j, 28) = tempStr
            End If
            
            'VTG Budget
            If ReadRow(i, 146) <> vbNullString Then
                WriteArr(j, 29) = "Date Finalised = " & Format(ReadRow(i, 94), "dd-mmm-yy") & Chr(10) & _
                                     "Date Submitted Finance = " & Format(ReadRow(i, 95), "dd-mmm-yy") & Chr(10) & _
                                     "Date Approved = " & Format(ReadRow(i, 96), "dd-mmm-yy")
                
                'Add reminder if incomplete
                If ReadRow(i, 146) = False And ReadRow(i, 97) <> vbNullString Then
                        WriteArr(j, 29) = WriteArr(j, 29) & Chr(10) & Chr(10) & "Reminder:" & Chr(10) & _
                                            ReadRow(i, 97)
                End If
            End If
            
            'TKI Budget
            If ReadRow(i, 147) <> vbNullString Then
                WriteArr(j, 30) = "Date Approved = " & Format(ReadRow(i, 98), "dd-mmm-yy")
                
                'Add reminder if incomplete
                If ReadRow(i, 147) = False And ReadRow(i, 99) <> vbNullString Then
                        WriteArr(j, 30) = WriteArr(j, 30) & Chr(10) & Chr(10) & "Reminder:" & Chr(10) & _
                                            ReadRow(i, 99)
                End If
            End If
            
            'Pharmacy Budget
            If ReadRow(i, 148) <> vbNullString Then
                WriteArr(j, 31) = "Date Quote Recv. = " & Format(ReadRow(i, 100), "dd-mmm-yy") & Chr(10) & _
                                     "Date PO Finalised = " & Format(ReadRow(i, 101), "dd-mmm-yy")
                
                'Add reminder if incomplete
                If ReadRow(i, 148) = False And ReadRow(i, 102) <> vbNullString Then
                        WriteArr(j, 31) = WriteArr(j, 31) & Chr(10) & Chr(10) & "Reminder:" & Chr(10) & _
                                            ReadRow(i, 102)
                End If
            End If
            
            'CTRA
            If ReadRow(i, 150) <> vbNullString Then
                WriteArr(j, 32) = "Date Submitted R,G&C = " & Format(ReadRow(i, 111), "dd-mmm-yy") & Chr(10) & _
                                     "Date UWA Review = " & Format(ReadRow(i, 112), "dd-mmm-yy") & Chr(10) & _
                                     "Date Finance Review = " & Format(ReadRow(i, 113), "dd-mmm-yy") & Chr(10) & _
                                     "Date COO/CFO Sign-off = " & Format(ReadRow(i, 114), "dd-mmm-yy") & Chr(10) & _
                                     "Date VTG Sign-off = " & Format(ReadRow(i, 115), "dd-mmm-yy") & Chr(10) & _
                                     "Date Subm. Company = " & Format(ReadRow(i, 116), "dd-mmm-yy") & Chr(10) & _
                                     "Date Finalised = " & Format(ReadRow(i, 117), "dd-mmm-yy")
                
                'Add reminder if incomplete
                If ReadRow(i, 150) = False And ReadRow(i, 118) <> vbNullString Then
                        WriteArr(j, 32) = WriteArr(j, 32) & Chr(10) & Chr(10) & "Reminder:" & Chr(10) & _
                                            ReadRow(i, 118)
                End If
            End If
            
            'Financial Disclosure
            If ReadRow(i, 151) <> vbNullString Then
                WriteArr(j, 33) = "Date Completed = " & Format(ReadRow(i, 121), "dd-mmm-yy")
                
                'Add reminder if incomplete
                If ReadRow(i, 151) = False And ReadRow(i, 122) <> vbNullString Then
                        WriteArr(j, 33) = WriteArr(j, 33) & Chr(10) & Chr(10) & "Reminder:" & Chr(10) & _
                                            ReadRow(i, 122)
                End If
            End If
            
            'SIV
            If ReadRow(i, 152) <> vbNullString Then
                WriteArr(j, 34) = "Date of SIV = " & Format(ReadRow(i, 125), "dd-mmm-yy")
                
                'Add reminder if incomplete
                If ReadRow(i, 152) = False And ReadRow(i, 126) <> vbNullString Then
                        WriteArr(j, 34) = WriteArr(j, 34) & Chr(10) & Chr(10) & "Reminder:" & Chr(10) & _
                                            ReadRow(i, 126)
                End If
            End If
            
            'Build colour reference
            'Study colour
            If ReadRow(i, 156) = True Then
                GreenRef = MergeRef(GreenRef, "A" & j + 11)
            ElseIf ReadRow(i, 156) <> vbNullString Then
                RedRef = MergeRef(RedRef, "A" & j + 11)
            End If

            'Study details colour
            For k = 5 To 8
                If WriteArr(j, k) = vbNullString Then
                    RedRef = MergeRef(RedRef, Chr(64 + k) & j + 11)
                End If
            Next k

            'CDA
            If ReadRow(i, 130) = True Then
                GreenRef = MergeRef(GreenRef, "I" & j + 11)
            ElseIf ReadRow(i, 130) <> vbNullString Then
                RedRef = MergeRef(RedRef, "I" & j + 11)
            End If

            'Feasibility
            If ReadRow(i, 131) = True Then
                GreenRef = MergeRef(GreenRef, "J" & j + 11)
            ElseIf ReadRow(i, 131) <> vbNullString Then
                RedRef = MergeRef(RedRef, "J" & j + 11)
            End If

            'Site Selection
             If ReadRow(i, 132) = True Then
                GreenRef = MergeRef(GreenRef, "K" & j + 11)
            ElseIf ReadRow(i, 132) <> vbNullString Then
                RedRef = MergeRef(RedRef, "K" & j + 11)
            End If

            'Recruitment
            If ReadRow(i, 133) = True Then
                GreenRef = MergeRef(GreenRef, "L" & j + 11)
            ElseIf ReadRow(i, 133) <> vbNullString Then
                RedRef = MergeRef(RedRef, "L" & j + 11)
            End If

            'Overall Ethics
            If ReadRow(i, 153) = True Then
                GreenRef = MergeRef(GreenRef, "M" & j + 11)
            ElseIf ReadRow(i, 153) <> vbNullString Then
                RedRef = MergeRef(RedRef, "M" & j + 11)
            End If

            'CAHS Ethics
            If ReadRow(i, 134) = True Then
                GreenRef = MergeRef(GreenRef, "N" & j + 11)
            ElseIf ReadRow(i, 134) <> vbNullString Then
                RedRef = MergeRef(RedRef, "N" & j + 11)
            End If

            'NMA Ethics
            If ReadRow(i, 135) = True Then
                GreenRef = MergeRef(GreenRef, "O" & j + 11)
            ElseIf ReadRow(i, 135) <> vbNullString Then
                RedRef = MergeRef(RedRef, "O" & j + 11)
            End If

            'WNHS Ethics
            If ReadRow(i, 136) = True Then
                GreenRef = MergeRef(GreenRef, "P" & j + 11)
            ElseIf ReadRow(i, 136) <> vbNullString Then
                RedRef = MergeRef(RedRef, "P" & j + 11)
            End If

            'SJOG Ethics
            If ReadRow(i, 137) = True Then
                GreenRef = MergeRef(GreenRef, "Q" & j + 11)
            ElseIf ReadRow(i, 137) <> vbNullString Then
                RedRef = MergeRef(RedRef, "Q" & j + 11)
            End If

            'Others Ethics
            If ReadRow(i, 138) = True Then
                GreenRef = MergeRef(GreenRef, "R" & j + 11)
            ElseIf ReadRow(i, 138) <> vbNullString Then
                RedRef = MergeRef(RedRef, "R" & j + 11)
            End If

            'Overall Governance
            If ReadRow(i, 154) = True Then
                GreenRef = MergeRef(GreenRef, "S" & j + 11)
            ElseIf ReadRow(i, 154) <> vbNullString Then
                RedRef = MergeRef(RedRef, "S" & j + 11)
            End If

            'PCH Governance
            If ReadRow(i, 139) = True Then
                GreenRef = MergeRef(GreenRef, "T" & j + 11)
            ElseIf ReadRow(i, 139) <> vbNullString Then
                RedRef = MergeRef(RedRef, "T" & j + 11)
            End If

            'TKI Governance
            If ReadRow(i, 140) = True Then
                GreenRef = MergeRef(GreenRef, "U" & j + 11)
            ElseIf ReadRow(i, 140) <> vbNullString Then
                RedRef = MergeRef(RedRef, "U" & j + 11)
            End If

            'KEMH Governance
            If ReadRow(i, 141) = True Then
                GreenRef = MergeRef(GreenRef, "V" & j + 11)
            ElseIf ReadRow(i, 141) <> vbNullString Then
                RedRef = MergeRef(RedRef, "V" & j + 11)
            End If

            'SJOG Subiaco Governance
            If ReadRow(i, 142) = True Then
                GreenRef = MergeRef(GreenRef, "W" & j + 11)
            ElseIf ReadRow(i, 142) <> vbNullString Then
                RedRef = MergeRef(RedRef, "W" & j + 11)
            End If

            'SJOG Mt Lawley Governance
            If ReadRow(i, 143) = True Then
                GreenRef = MergeRef(GreenRef, "X" & j + 11)
            ElseIf ReadRow(i, 143) <> vbNullString Then
                RedRef = MergeRef(RedRef, "X" & j + 11)
            End If

            'SJOG Murdoch Governance
            If ReadRow(i, 144) = True Then
                GreenRef = MergeRef(GreenRef, "Y" & j + 11)
            ElseIf ReadRow(i, 144) <> vbNullString Then
                RedRef = MergeRef(RedRef, "Y" & j + 11)
            End If

            'Others Governance
            If ReadRow(i, 145) = True Then
                GreenRef = MergeRef(GreenRef, "Z" & j + 11)
            ElseIf ReadRow(i, 145) <> vbNullString Then
                RedRef = MergeRef(RedRef, "Z" & j + 11)
            End If

            'Indemnity
            If ReadRow(i, 149) = True Then
                GreenRef = MergeRef(GreenRef, "AA" & j + 11)
            ElseIf ReadRow(i, 149) <> vbNullString Then
                RedRef = MergeRef(RedRef, "AA" & j + 11)
            End If

            'Overall Budget
            If ReadRow(i, 155) = True Then
                GreenRef = MergeRef(GreenRef, "AB" & j + 11)
            ElseIf ReadRow(i, 155) <> vbNullString Then
                RedRef = MergeRef(RedRef, "AB" & j + 11)
            End If

            'VTG Budget
            If ReadRow(i, 146) = True Then
                GreenRef = MergeRef(GreenRef, "AC" & j + 11)
            ElseIf ReadRow(i, 146) <> vbNullString Then
                RedRef = MergeRef(RedRef, "AC" & j + 11)
            End If

            'TKI Budget
            If ReadRow(i, 147) = True Then
                GreenRef = MergeRef(GreenRef, "AD" & j + 11)
            ElseIf ReadRow(i, 147) <> vbNullString Then
                RedRef = MergeRef(RedRef, "AD" & j + 11)
            End If

            'Pharmacy Budget
            If ReadRow(i, 148) = True Then
                GreenRef = MergeRef(GreenRef, "AE" & j + 11)
            ElseIf ReadRow(i, 148) <> vbNullString Then
                RedRef = MergeRef(RedRef, "AE" & j + 11)
            End If

            'CTRA
            If ReadRow(i, 150) = True Then
                GreenRef = MergeRef(GreenRef, "AF" & j + 11)
            ElseIf ReadRow(i, 150) <> vbNullString Then
                RedRef = MergeRef(RedRef, "AF" & j + 11)
            End If

            'Financial Disclosure
            If ReadRow(i, 151) = True Then
                GreenRef = MergeRef(GreenRef, "AG" & j + 11)
            ElseIf ReadRow(i, 151) <> vbNullString Then
                RedRef = MergeRef(RedRef, "AG" & j + 11)
            End If

            'SIV
            If ReadRow(i, 152) = True Then
                GreenRef = MergeRef(GreenRef, "AH" & j + 11)
            ElseIf ReadRow(i, 152) <> vbNullString Then
                RedRef = MergeRef(RedRef, "AH" & j + 11)
            End If

            'Store string range reference
            GreenArr(j) = GreenRef
            RedArr(j) = RedRef
            
            j = j + 1
        End If
    Next i
    
    Application.StatusBar = "Copying data into report"
    
    'Write into table
    'SOURCE: https://stackoverflow.com/questions/37603174/what-is-the-fastest-way-to-unload-a-2-dimensional-array-into-an-excel-worksheet
    Rpt.ListRows.Add.Range.Resize(cRows, 35).Value = WriteArr
    
    Application.StatusBar = "Applying colours"
    
    For i = LBound(GreenArr) To UBound(GreenArr)
        If GreenArr(i) <> vbNullString Then
            Set GreenCells = Sheet_Report.Range(GreenArr(i))
            
            With GreenCells
                .Borders.LineStyle = xlContinuous
                .Borders.Weight = xlThin
                .Interior.Color = cGreen
                .WrapText = True
                .HorizontalAlignment = xlHAlignLeft
                .VerticalAlignment = xlVAlignTop
                .Font.Bold = False
            End With
        End If
        
        If RedArr(i) <> vbNullString Then
            Set RedCells = Sheet_Report.Range(RedArr(i))
        
            With RedCells
                .Borders.LineStyle = xlContinuous
                .Borders.Weight = xlThin
                .Interior.Color = cRed
                .WrapText = True
                .HorizontalAlignment = xlHAlignLeft
                .VerticalAlignment = xlVAlignTop
                .Font.Bold = False
            End With
        End If
      
    Next i
    
    'Revert Status bar
    Application.StatusBar = False
    
    'Reinstate Settings
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.EnableEvents = True
    
    MsgBox ("Data Succesfully Imported")
    err.Font.Color = vbBlack
    err.Value = "Data retrieved " & Format(Now, "dd-mmm-yyyy hh:mm AM/PM")
End Sub

Private Function MergeRef(ref As String, Add As String)
    'PURPOSE: To merge string reference with changes
    
    If ref = vbNullString Then
        ref = Add
    Else
        ref = ref & "," & Add
    End If
    
    MergeRef = ref
    
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
