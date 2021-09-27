Attribute VB_Name = "Public_Variables"
Option Explicit

'List of public variables to share between userforms
'---------------------------------------------------

'Row of table being read
Public RowIndex As Long
Public RegTable As ListObject

'Version control
Public Username As String
Public LastUpdate As Date

'Conditional Navigation
Public Tick As Boolean
Public FC_Tick As Boolean
Public StudyStatus As Variant

'Search
Public DisplayArr() As Variant

'Undo delete
Public OldStudyStatus As String

'Userform Position storage
Public UserFormLeftPos As Long
Public UserFormTopPos As Long
Public UserFormLeftPosC As Long
Public UserFormTopPosC As Long
Public UserFormLeftPosR As Long
Public UserFormTopPosR As Long

'Userform dimension control
Public Const UHeight As Long = 470 '610
Public Const UWidth As Long = 650 '500


'list of public functions to share between userforms
'---------------------------------------------------

Public Sub LogLastAccess()
    
    'PURPOSE: Log last time entry was accessed
    
    If RowIndex > 0 Then
        With RegTable.ListRows(RowIndex)
            .Range(5) = Now
            .Range(6) = Username
        End With
    End If
End Sub

Public Function String_to_Date(Txt As String)
    'PURPOSE: Convert string input to date value if it is a valid date
    
    If IsDate(Txt) Then
        String_to_Date = DateValue(Txt)
    Else
        String_to_Date = Txt
    End If
    
End Function

Public Function Date_Validation(CurrDate As String, Optional PrevDate As String = "", Optional err2 As String = "") As String
    'PURPOSE: Assess data input is in correct format and output error message string
    
    Dim err As String
    Dim d1 As Variant
    Dim d2 As Variant
    
    err = vbNullString
    
    If CurrDate <> vbNullString And Not IsDate(CurrDate) Then
        err = "Please enter a valid date:" & Chr(10) & "DD-MMM-YYYY"
    End If
    
    d1 = String_to_Date(PrevDate)
    d2 = String_to_Date(CurrDate)
    
    'If no date entry issue, check date for chronology
    If err = "" And d1 <> "" And d2 <> "" And _
        IsDate(d1) And IsDate(d2) And d2 < d1 Then
        err = err2
    End If
    
    Date_Validation = err
    
End Function

Public Sub Fill_Completion_Status()

    'PURPOSE: Evaluate entry completion status
    
    Dim db As Range
    Dim ReadRow As Variant, Correct As Variant
    Dim i As Integer, cntTrue As Integer, cntEmpty As Integer
    Dim Status As Boolean
    
    'Exit if register is empty
    If RegTable.DataBodyRange Is Nothing Then
        Exit Sub
    End If
    
    Set db = RegTable.DataBodyRange
    
    'Initialise by making all values to be null
    Range(db.Cells(RowIndex, 129), db.Cells(RowIndex, 155)).Value = vbNullString
    
    'Tranpose twice to get 1D Array
    ReadRow = Application.Transpose(Application.Transpose(Range(db.Cells(RowIndex, 7), db.Cells(RowIndex, 125))))
    
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
                    
    'Apply correct test on each field
    For i = LBound(ReadRow) To UBound(ReadRow)
        If ReadRow(i) <> vbNullString Then
    
            Select Case Correct(i - 1)
                Case 0
                    ReadRow(i) = "Skip"
                Case 1
                    ReadRow(i) = Not (IsEmpty(ReadRow(i)))
                Case 2
                    ReadRow(i) = WorksheetFunction.IsText(ReadRow(i))
                Case 3
                    ReadRow(i) = IsDate(Format(ReadRow(i), "dd-mmm-yyyy"))
            End Select
            
        End If
    Next i
    
    
    'Completion status
    
    'Study Details
    'Criteria - all fields filled
    cntTrue = 0
    cntEmpty = 0
    For i = 2 To 6
        If ReadRow(i) = vbNullString Then
            cntEmpty = cntEmpty + 1
        ElseIf ReadRow(i) Then
            cntTrue = cntTrue + 1
        End If
    Next i
    
    If cntEmpty = 5 Then
        db.Cells(RowIndex, 129) = vbNullString
    ElseIf cntTrue = 5 Then
        db.Cells(RowIndex, 129) = True
    Else
        db.Cells(RowIndex, 129) = False
    End If
    
    'CDA
    'Criteria - all fields filled with dates
    cntTrue = 0
    cntEmpty = 0
    For i = 10 To 14
        If ReadRow(i) = vbNullString Then
            cntEmpty = cntEmpty + 1
        ElseIf ReadRow(i) Then
            cntTrue = cntTrue + 1
        End If
    Next i
    
    If cntEmpty = 5 Then
        db.Cells(RowIndex, 130) = vbNullString
    ElseIf cntTrue = 5 Then
        db.Cells(RowIndex, 130) = True
    Else
        db.Cells(RowIndex, 130) = False
    End If
    
        
    'Feasibility
    'Criteria - all fields filled with dates
    cntTrue = 0
    cntEmpty = 0
    For i = 18 To 20
        If ReadRow(i) = vbNullString Then
            cntEmpty = cntEmpty + 1
        ElseIf ReadRow(i) Then
            cntTrue = cntTrue + 1
        End If
    Next i
    
    If cntEmpty = 3 Then
        db.Cells(RowIndex, 131) = vbNullString
    ElseIf cntTrue = 3 Then
        db.Cells(RowIndex, 131) = True
    Else
        db.Cells(RowIndex, 131) = False
    End If
    
    'Site Selection
    'Criteria - all fields filled with dates and text (for combo box)
    cntTrue = 0
    cntEmpty = 0
    For i = 24 To 28
        If ReadRow(i) = vbNullString Then
            cntEmpty = cntEmpty + 1
        ElseIf ReadRow(i) Then
            cntTrue = cntTrue + 1
        End If
    Next i
    
    If cntEmpty = 5 Then
        db.Cells(RowIndex, 132) = vbNullString
    ElseIf cntTrue = 5 Then
        db.Cells(RowIndex, 132) = True
    Else
        db.Cells(RowIndex, 132) = False
    End If
    
    'Recruitment
    'Criteria - has to be date
    db.Cells(RowIndex, 133) = ReadRow(32)
    
    
    'CAHS Ethics
    'Criteria - all fields filled with dates
    cntTrue = 0
    cntEmpty = 0
    For i = 36 To 39
        If ReadRow(i) = vbNullString Then
            cntEmpty = cntEmpty + 1
        ElseIf ReadRow(i) Then
            cntTrue = cntTrue + 1
        End If
    Next i
    
    If cntEmpty = 4 Then
        db.Cells(RowIndex, 134) = vbNullString
    ElseIf cntTrue = 4 Then
        db.Cells(RowIndex, 134) = True
    Else
        db.Cells(RowIndex, 134) = False
    End If
    
    'NMA Ethics
    'Criteria - all fields filled with dates
    cntTrue = 0
    cntEmpty = 0
    For i = 41 To 43
        If ReadRow(i) = vbNullString Then
            cntEmpty = cntEmpty + 1
        ElseIf ReadRow(i) Then
            cntTrue = cntTrue + 1
        End If
    Next i
    
    If cntEmpty = 3 Then
        db.Cells(RowIndex, 135) = vbNullString
    ElseIf cntTrue = 3 Then
        db.Cells(RowIndex, 135) = True
    Else
        db.Cells(RowIndex, 135) = False
    End If
    
    'WNHS Ethics
    'Criteria - all fields filled with dates
    cntTrue = 0
    cntEmpty = 0
    For i = 45 To 46
        If ReadRow(i) = vbNullString Then
            cntEmpty = cntEmpty + 1
        ElseIf ReadRow(i) Then
            cntTrue = cntTrue + 1
        End If
    Next i
    
    If cntEmpty = 2 Then
        db.Cells(RowIndex, 136) = vbNullString
    ElseIf cntTrue = 2 Then
        db.Cells(RowIndex, 136) = True
    Else
        db.Cells(RowIndex, 136) = False
    End If
    
    'SJOG Ethics
    'Criteria - all fields filled with dates
    cntTrue = 0
    cntEmpty = 0
    For i = 48 To 49
        If ReadRow(i) = vbNullString Then
            cntEmpty = cntEmpty + 1
        ElseIf ReadRow(i) Then
            cntTrue = cntTrue + 1
        End If
    Next i
    
    If cntEmpty = 2 Then
        db.Cells(RowIndex, 137) = vbNullString
    ElseIf cntTrue = 2 Then
        db.Cells(RowIndex, 137) = True
    Else
        db.Cells(RowIndex, 137) = False
    End If
    
    'Other Ethics
    'Criteria - all fields filled with text (for committeee) and dates
    cntTrue = 0
    cntEmpty = 0
    For i = 51 To 53
        If ReadRow(i) = vbNullString Then
            cntEmpty = cntEmpty + 1
        ElseIf ReadRow(i) Then
            cntTrue = cntTrue + 1
        End If
    Next i
    
    If cntEmpty = 3 Then
        db.Cells(RowIndex, 138) = vbNullString
    ElseIf cntTrue = 3 Then
        db.Cells(RowIndex, 138) = True
    Else
        db.Cells(RowIndex, 138) = False
    End If
    
    'PCH Governance
    'Criteria - all fields filled with dates
    cntTrue = 0
    cntEmpty = 0
    For i = 57 To 59
        If ReadRow(i) = vbNullString Then
            cntEmpty = cntEmpty + 1
        ElseIf ReadRow(i) Then
            cntTrue = cntTrue + 1
        End If
    Next i
    
    If cntEmpty = 3 Then
        db.Cells(RowIndex, 139) = vbNullString
    ElseIf cntTrue = 3 Then
        db.Cells(RowIndex, 139) = True
    Else
        db.Cells(RowIndex, 139) = False
    End If
    
    'TKI Governance
    'Criteria - all fields filled with dates
    cntTrue = 0
    cntEmpty = 0
    For i = 61 To 63
        If ReadRow(i) = vbNullString Then
            cntEmpty = cntEmpty + 1
        ElseIf ReadRow(i) Then
            cntTrue = cntTrue + 1
        End If
    Next i
    
    If cntEmpty = 3 Then
        db.Cells(RowIndex, 140) = vbNullString
    ElseIf cntTrue = 3 Then
        db.Cells(RowIndex, 140) = True
    Else
        db.Cells(RowIndex, 140) = False
    End If
    
    'KEMH Governance
    'Criteria - all fields filled with dates
    cntTrue = 0
    cntEmpty = 0
    For i = 65 To 67
        If ReadRow(i) = vbNullString Then
            cntEmpty = cntEmpty + 1
        ElseIf ReadRow(i) Then
            cntTrue = cntTrue + 1
        End If
    Next i
    
    If cntEmpty = 3 Then
        db.Cells(RowIndex, 141) = vbNullString
    ElseIf cntTrue = 3 Then
        db.Cells(RowIndex, 141) = True
    Else
        db.Cells(RowIndex, 141) = False
    End If
    
    'SJOG Subiaco Governance
    'Criteria - all fields filled with dates
    cntTrue = 0
    cntEmpty = 0
    For i = 69 To 71
        If ReadRow(i) = vbNullString Then
            cntEmpty = cntEmpty + 1
        ElseIf ReadRow(i) Then
            cntTrue = cntTrue + 1
        End If
    Next i
    
    If cntEmpty = 3 Then
        db.Cells(RowIndex, 142) = vbNullString
    ElseIf cntTrue = 3 Then
        db.Cells(RowIndex, 142) = True
    Else
        db.Cells(RowIndex, 142) = False
    End If
    
    'SJOG Mt Lawley Governance
    'Criteria - all fields filled with dates
    cntTrue = 0
    cntEmpty = 0
    For i = 73 To 75
        If ReadRow(i) = vbNullString Then
            cntEmpty = cntEmpty + 1
        ElseIf ReadRow(i) Then
            cntTrue = cntTrue + 1
        End If
    Next i
    
    If cntEmpty = 3 Then
        db.Cells(RowIndex, 143) = vbNullString
    ElseIf cntTrue = 3 Then
        db.Cells(RowIndex, 143) = True
    Else
        db.Cells(RowIndex, 143) = False
    End If
    
    'SJOG Murdoch Governance
    'Criteria - all fields filled with dates
    cntTrue = 0
    cntEmpty = 0
    For i = 77 To 79
        If ReadRow(i) = vbNullString Then
            cntEmpty = cntEmpty + 1
        ElseIf ReadRow(i) Then
            cntTrue = cntTrue + 1
        End If
    Next i
    
    If cntEmpty = 3 Then
        db.Cells(RowIndex, 144) = vbNullString
    ElseIf cntTrue = 3 Then
        db.Cells(RowIndex, 144) = True
    Else
        db.Cells(RowIndex, 144) = False
    End If
    
    'Other Governance
    'Criteria - all fields filled with text (for committee) and dates
    cntTrue = 0
    cntEmpty = 0
    For i = 81 To 84
        If ReadRow(i) = vbNullString Then
            cntEmpty = cntEmpty + 1
        ElseIf ReadRow(i) Then
            cntTrue = cntTrue + 1
        End If
    Next i
    
    If cntEmpty = 4 Then
        db.Cells(RowIndex, 145) = vbNullString
    ElseIf cntTrue = 4 Then
        db.Cells(RowIndex, 145) = True
    Else
        db.Cells(RowIndex, 145) = False
    End If
    
    'VTG Budget
    'Criteria - all fields filled with dates
    cntTrue = 0
    cntEmpty = 0
    For i = 88 To 90
        If ReadRow(i) = vbNullString Then
            cntEmpty = cntEmpty + 1
        ElseIf ReadRow(i) Then
            cntTrue = cntTrue + 1
        End If
    Next i
    
    If cntEmpty = 3 Then
        db.Cells(RowIndex, 146) = vbNullString
    ElseIf cntTrue = 3 Then
        db.Cells(RowIndex, 146) = True
    Else
        db.Cells(RowIndex, 146) = False
    End If
    
    'TKI Budget
    'Criteria - has to be date
    db.Cells(RowIndex, 147) = ReadRow(92)
    
    
    'Pharmacy Budget
    'Criteria - all fields filled with dates
    cntTrue = 0
    cntEmpty = 0
    For i = 94 To 95
        If ReadRow(i) = vbNullString Then
            cntEmpty = cntEmpty + 1
        ElseIf ReadRow(i) Then
            cntTrue = cntTrue + 1
        End If
    Next i
    
    If cntEmpty = 2 Then
        db.Cells(RowIndex, 148) = vbNullString
    ElseIf cntTrue = 2 Then
        db.Cells(RowIndex, 148) = True
    Else
        db.Cells(RowIndex, 148) = False
    End If
    
    'Indemnity
    'Criteria - all fields filled with dates
    cntTrue = 0
    cntEmpty = 0
    For i = 99 To 101
        If ReadRow(i) = vbNullString Then
            cntEmpty = cntEmpty + 1
        ElseIf ReadRow(i) Then
            cntTrue = cntTrue + 1
        End If
    Next i
    
    If cntEmpty = 3 Then
        db.Cells(RowIndex, 149) = vbNullString
    ElseIf cntTrue = 3 Then
        db.Cells(RowIndex, 149) = True
    Else
        db.Cells(RowIndex, 149) = False
    End If
    
    'CTRA
    'Criteria - all fields filled with dates
    cntTrue = 0
    cntEmpty = 0
    For i = 105 To 111
        If ReadRow(i) = vbNullString Then
            cntEmpty = cntEmpty + 1
        ElseIf ReadRow(i) Then
            cntTrue = cntTrue + 1
        End If
    Next i
    
    If cntEmpty = 7 Then
        db.Cells(RowIndex, 150) = vbNullString
    ElseIf cntTrue = 7 Then
        db.Cells(RowIndex, 150) = True
    Else
        db.Cells(RowIndex, 150) = False
    End If
    
    'Financial Disclosure
    'Criteria - has to be date
    db.Cells(RowIndex, 151) = ReadRow(115)
    
    'SIV
    'Criteria - has to be date
    db.Cells(RowIndex, 152) = ReadRow(119)
    
    'Add table formulae
    'Overall Ethics true if at least one ethics committee complete
    db.Cells(RowIndex, 153).Formula = "=IF(COUNTA(Register[@[Ethics - CAHS Complete]:[Ethics - Others Complete]])=0, """"," & _
                                        "IF(COUNTIF(Register[@[Ethics - CAHS Complete]:[Ethics - Others Complete]],TRUE)>0,TRUE,FALSE))"
    
    'Overall Governance true if at least one ethics committee complete
    db.Cells(RowIndex, 154).Formula = "=IF(COUNTA(Register[@[Gov - PCH Complete]:[Gov - Others Complete]])=0,""""," & _
                                        "IF(COUNTIF(Register[@[Gov - PCH Complete]:[Gov - Others Complete]],TRUE)>0,TRUE,FALSE))"
    
    'Overall Budget true if at all budget committee approve
    db.Cells(RowIndex, 155).Formula = "=IF(COUNTA(Register[@[Budget - VTG Complete]:[Budget - Pharmacy Complete]])=0,""""," & _
                                        "IF(COUNTIF(Register[@[Budget - VTG Complete]:[Budget - Pharmacy Complete]],TRUE)=3,TRUE,FALSE))"
                              
    'Study complete if all core sections complete
    db.Cells(RowIndex, 156).Formula = "=IF(AND([@[Study Details Complete]]=TRUE,[@[CDA Complete]]=TRUE,[@[FS Complete]]=TRUE," & _
                                        "[@[Site Selection Complete]]=TRUE,[@[Recruitment Complete]]=TRUE,[@[Overall Ethics]]=TRUE," & _
                                        "[@[Overall Governance]]=TRUE,[@[Budget - VTG Complete]]=TRUE,[@[Budget - TKI Complete]]=TRUE," & _
                                        "[@[Budget - Pharmacy Complete]]=TRUE,[@[Indemnity Complete]]=TRUE,[@[CTRA Complete]]=TRUE," & _
                                        "[@[Fin Disc Complete]]=TRUE,[@[SIV Complete]]=TRUE),TRUE,FALSE)"
    
    'Fast cycle location based on last incomplete form. If none found then reverts to starting position
    db.Cells(RowIndex, 157).Formula = "=IFERROR(MATCH(FALSE,Register[@[Study Details Complete]:[SIV Complete]],0),1)"
        
End Sub
