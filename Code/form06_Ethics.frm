VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form06_Ethics 
   Caption         =   "Ethics Review"
   ClientHeight    =   5400
   ClientLeft      =   -468
   ClientTop       =   -2184
   ClientWidth     =   8664.001
   OleObjectBlob   =   "form06_Ethics.frx":0000
End
Attribute VB_Name = "form06_Ethics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub UserForm_Activate()
    'PURPOSE: Reposition userform to Top Left of application Window and fix size
    'source: https://www.mrexcel.com/board/threads/userform-startup-position.671108/
    Me.Top = UserFormTopPos
    Me.Left = UserFormLeftPos
    Me.Height = UHeight
    Me.Width = UWidth

End Sub

Private Sub UserForm_Deactivate()
    'Store form position
    UserFormTopPos = Me.Top
    UserFormLeftPos = Me.Left
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    'PURPOSE: On Close Userform this code saves the last Userform position to Defined Names
    'SOURCE: https://answers.microsoft.com/en-us/msoffice/forum/all/saving-last-position-of-userform/9399e735-9a9e-47c4-a1e0-e0d5cedd15ca
    UserFormTopPos = Me.Top
    UserFormLeftPos = Me.Left
End Sub

Private Sub UserForm_Initialize()
    'PURPOSE: Clear form on initialization
    'Source: https://www.contextures.com/xlUserForm02.html
    'Source: https://www.contextures.com/Excel-VBA-ComboBox-Lists.html
    Dim ctrl As MSForms.Control
    Dim pPage As MSForms.Page
       
    'Turn off Settings to speed up
    'source: https://www.automateexcel.com/vba/turn-off-screen-updating/
    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayStatusBar = False
    Application.EnableEvents = False
    
    'Erase and initialise arrays
    ReDim OldValues(1 To 19)
    ReDim NxtOldValues(1 To 19)
                    
    'Clear user form
    'SOURCE: https://www.mrexcel.com/board/threads/loop-through-controls-on-a-userform.427103/
    For Each ctrl In Me.Controls
        Select Case True
                Case TypeOf ctrl Is MSForms.TextBox
                    ctrl.Value = ""
                Case TypeOf ctrl Is MSForms.Label
                    'Empty error captions
                    If Left(ctrl.Name, 3) = "err" Then
                        ctrl.Caption = ""
                    End If
                Case TypeOf ctrl Is MSForms.ComboBox
                    ctrl.Value = ""
                    ctrl.Clear
                Case TypeOf ctrl Is MSForms.ListBox
                    ctrl.Value = ""
                    ctrl.Clear
            End Select
    Next ctrl
    
    For Each pPage In Me.multiEthics.Pages
        For Each ctrl In pPage.Controls
            Select Case True
                Case TypeOf ctrl Is MSForms.CheckBox
                    ctrl.Value = False
                Case TypeOf ctrl Is MSForms.TextBox
                    ctrl.Value = ""
                Case TypeOf ctrl Is MSForms.ComboBox
                    ctrl.Value = ""
                Case TypeOf ctrl Is MSForms.ListBox
                    ctrl.Value = ""
            End Select
                
        Next ctrl
    Next pPage
    
    'Read information from register table
    With RegTable.ListRows(RowIndex)
        Me.txtStudyName.Value = .Range(9).Value
        
        Me.txtCAHS_Date_Submitted.Value = ReadDate(.Range(42).Value)
        Me.txtCAHS_Date_Responded.Value = ReadDate(.Range(43).Value)
        Me.txtCAHS_Date_Resubmitted.Value = ReadDate(.Range(44).Value)
        Me.txtCAHS_Date_Approved.Value = ReadDate(.Range(45).Value)
        Me.txtCAHS_Reminder.Value = .Range(46).Value
        
        Me.txtNMA_Committee.Value = .Range(47).Value
        Me.txtNMA_Date_Submitted.Value = ReadDate(.Range(48).Value)
        Me.txtNMA_Date_Approved.Value = ReadDate(.Range(49).Value)
        Me.txtNMA_Reminder.Value = .Range(50).Value
        
        Me.txtWNHS_Date_Submitted.Value = ReadDate(.Range(51).Value)
        Me.txtWNHS_Date_Approved.Value = ReadDate(.Range(52).Value)
        Me.txtWNHS_Reminder.Value = .Range(53).Value
        
        Me.txtSJOG_Date_Submitted.Value = ReadDate(.Range(54).Value)
        Me.txtSJOG_Date_Approved.Value = ReadDate(.Range(55).Value)
        Me.txtSJOG_Reminder.Value = .Range(56).Value
        
        Me.txtOthers_Committee.Value = .Range(57).Value
        Me.txtOthers_Date_Submitted.Value = ReadDate(.Range(58).Value)
        Me.txtOthers_Date_Approved.Value = ReadDate(.Range(59).Value)
        Me.txtOthers_Reminder.Value = .Range(60).Value
        
    End With
    
    'Populate Old Values Array - for undo
    OldValues(1) = String_to_Date(Me.txtCAHS_Date_Submitted.Value)
    OldValues(2) = String_to_Date(Me.txtCAHS_Date_Responded.Value)
    OldValues(3) = String_to_Date(Me.txtCAHS_Date_Resubmitted.Value)
    OldValues(4) = String_to_Date(Me.txtCAHS_Date_Approved.Value)
    OldValues(5) = Me.txtCAHS_Reminder.Value
    
    OldValues(6) = Me.txtNMA_Committee.Value
    OldValues(7) = String_to_Date(Me.txtNMA_Date_Submitted.Value)
    OldValues(8) = String_to_Date(Me.txtNMA_Date_Approved.Value)
    OldValues(9) = Me.txtNMA_Reminder.Value
    
    OldValues(10) = String_to_Date(Me.txtWNHS_Date_Submitted.Value)
    OldValues(11) = String_to_Date(Me.txtWNHS_Date_Approved.Value)
    OldValues(12) = Me.txtWNHS_Reminder.Value
    
    OldValues(13) = String_to_Date(Me.txtSJOG_Date_Submitted.Value)
    OldValues(14) = String_to_Date(Me.txtSJOG_Date_Approved.Value)
    OldValues(15) = Me.txtSJOG_Reminder.Value
    
    OldValues(16) = Me.txtOthers_Committee.Value
    OldValues(17) = String_to_Date(Me.txtOthers_Date_Submitted.Value)
    OldValues(18) = String_to_Date(Me.txtOthers_Date_Approved.Value)
    OldValues(19) = Me.txtOthers_Reminder.Value
    
    'Initialize NxtOldValues
    NxtOldValues = OldValues
    
    'Access version control
    Call LogLastAccess
    
    'Depress and make toggle green on nav bar
    Me.tglEthics.Value = True
    Me.tglEthics.BackColor = vbGreen
    
    'Allocate tick box values
    Me.cbSaveonUnload.Value = SAG_Tick
    
    'Run date validation on data entered
    Call txtCAHS_Date_Submitted_AfterUpdate
    Call txtCAHS_Date_Responded_AfterUpdate
    Call txtCAHS_Date_Resubmitted_AfterUpdate
    Call txtCAHS_Date_Approved_AfterUpdate
    
    Call txtNMA_Date_Submitted_AfterUpdate
    Call txtNMA_Date_Approved_AfterUpdate
    
    Call txtWNHS_Date_Submitted_AfterUpdate
    Call txtWNHS_Date_Approved_AfterUpdate
    
    Call txtSJOG_Date_Submitted_AfterUpdate
    Call txtSJOG_Date_Approved_AfterUpdate
    
    Call txtOthers_Date_Submitted_AfterUpdate
    Call txtOthers_Date_Approved_AfterUpdate
    
    'Reinstate Settings
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayStatusBar = True
    Application.EnableEvents = True
    
End Sub

Private Sub txtCAHS_Date_Submitted_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtCAHS_Date_Submitted.Value)
    
    'Display error message
    Me.errCAHS_Date_Submitted.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtCAHS_Date_Submitted.Value) Then
        Me.txtCAHS_Date_Submitted.Value = Format(Me.txtCAHS_Date_Submitted.Value, "dd-mmm-yyyy")
    End If
    
End Sub

Private Sub txtCAHS_Date_Responded_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtCAHS_Date_Responded.Value, Me.txtCAHS_Date_Submitted.Value, _
            "Date entered earlier than date Submitted")

    'Display error message
    Me.errCAHS_Date_Responded.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtCAHS_Date_Responded.Value) Then
        Me.txtCAHS_Date_Responded.Value = Format(Me.txtCAHS_Date_Responded.Value, "dd-mmm-yyyy")
    End If
     
End Sub

Private Sub txtCAHS_Date_Resubmitted_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtCAHS_Date_Resubmitted.Value, Me.txtCAHS_Date_Responded.Value, _
            "Date entered earlier than date Responded")

    'Display error message
    Me.errCAHS_Date_Resubmitted.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtCAHS_Date_Resubmitted.Value) Then
        Me.txtCAHS_Date_Resubmitted.Value = Format(Me.txtCAHS_Date_Resubmitted.Value, "dd-mmm-yyyy")
    End If
     
End Sub

Private Sub txtCAHS_Date_Approved_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtCAHS_Date_Approved.Value, Me.txtCAHS_Date_Submitted.Value, _
            "Date entered earlier than date Submitted")

    'Display error message
    Me.errCAHS_Date_Approved.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtCAHS_Date_Approved.Value) Then
        Me.txtCAHS_Date_Approved.Value = Format(Me.txtCAHS_Date_Approved.Value, "dd-mmm-yyyy")
    End If
     
End Sub

Private Sub txtNMA_Date_Submitted_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtNMA_Date_Submitted.Value)
    
    'Display error message
    Me.errNMA_Date_Submitted.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtNMA_Date_Submitted.Value) Then
        Me.txtNMA_Date_Submitted.Value = Format(Me.txtNMA_Date_Submitted.Value, "dd-mmm-yyyy")
    End If
    
End Sub

Private Sub txtNMA_Date_Approved_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtNMA_Date_Approved.Value, Me.txtNMA_Date_Submitted.Value, _
            "Date entered earlier than date Submitted")

    'Display error message
    Me.errNMA_Date_Approved.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtNMA_Date_Approved.Value) Then
        Me.txtNMA_Date_Approved.Value = Format(Me.txtNMA_Date_Approved.Value, "dd-mmm-yyyy")
    End If
     
End Sub

Private Sub txtWNHS_Date_Submitted_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtWNHS_Date_Submitted.Value)
    
    'Display error message
    Me.errWNHS_Date_Submitted.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtWNHS_Date_Submitted.Value) Then
        Me.txtWNHS_Date_Submitted.Value = Format(Me.txtWNHS_Date_Submitted.Value, "dd-mmm-yyyy")
    End If
    
End Sub

Private Sub txtWNHS_Date_Approved_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtWNHS_Date_Approved.Value, Me.txtWNHS_Date_Submitted.Value, _
            "Date entered earlier than date Submitted")

    'Display error message
    Me.errWNHS_Date_Approved.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtWNHS_Date_Approved.Value) Then
        Me.txtWNHS_Date_Approved.Value = Format(Me.txtWNHS_Date_Approved.Value, "dd-mmm-yyyy")
    End If
     
End Sub

Private Sub txtSJOG_Date_Submitted_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtSJOG_Date_Submitted.Value)
    
    'Display error message
    Me.errSJOG_Date_Submitted.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtSJOG_Date_Submitted.Value) Then
        Me.txtSJOG_Date_Submitted.Value = Format(Me.txtSJOG_Date_Submitted.Value, "dd-mmm-yyyy")
    End If
    
End Sub

Private Sub txtSJOG_Date_Approved_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtSJOG_Date_Approved.Value, Me.txtSJOG_Date_Submitted.Value, _
            "Date entered earlier than date Submitted")

    'Display error message
    Me.errSJOG_Date_Approved.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtSJOG_Date_Approved.Value) Then
        Me.txtSJOG_Date_Approved.Value = Format(Me.txtSJOG_Date_Approved.Value, "dd-mmm-yyyy")
    End If
     
End Sub

Private Sub txtOthers_Date_Submitted_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtOthers_Date_Submitted.Value)
    
    'Display error message
    Me.errOthers_Date_Submitted.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtOthers_Date_Submitted.Value) Then
        Me.txtOthers_Date_Submitted.Value = Format(Me.txtOthers_Date_Submitted.Value, "dd-mmm-yyyy")
    End If
    
End Sub

Private Sub txtOthers_Date_Approved_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtOthers_Date_Approved.Value, Me.txtOthers_Date_Submitted.Value, _
            "Date entered earlier than date Submitted")

    'Display error message
    Me.errOthers_Date_Approved.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtOthers_Date_Approved.Value) Then
        Me.txtOthers_Date_Approved.Value = Format(Me.txtOthers_Date_Approved.Value, "dd-mmm-yyyy")
    End If
     
End Sub

Private Sub cmdUndo_Click()
    'PURPOSE: Recall values read from register table when the form was loaded initially
    
    Me.txtCAHS_Date_Submitted.Value = ReadDate(CStr(OldValues(1)))
    Me.txtCAHS_Date_Responded.Value = ReadDate(CStr(OldValues(2)))
    Me.txtCAHS_Date_Resubmitted.Value = ReadDate(CStr(OldValues(3)))
    Me.txtCAHS_Date_Approved.Value = ReadDate(CStr(OldValues(4)))
    Me.txtCAHS_Reminder.Value = OldValues(5)
    
    Me.txtNMA_Committee.Value = OldValues(6)
    Me.txtNMA_Date_Submitted.Value = ReadDate(CStr(OldValues(7)))
    Me.txtNMA_Date_Approved.Value = ReadDate(CStr(OldValues(8)))
    Me.txtNMA_Reminder.Value = OldValues(9)
    
    Me.txtWNHS_Date_Submitted.Value = ReadDate(CStr(OldValues(10)))
    Me.txtWNHS_Date_Approved.Value = ReadDate(CStr(OldValues(11)))
    Me.txtWNHS_Reminder.Value = OldValues(12)
    
    Me.txtSJOG_Date_Submitted.Value = ReadDate(CStr(OldValues(13)))
    Me.txtSJOG_Date_Approved.Value = ReadDate(CStr(OldValues(14)))
    Me.txtSJOG_Reminder.Value = OldValues(15)
    
    Me.txtOthers_Committee.Value = OldValues(16)
    Me.txtOthers_Date_Submitted.Value = ReadDate(CStr(OldValues(17)))
    Me.txtOthers_Date_Approved.Value = ReadDate(CStr(OldValues(18)))
    Me.txtOthers_Reminder.Value = OldValues(19)
    
End Sub

Private Sub cmdRedo_Click()
    'PURPOSE: Recall values replaced by undo
    
    Me.txtCAHS_Date_Submitted.Value = ReadDate(CStr(NxtOldValues(1)))
    Me.txtCAHS_Date_Responded.Value = ReadDate(CStr(NxtOldValues(2)))
    Me.txtCAHS_Date_Resubmitted.Value = ReadDate(CStr(NxtOldValues(3)))
    Me.txtCAHS_Date_Approved.Value = ReadDate(CStr(NxtOldValues(4)))
    Me.txtCAHS_Reminder.Value = NxtOldValues(5)
    
    Me.txtNMA_Committee.Value = NxtOldValues(6)
    Me.txtNMA_Date_Submitted.Value = ReadDate(CStr(NxtOldValues(7)))
    Me.txtNMA_Date_Approved.Value = ReadDate(CStr(NxtOldValues(8)))
    Me.txtNMA_Reminder.Value = NxtOldValues(9)
    
    Me.txtWNHS_Date_Submitted.Value = ReadDate(CStr(NxtOldValues(10)))
    Me.txtWNHS_Date_Approved.Value = ReadDate(CStr(NxtOldValues(11)))
    Me.txtWNHS_Reminder.Value = NxtOldValues(12)
    
    Me.txtSJOG_Date_Submitted.Value = ReadDate(CStr(NxtOldValues(13)))
    Me.txtSJOG_Date_Approved.Value = ReadDate(CStr(NxtOldValues(14)))
    Me.txtSJOG_Reminder.Value = NxtOldValues(15)
    
    Me.txtOthers_Committee.Value = NxtOldValues(16)
    Me.txtOthers_Date_Submitted.Value = ReadDate(CStr(NxtOldValues(17)))
    Me.txtOthers_Date_Approved.Value = ReadDate(CStr(NxtOldValues(18)))
    Me.txtOthers_Reminder.Value = NxtOldValues(19)
    
End Sub

Private Sub cmdClose_Click()
    'PURPOSE: Closes current form
    
    Unload Me
    
End Sub

Private Sub cmdEdit_Click()
    'PURPOSE: Apply changes into Register table when edit button clicked
    
    'Overwrite prev. old values with new backup values
    OldValues = NxtOldValues
    
    'Apply changes
    Call UpdateRegister
    DoEvents
    
    'Access version control
    Call LogLastAccess
                    
    'Run date validation on data entered
    Call txtCAHS_Date_Submitted_AfterUpdate
    Call txtCAHS_Date_Responded_AfterUpdate
    Call txtCAHS_Date_Resubmitted_AfterUpdate
    Call txtCAHS_Date_Approved_AfterUpdate
    
    Call txtNMA_Date_Submitted_AfterUpdate
    Call txtNMA_Date_Approved_AfterUpdate
    
    Call txtWNHS_Date_Submitted_AfterUpdate
    Call txtWNHS_Date_Approved_AfterUpdate
    
    Call txtSJOG_Date_Submitted_AfterUpdate
    Call txtSJOG_Date_Approved_AfterUpdate
    
    Call txtOthers_Date_Submitted_AfterUpdate
    Call txtOthers_Date_Approved_AfterUpdate
    
End Sub

Private Sub Userform_Terminate()
    'PURPOSE: Update register when unloaded
        
    If cbSaveonUnload.Value Then
        'Apply changes
        Call UpdateRegister
        DoEvents
    End If
    
    'Access version control
    Call LogLastAccess
    
End Sub

Private Sub UpdateRegister()
    'PURPOSE: Apply changes into Register table
    Dim ReadRow(1 To 19) As Variant
    
    'Turn off Settings to speed up
    'source: https://www.automateexcel.com/vba/turn-off-screen-updating/
    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayStatusBar = False
    Application.EnableEvents = False
    
    With RegTable.ListRows(RowIndex)
        
        'Populate ReadRow Array - faster than double transpose
        ReadRow(1) = String_to_Date(Me.txtCAHS_Date_Submitted.Value)
        ReadRow(2) = String_to_Date(Me.txtCAHS_Date_Responded.Value)
        ReadRow(3) = String_to_Date(Me.txtCAHS_Date_Resubmitted.Value)
        ReadRow(4) = String_to_Date(Me.txtCAHS_Date_Approved.Value)
        ReadRow(5) = Me.txtCAHS_Reminder.Value
        
        ReadRow(6) = Me.txtNMA_Committee.Value
        ReadRow(7) = String_to_Date(Me.txtNMA_Date_Submitted.Value)
        ReadRow(8) = String_to_Date(Me.txtNMA_Date_Approved.Value)
        ReadRow(9) = Me.txtNMA_Reminder.Value
        
        ReadRow(10) = String_to_Date(Me.txtWNHS_Date_Submitted.Value)
        ReadRow(11) = String_to_Date(Me.txtWNHS_Date_Approved.Value)
        ReadRow(12) = Me.txtWNHS_Reminder.Value
        
        ReadRow(13) = String_to_Date(Me.txtSJOG_Date_Submitted.Value)
        ReadRow(14) = String_to_Date(Me.txtSJOG_Date_Approved.Value)
        ReadRow(15) = Me.txtSJOG_Reminder.Value
        
        ReadRow(16) = Me.txtOthers_Committee.Value
        ReadRow(17) = String_to_Date(Me.txtOthers_Date_Submitted.Value)
        ReadRow(18) = String_to_Date(Me.txtOthers_Date_Approved.Value)
        ReadRow(19) = Me.txtOthers_Reminder.Value
        
        'Write to Register Table
        .Range(42) = ReadRow(1)
        .Range(43) = ReadRow(2)
        .Range(44) = ReadRow(3)
        .Range(45) = ReadRow(4)
        .Range(46) = ReadRow(5)
        
        .Range(47) = ReadRow(6)
        .Range(48) = ReadRow(7)
        .Range(49) = ReadRow(8)
        .Range(50) = ReadRow(9)
        
        .Range(51) = ReadRow(10)
        .Range(52) = ReadRow(11)
        .Range(53) = ReadRow(12)
        
        .Range(54) = ReadRow(13)
        .Range(55) = ReadRow(14)
        .Range(56) = ReadRow(15)
        
        .Range(57) = ReadRow(16)
        .Range(58) = ReadRow(17)
        .Range(59) = ReadRow(18)
        .Range(60) = ReadRow(19)
           
        'Store next old values
        NxtOldValues = ReadRow
        
        'Check if values changed
        If Not ArraysSame(ReadRow, OldValues) Then
            'Update version control
            .Range(61) = Now
            .Range(62) = Username
            
            'Apply completion status
             Call Fill_Completion_Status
             DoEvents
        End If
        
        'Clear array elements
        Erase ReadRow
        
    End With
    
    'Reinstate Settings
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayStatusBar = True
    Application.EnableEvents = True
    
End Sub

Private Sub Fill_Completion_Status()

    'PURPOSE: Evaluate entry completion status
    
    Dim db As Range
    Dim ReadRow As Variant
    Dim i As Integer, cntTrue As Integer, cntEmpty As Integer
    
    'Turn off Settings to speed up
    'source: https://www.automateexcel.com/vba/turn-off-screen-updating/
    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayStatusBar = False
    Application.EnableEvents = False
    
    'Exit if register is empty
    If RegTable.DataBodyRange Is Nothing Then
        GoTo ErrHandler
    End If
    
    Set db = RegTable.DataBodyRange
    
    'Tranpose twice to get 1D Array
    ReadRow = Application.Transpose(Application.Transpose(Range(db.Cells(RowIndex, 42), db.Cells(RowIndex, 59))))
                   
    'Apply correct test on each field
    For i = LBound(ReadRow) To UBound(ReadRow)
        If ReadRow(i) <> vbNullString Then
    
            Select Case Correct(i + 34)
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
    
    'CAHS Ethics
    'Criteria - all fields filled with dates
    cntTrue = 0
    cntEmpty = 0
    For i = 1 To 4
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
    For i = 6 To 8
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
    For i = 10 To 11
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
    For i = 13 To 14
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
    For i = 16 To 18
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
   
ErrHandler:
    'Clear array
    Erase ReadRow
    
    'Reinstate Settings
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayStatusBar = True
    Application.EnableEvents = True
    
End Sub

Private Sub cbSaveonUnload_Click()
    'PURPOSE: Change value of SAG_Tick variable
    SAG_Tick = Me.cbSaveonUnload.Value
    
    If SAG_Tick Then
        Me.cbSaveonUnload.Caption = "Save via Navigation"
    Else
        Me.cbSaveonUnload.Caption = "Save via Button"
    End If
    
End Sub


'----------------- Navigation section Toggles ----------------

Private Sub tglNav_Click()
    'PURPOSE: Closes current form and open Nav form
    Unload form06_Ethics
    
    form00_Nav.show False
End Sub

Private Sub tglCDA_Click()
    'PURPOSE: Closes current form and open CDA form
    Unload form06_Ethics
    
    form02_CDA.show False
End Sub

Private Sub tglFS_Click()
    'PURPOSE: Closes current form and open Feasibility form
    Unload form06_Ethics
    
    form03_FS.show False
End Sub

Private Sub tglSiteSelect_Click()
    'PURPOSE: Closes current form and open Site Select form
    Unload form06_Ethics
    
    form04_SiteSelect.show False
End Sub

Private Sub tglRecruit_Click()
    'PURPOSE: Closes current form and open Recruitment form
    Unload form06_Ethics
    
    form05_Recruitment.show False
End Sub

Private Sub tglStudyDetail_Click()
    'PURPOSE: Closes current form and open Study Detail form
    Unload form06_Ethics
    
    form01_StudyDetail.show False
End Sub

Private Sub tglGov_Click()
    'PURPOSE: Closes current form and open Governance form
    Unload form06_Ethics
    
    form07_Governance.show False
    form07_Governance.multiGov.Value = 0
End Sub

Private Sub tglBudget_Click()
    'PURPOSE: Closes current form and open Budget form
    Unload form06_Ethics
    
    form08_Budget.show False
    form08_Budget.multiBudget.Value = 0
End Sub

Private Sub tglIndemnity_Click()
    'PURPOSE: Closes current form and open Indemnity form
    Unload form06_Ethics
    
    form09_Indemnity.show False
End Sub

Private Sub tglCTRA_Click()
    'PURPOSE: Closes current form and open CTRA form
    Unload form06_Ethics
    
    form10_CTRA.show False
End Sub

Private Sub tglFinDisc_Click()
    'PURPOSE: Closes current form and open Fin. Disc. form
    Unload form06_Ethics
    
    form11_FinDisc.show False
End Sub

Private Sub tglSIV_Click()
    'PURPOSE: Closes current form and open SIV form
    Unload form06_Ethics
    
    form12_SIV.show False
End Sub

Private Function ArraysSame(ArrX As Variant, ArrY As Variant) As Boolean
    'PURPOSE: Compare values of two 1D arrays
    
    Dim Check As Boolean
    Dim Upper As Long, i As Long
    
    Check = True
    Upper = UBound(ArrX)
    
    'Shift upper bound to smaller array
    If UBound(ArrX) >= UBound(ArrY) Then
        Upper = UBound(ArrY)
    End If
    
    For i = LBound(ArrX) To Upper
        If ArrX(i) <> ArrY(i) Then
            Check = False
            Exit For
        End If
    Next i
    
    ArraysSame = Check
End Function
