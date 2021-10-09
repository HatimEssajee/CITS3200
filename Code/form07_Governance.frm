VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form07_Governance 
   Caption         =   "Governance Review"
   ClientHeight    =   6540
   ClientLeft      =   -552
   ClientTop       =   -2484
   ClientWidth     =   10584
   OleObjectBlob   =   "form07_Governance.frx":0000
End
Attribute VB_Name = "form07_Governance"
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
    ReDim OldValues(1 To 29)
    ReDim NxtOldValues(1 To 29)
                    
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
    
    For Each pPage In Me.multiGov.Pages
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
        
        Me.txtPCH_Date_Submitted.Value = ReadDate(.Range(63).Value)
        Me.txtPCH_Date_Responded.Value = ReadDate(.Range(64).Value)
        Me.txtPCH_Date_Approved.Value = ReadDate(.Range(65).Value)
        Me.txtPCH_Reminder.Value = .Range(66).Value
        
        Me.txtTKI_Date_Submitted.Value = ReadDate(.Range(67).Value)
        Me.txtTKI_Date_Responded.Value = ReadDate(.Range(68).Value)
        Me.txtTKI_Date_Approved.Value = ReadDate(.Range(69).Value)
        Me.txtTKI_Reminder.Value = .Range(70).Value
        
        Me.txtKEMH_Date_Submitted.Value = ReadDate(.Range(71).Value)
        Me.txtKEMH_Date_Responded.Value = ReadDate(.Range(72).Value)
        Me.txtKEMH_Date_Approved.Value = ReadDate(.Range(73).Value)
        Me.txtKEMH_Reminder.Value = .Range(74).Value
        
        Me.txtSJOG_S_Date_Submitted.Value = ReadDate(.Range(75).Value)
        Me.txtSJOG_S_Date_Responded.Value = ReadDate(.Range(76).Value)
        Me.txtSJOG_S_Date_Approved.Value = ReadDate(.Range(77).Value)
        Me.txtSJOG_S_Reminder.Value = .Range(78).Value
        
        Me.txtSJOG_L_Date_Submitted.Value = ReadDate(.Range(79).Value)
        Me.txtSJOG_L_Date_Responded.Value = ReadDate(.Range(80).Value)
        Me.txtSJOG_L_Date_Approved.Value = ReadDate(.Range(81).Value)
        Me.txtSJOG_L_Reminder.Value = .Range(82).Value
        
        Me.txtSJOG_M_Date_Submitted.Value = ReadDate(.Range(83).Value)
        Me.txtSJOG_M_Date_Responded.Value = ReadDate(.Range(84).Value)
        Me.txtSJOG_M_Date_Approved.Value = ReadDate(.Range(85).Value)
        Me.txtSJOG_M_Reminder.Value = .Range(86).Value
        
        Me.txtOthers_Committee.Value = .Range(87).Value
        Me.txtOthers_Date_Submitted.Value = ReadDate(.Range(88).Value)
        Me.txtOthers_Date_Responded.Value = ReadDate(.Range(89).Value)
        Me.txtOthers_Date_Approved.Value = ReadDate(.Range(90).Value)
        Me.txtOthers_Reminder.Value = .Range(91).Value
        
    End With
    
    'Populate Old Values Array - for undo
    OldValues(1) = String_to_Date(Me.txtPCH_Date_Submitted.Value)
    OldValues(2) = String_to_Date(Me.txtPCH_Date_Responded.Value)
    OldValues(3) = String_to_Date(Me.txtPCH_Date_Approved.Value)
    OldValues(4) = Me.txtPCH_Reminder.Value
    
    OldValues(5) = String_to_Date(Me.txtTKI_Date_Submitted.Value)
    OldValues(6) = String_to_Date(Me.txtTKI_Date_Responded.Value)
    OldValues(7) = String_to_Date(Me.txtTKI_Date_Approved.Value)
    OldValues(8) = Me.txtTKI_Reminder.Value
    
    OldValues(9) = String_to_Date(Me.txtKEMH_Date_Submitted.Value)
    OldValues(10) = String_to_Date(Me.txtKEMH_Date_Responded.Value)
    OldValues(11) = String_to_Date(Me.txtKEMH_Date_Approved.Value)
    OldValues(12) = Me.txtKEMH_Reminder.Value
    
    OldValues(13) = String_to_Date(Me.txtSJOG_S_Date_Submitted.Value)
    OldValues(14) = String_to_Date(Me.txtSJOG_S_Date_Responded.Value)
    OldValues(15) = String_to_Date(Me.txtSJOG_S_Date_Approved.Value)
    OldValues(16) = Me.txtSJOG_S_Reminder.Value
    
    OldValues(17) = String_to_Date(Me.txtSJOG_L_Date_Submitted.Value)
    OldValues(18) = String_to_Date(Me.txtSJOG_L_Date_Responded.Value)
    OldValues(19) = String_to_Date(Me.txtSJOG_L_Date_Approved.Value)
    OldValues(20) = Me.txtSJOG_L_Reminder.Value
    
    OldValues(21) = String_to_Date(Me.txtSJOG_M_Date_Submitted.Value)
    OldValues(22) = String_to_Date(Me.txtSJOG_M_Date_Responded.Value)
    OldValues(23) = String_to_Date(Me.txtSJOG_M_Date_Approved.Value)
    OldValues(24) = Me.txtSJOG_M_Reminder.Value

    OldValues(25) = Me.txtOthers_Committee.Value
    OldValues(26) = String_to_Date(Me.txtOthers_Date_Submitted.Value)
    OldValues(27) = String_to_Date(Me.txtOthers_Date_Responded.Value)
    OldValues(28) = String_to_Date(Me.txtOthers_Date_Approved.Value)
    OldValues(29) = Me.txtOthers_Reminder.Value
    
    'Initialize NxtOldValues
    NxtOldValues = OldValues
    
    'Access version control
    Call LogLastAccess
    
    'Depress and make toggle green on nav bar
    Me.tglGov.Value = True
    Me.tglGov.BackColor = vbGreen
    
    'Allocate tick box values
    Me.cbSaveonUnload.Value = SAG_Tick
    
    'Run date validation on data entered
    Call txtPCH_Date_Submitted_AfterUpdate
    Call txtPCH_Date_Responded_AfterUpdate
    Call txtPCH_Date_Approved_AfterUpdate
    
    Call txtTKI_Date_Submitted_AfterUpdate
    Call txtTKI_Date_Responded_AfterUpdate
    Call txtTKI_Date_Approved_AfterUpdate
    
    Call txtKEMH_Date_Submitted_AfterUpdate
    Call txtKEMH_Date_Responded_AfterUpdate
    Call txtKEMH_Date_Approved_AfterUpdate
    
    Call txtSJOG_S_Date_Submitted_AfterUpdate
    Call txtSJOG_S_Date_Responded_AfterUpdate
    Call txtSJOG_S_Date_Approved_AfterUpdate
    
    Call txtSJOG_L_Date_Submitted_AfterUpdate
    Call txtSJOG_L_Date_Responded_AfterUpdate
    Call txtSJOG_L_Date_Approved_AfterUpdate
    
    Call txtSJOG_M_Date_Submitted_AfterUpdate
    Call txtSJOG_M_Date_Responded_AfterUpdate
    Call txtSJOG_M_Date_Approved_AfterUpdate
    
    Call txtOthers_Date_Submitted_AfterUpdate
    Call txtOthers_Date_Responded_AfterUpdate
    Call txtOthers_Date_Approved_AfterUpdate
       
    'Reinstate Settings
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayStatusBar = True
    Application.EnableEvents = True
    
End Sub

Private Sub txtPCH_Date_Submitted_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtPCH_Date_Submitted.Value)
    
    'Display error message
    Me.errPCH_Date_Submitted.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtPCH_Date_Submitted.Value) Then
        Me.txtPCH_Date_Submitted.Value = Format(Me.txtPCH_Date_Submitted.Value, "dd-mmm-yyyy")
    End If
    
End Sub

Private Sub txtPCH_Date_Responded_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtPCH_Date_Responded.Value, Me.txtPCH_Date_Submitted.Value, _
            "Date entered earlier than date Submitted")

    'Display error message
    Me.errPCH_Date_Responded.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtPCH_Date_Responded.Value) Then
        Me.txtPCH_Date_Responded.Value = Format(Me.txtPCH_Date_Responded.Value, "dd-mmm-yyyy")
    End If
     
End Sub

Private Sub txtPCH_Date_Approved_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtPCH_Date_Approved.Value, Me.txtPCH_Date_Submitted.Value, _
            "Date entered earlier than date Submitted")

    'Display error message
    Me.errPCH_Date_Approved.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtPCH_Date_Approved.Value) Then
        Me.txtPCH_Date_Approved.Value = Format(Me.txtPCH_Date_Approved.Value, "dd-mmm-yyyy")
    End If
     
End Sub

Private Sub txtTKI_Date_Submitted_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtTKI_Date_Submitted.Value)
    
    'Display error message
    Me.errTKI_Date_Submitted.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtTKI_Date_Submitted.Value) Then
        Me.txtTKI_Date_Submitted.Value = Format(Me.txtTKI_Date_Submitted.Value, "dd-mmm-yyyy")
    End If
    
End Sub

Private Sub txtTKI_Date_Responded_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtTKI_Date_Responded.Value, Me.txtTKI_Date_Submitted.Value, _
            "Date entered earlier than date Submitted")

    'Display error message
    Me.errTKI_Date_Responded.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtTKI_Date_Responded.Value) Then
        Me.txtTKI_Date_Responded.Value = Format(Me.txtTKI_Date_Responded.Value, "dd-mmm-yyyy")
    End If
     
End Sub

Private Sub txtTKI_Date_Approved_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtTKI_Date_Approved.Value, Me.txtTKI_Date_Submitted.Value, _
            "Date entered earlier than date Submitted")

    'Display error message
    Me.errTKI_Date_Approved.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtTKI_Date_Approved.Value) Then
        Me.txtTKI_Date_Approved.Value = Format(Me.txtTKI_Date_Approved.Value, "dd-mmm-yyyy")
    End If
     
End Sub

Private Sub txtKEMH_Date_Submitted_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtKEMH_Date_Submitted.Value)
    
    'Display error message
    Me.errKEMH_Date_Submitted.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtKEMH_Date_Submitted.Value) Then
        Me.txtKEMH_Date_Submitted.Value = Format(Me.txtKEMH_Date_Submitted.Value, "dd-mmm-yyyy")
    End If
    
End Sub

Private Sub txtKEMH_Date_Responded_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtKEMH_Date_Responded.Value, Me.txtKEMH_Date_Submitted.Value, _
            "Date entered earlier than date Submitted")

    'Display error message
    Me.errKEMH_Date_Responded.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtKEMH_Date_Responded.Value) Then
        Me.txtKEMH_Date_Responded.Value = Format(Me.txtKEMH_Date_Responded.Value, "dd-mmm-yyyy")
    End If
     
End Sub

Private Sub txtKEMH_Date_Approved_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtKEMH_Date_Approved.Value, Me.txtKEMH_Date_Submitted.Value, _
            "Date entered earlier than date Submitted")

    'Display error message
    Me.errKEMH_Date_Approved.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtKEMH_Date_Approved.Value) Then
        Me.txtKEMH_Date_Approved.Value = Format(Me.txtKEMH_Date_Approved.Value, "dd-mmm-yyyy")
    End If
     
End Sub

Private Sub txtSJOG_S_Date_Submitted_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtSJOG_S_Date_Submitted.Value)
    
    'Display error message
    Me.errSJOG_S_Date_Submitted.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtSJOG_S_Date_Submitted.Value) Then
        Me.txtSJOG_S_Date_Submitted.Value = Format(Me.txtSJOG_S_Date_Submitted.Value, "dd-mmm-yyyy")
    End If
    
End Sub

Private Sub txtSJOG_S_Date_Responded_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtSJOG_S_Date_Responded.Value, Me.txtSJOG_S_Date_Submitted.Value, _
            "Date entered earlier than date Submitted")

    'Display error message
    Me.errSJOG_S_Date_Responded.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtSJOG_S_Date_Responded.Value) Then
        Me.txtSJOG_S_Date_Responded.Value = Format(Me.txtSJOG_S_Date_Responded.Value, "dd-mmm-yyyy")
    End If
     
End Sub

Private Sub txtSJOG_S_Date_Approved_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtSJOG_S_Date_Approved.Value, Me.txtSJOG_S_Date_Submitted.Value, _
            "Date entered earlier than date Submitted")

    'Display error message
    Me.errSJOG_S_Date_Approved.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtSJOG_S_Date_Approved.Value) Then
        Me.txtSJOG_S_Date_Approved.Value = Format(Me.txtSJOG_S_Date_Approved.Value, "dd-mmm-yyyy")
    End If
     
End Sub

Private Sub txtSJOG_L_Date_Submitted_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtSJOG_L_Date_Submitted.Value)
    
    'Display error message
    Me.errSJOG_L_Date_Submitted.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtSJOG_L_Date_Submitted.Value) Then
        Me.txtSJOG_L_Date_Submitted.Value = Format(Me.txtSJOG_L_Date_Submitted.Value, "dd-mmm-yyyy")
    End If
    
End Sub

Private Sub txtSJOG_L_Date_Responded_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtSJOG_L_Date_Responded.Value, Me.txtSJOG_L_Date_Submitted.Value, _
            "Date entered earlier than date Submitted")

    'Display error message
    Me.errSJOG_L_Date_Responded.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtSJOG_L_Date_Responded.Value) Then
        Me.txtSJOG_L_Date_Responded.Value = Format(Me.txtSJOG_L_Date_Responded.Value, "dd-mmm-yyyy")
    End If
     
End Sub

Private Sub txtSJOG_L_Date_Approved_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtSJOG_L_Date_Approved.Value, Me.txtSJOG_L_Date_Submitted.Value, _
            "Date entered earlier than date Submitted")

    'Display error message
    Me.errSJOG_L_Date_Approved.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtSJOG_L_Date_Approved.Value) Then
        Me.txtSJOG_L_Date_Approved.Value = Format(Me.txtSJOG_L_Date_Approved.Value, "dd-mmm-yyyy")
    End If
     
End Sub

Private Sub txtSJOG_M_Date_Submitted_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtSJOG_M_Date_Submitted.Value)
    
    'Display error message
    Me.errSJOG_M_Date_Submitted.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtSJOG_M_Date_Submitted.Value) Then
        Me.txtSJOG_M_Date_Submitted.Value = Format(Me.txtSJOG_M_Date_Submitted.Value, "dd-mmm-yyyy")
    End If
    
End Sub

Private Sub txtSJOG_M_Date_Responded_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtSJOG_M_Date_Responded.Value, Me.txtSJOG_M_Date_Submitted.Value, _
            "Date entered earlier than date Submitted")

    'Display error message
    Me.errSJOG_M_Date_Responded.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtSJOG_M_Date_Responded.Value) Then
        Me.txtSJOG_M_Date_Responded.Value = Format(Me.txtSJOG_M_Date_Responded.Value, "dd-mmm-yyyy")
    End If
     
End Sub

Private Sub txtSJOG_M_Date_Approved_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtSJOG_M_Date_Approved.Value, Me.txtSJOG_M_Date_Submitted.Value, _
            "Date entered earlier than date Submitted")

    'Display error message
    Me.errSJOG_M_Date_Approved.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtSJOG_M_Date_Approved.Value) Then
        Me.txtSJOG_M_Date_Approved.Value = Format(Me.txtSJOG_M_Date_Approved.Value, "dd-mmm-yyyy")
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

Private Sub txtOthers_Date_Responded_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtOthers_Date_Responded.Value, Me.txtOthers_Date_Submitted.Value, _
            "Date entered earlier than date Submitted")

    'Display error message
    Me.errOthers_Date_Responded.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtOthers_Date_Responded.Value) Then
        Me.txtOthers_Date_Responded.Value = Format(Me.txtOthers_Date_Responded.Value, "dd-mmm-yyyy")
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
    
    Me.txtPCH_Date_Submitted.Value = ReadDate(CStr(OldValues(1)))
    Me.txtPCH_Date_Responded.Value = ReadDate(CStr(OldValues(2)))
    Me.txtPCH_Date_Approved.Value = ReadDate(CStr(OldValues(3)))
    Me.txtPCH_Reminder.Value = OldValues(4)
    
    Me.txtTKI_Date_Submitted.Value = ReadDate(CStr(OldValues(5)))
    Me.txtTKI_Date_Responded.Value = ReadDate(CStr(OldValues(6)))
    Me.txtTKI_Date_Approved.Value = ReadDate(CStr(OldValues(7)))
    Me.txtTKI_Reminder.Value = OldValues(8)
    
    Me.txtKEMH_Date_Submitted.Value = ReadDate(CStr(OldValues(9)))
    Me.txtKEMH_Date_Responded.Value = ReadDate(CStr(OldValues(10)))
    Me.txtKEMH_Date_Approved.Value = ReadDate(CStr(OldValues(11)))
    Me.txtKEMH_Reminder.Value = OldValues(12)
    
    Me.txtSJOG_S_Date_Submitted.Value = ReadDate(CStr(OldValues(13)))
    Me.txtSJOG_S_Date_Responded.Value = ReadDate(CStr(OldValues(14)))
    Me.txtSJOG_S_Date_Approved.Value = ReadDate(CStr(OldValues(15)))
    Me.txtSJOG_S_Reminder.Value = OldValues(16)
    
    Me.txtSJOG_L_Date_Submitted.Value = ReadDate(CStr(OldValues(17)))
    Me.txtSJOG_L_Date_Responded.Value = ReadDate(CStr(OldValues(18)))
    Me.txtSJOG_L_Date_Approved.Value = ReadDate(CStr(OldValues(19)))
    Me.txtSJOG_L_Reminder.Value = OldValues(20)
    
    Me.txtSJOG_M_Date_Submitted.Value = ReadDate(CStr(OldValues(21)))
    Me.txtSJOG_M_Date_Responded.Value = ReadDate(CStr(OldValues(22)))
    Me.txtSJOG_M_Date_Approved.Value = ReadDate(CStr(OldValues(23)))
    Me.txtSJOG_M_Reminder.Value = OldValues(24)
    
    Me.txtOthers_Committee.Value = OldValues(25)
    Me.txtOthers_Date_Submitted.Value = ReadDate(CStr(OldValues(26)))
    Me.txtOthers_Date_Responded.Value = ReadDate(CStr(OldValues(27)))
    Me.txtOthers_Date_Approved.Value = ReadDate(CStr(OldValues(28)))
    Me.txtOthers_Reminder.Value = OldValues(29)
    
End Sub

Private Sub cmdRedo_Click()
    'PURPOSE: Recall values replaced by undo
    
    Me.txtPCH_Date_Submitted.Value = ReadDate(CStr(NxtOldValues(1)))
    Me.txtPCH_Date_Responded.Value = ReadDate(CStr(NxtOldValues(2)))
    Me.txtPCH_Date_Approved.Value = ReadDate(CStr(NxtOldValues(3)))
    Me.txtPCH_Reminder.Value = NxtOldValues(4)
    
    Me.txtTKI_Date_Submitted.Value = ReadDate(CStr(NxtOldValues(5)))
    Me.txtTKI_Date_Responded.Value = ReadDate(CStr(NxtOldValues(6)))
    Me.txtTKI_Date_Approved.Value = ReadDate(CStr(NxtOldValues(7)))
    Me.txtTKI_Reminder.Value = NxtOldValues(8)
    
    Me.txtKEMH_Date_Submitted.Value = ReadDate(CStr(NxtOldValues(9)))
    Me.txtKEMH_Date_Responded.Value = ReadDate(CStr(NxtOldValues(10)))
    Me.txtKEMH_Date_Approved.Value = ReadDate(CStr(NxtOldValues(11)))
    Me.txtKEMH_Reminder.Value = NxtOldValues(12)
    
    Me.txtSJOG_S_Date_Submitted.Value = ReadDate(CStr(NxtOldValues(13)))
    Me.txtSJOG_S_Date_Responded.Value = ReadDate(CStr(NxtOldValues(14)))
    Me.txtSJOG_S_Date_Approved.Value = ReadDate(CStr(NxtOldValues(15)))
    Me.txtSJOG_S_Reminder.Value = NxtOldValues(16)
    
    Me.txtSJOG_L_Date_Submitted.Value = ReadDate(CStr(NxtOldValues(17)))
    Me.txtSJOG_L_Date_Responded.Value = ReadDate(CStr(NxtOldValues(18)))
    Me.txtSJOG_L_Date_Approved.Value = ReadDate(CStr(NxtOldValues(19)))
    Me.txtSJOG_L_Reminder.Value = NxtOldValues(20)
    
    Me.txtSJOG_M_Date_Submitted.Value = ReadDate(CStr(NxtOldValues(21)))
    Me.txtSJOG_M_Date_Responded.Value = ReadDate(CStr(NxtOldValues(22)))
    Me.txtSJOG_M_Date_Approved.Value = ReadDate(CStr(NxtOldValues(23)))
    Me.txtSJOG_M_Reminder.Value = NxtOldValues(24)
    
    Me.txtOthers_Committee.Value = NxtOldValues(25)
    Me.txtOthers_Date_Submitted.Value = ReadDate(CStr(NxtOldValues(26)))
    Me.txtOthers_Date_Responded.Value = ReadDate(CStr(NxtOldValues(27)))
    Me.txtOthers_Date_Approved.Value = ReadDate(CStr(NxtOldValues(28)))
    Me.txtOthers_Reminder.Value = NxtOldValues(29)
    
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
    Call txtPCH_Date_Submitted_AfterUpdate
    Call txtPCH_Date_Responded_AfterUpdate
    Call txtPCH_Date_Approved_AfterUpdate
    
    Call txtTKI_Date_Submitted_AfterUpdate
    Call txtTKI_Date_Responded_AfterUpdate
    Call txtTKI_Date_Approved_AfterUpdate
    
    Call txtKEMH_Date_Submitted_AfterUpdate
    Call txtKEMH_Date_Responded_AfterUpdate
    Call txtKEMH_Date_Approved_AfterUpdate
    
    Call txtSJOG_S_Date_Submitted_AfterUpdate
    Call txtSJOG_S_Date_Responded_AfterUpdate
    Call txtSJOG_S_Date_Approved_AfterUpdate
    
    Call txtSJOG_L_Date_Submitted_AfterUpdate
    Call txtSJOG_L_Date_Responded_AfterUpdate
    Call txtSJOG_L_Date_Approved_AfterUpdate
    
    Call txtSJOG_M_Date_Submitted_AfterUpdate
    Call txtSJOG_M_Date_Responded_AfterUpdate
    Call txtSJOG_M_Date_Approved_AfterUpdate
    
    Call txtOthers_Date_Submitted_AfterUpdate
    Call txtOthers_Date_Responded_AfterUpdate
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
    Dim ReadRow(1 To 29) As Variant
    
    'Turn off Settings to speed up
    'source: https://www.automateexcel.com/vba/turn-off-screen-updating/
    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayStatusBar = False
    Application.EnableEvents = False
    
    With RegTable.ListRows(RowIndex)
        
        'Populate ReadRow Array - faster than double transpose
        ReadRow(1) = String_to_Date(Me.txtPCH_Date_Submitted.Value)
        ReadRow(2) = String_to_Date(Me.txtPCH_Date_Responded.Value)
        ReadRow(3) = String_to_Date(Me.txtPCH_Date_Approved.Value)
        ReadRow(4) = Me.txtPCH_Reminder.Value
        
        ReadRow(5) = String_to_Date(Me.txtTKI_Date_Submitted.Value)
        ReadRow(6) = String_to_Date(Me.txtTKI_Date_Responded.Value)
        ReadRow(7) = String_to_Date(Me.txtTKI_Date_Approved.Value)
        ReadRow(8) = Me.txtTKI_Reminder.Value
        
        ReadRow(9) = String_to_Date(Me.txtKEMH_Date_Submitted.Value)
        ReadRow(10) = String_to_Date(Me.txtKEMH_Date_Responded.Value)
        ReadRow(11) = String_to_Date(Me.txtKEMH_Date_Approved.Value)
        ReadRow(12) = Me.txtKEMH_Reminder.Value
        
        ReadRow(13) = String_to_Date(Me.txtSJOG_S_Date_Submitted.Value)
        ReadRow(14) = String_to_Date(Me.txtSJOG_S_Date_Responded.Value)
        ReadRow(15) = String_to_Date(Me.txtSJOG_S_Date_Approved.Value)
        ReadRow(16) = Me.txtSJOG_S_Reminder.Value
        
        ReadRow(17) = String_to_Date(Me.txtSJOG_L_Date_Submitted.Value)
        ReadRow(18) = String_to_Date(Me.txtSJOG_L_Date_Responded.Value)
        ReadRow(19) = String_to_Date(Me.txtSJOG_L_Date_Approved.Value)
        ReadRow(20) = Me.txtSJOG_L_Reminder.Value
        
        ReadRow(21) = String_to_Date(Me.txtSJOG_M_Date_Submitted.Value)
        ReadRow(22) = String_to_Date(Me.txtSJOG_M_Date_Responded.Value)
        ReadRow(23) = String_to_Date(Me.txtSJOG_M_Date_Approved.Value)
        ReadRow(24) = Me.txtSJOG_M_Reminder.Value
    
        ReadRow(25) = Me.txtOthers_Committee.Value
        ReadRow(26) = String_to_Date(Me.txtOthers_Date_Submitted.Value)
        ReadRow(27) = String_to_Date(Me.txtOthers_Date_Responded.Value)
        ReadRow(28) = String_to_Date(Me.txtOthers_Date_Approved.Value)
        ReadRow(29) = Me.txtOthers_Reminder.Value
        
        'Write to Register Table
        .Range(63) = ReadRow(1)
        .Range(64) = ReadRow(2)
        .Range(65) = ReadRow(3)
        .Range(66) = WriteText(ReadRow(4))
        
        .Range(67) = ReadRow(5)
        .Range(68) = ReadRow(6)
        .Range(69) = ReadRow(7)
        .Range(70) = WriteText(ReadRow(8))
        
        .Range(71) = ReadRow(9)
        .Range(72) = ReadRow(10)
        .Range(73) = ReadRow(11)
        .Range(74) = WriteText(ReadRow(12))
        
        .Range(75) = ReadRow(13)
        .Range(76) = ReadRow(14)
        .Range(77) = ReadRow(15)
        .Range(78) = WriteText(ReadRow(16))
        
        .Range(79) = ReadRow(17)
        .Range(80) = ReadRow(18)
        .Range(81) = ReadRow(19)
        .Range(82) = WriteText(ReadRow(20))
        
        .Range(83) = ReadRow(21)
        .Range(84) = ReadRow(22)
        .Range(85) = ReadRow(23)
        .Range(86) = WriteText(ReadRow(24))
    
        .Range(87) = WriteText(ReadRow(25))
        .Range(88) = ReadRow(26)
        .Range(89) = ReadRow(27)
        .Range(90) = ReadRow(28)
        .Range(91) = WriteText(ReadRow(29))
        
        'Store next old values
        NxtOldValues = ReadRow
        
        'Check if values changed
        If Not ArraysSame(ReadRow, OldValues) Then
        
            'Update version control
            .Range(92) = Now
            .Range(93) = Username
            
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
    Dim ReadRow(1 To 28) As Variant
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
    For i = 1 To 28
        ReadRow(i) = db.Cells(RowIndex, 62 + i).Value
    Next i
    'ReadRow = Application.Transpose(Application.Transpose(Range(db.Cells(RowIndex, 63), db.Cells(RowIndex, 90))))
                   
    'Apply correct test on each field
    For i = LBound(ReadRow) To UBound(ReadRow)
        If ReadRow(i) <> vbNullString Then
    
            Select Case Correct(i + 55)
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
    
    'PCH Governance
    'Criteria - all fields filled with dates
    cntTrue = 0
    cntEmpty = 0
    For i = 1 To 3
        If ReadRow(i) = vbNullString Then
            cntEmpty = cntEmpty + 1
        ElseIf ReadRow(i) Then
            cntTrue = cntTrue + 1
        End If
    Next i
    
    If cntEmpty = 3 And db.Cells(RowIndex, 66).Value <> vbNullString Then
        db.Cells(RowIndex, 139) = False
    ElseIf cntEmpty = 3 Then
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
    For i = 5 To 7
        If ReadRow(i) = vbNullString Then
            cntEmpty = cntEmpty + 1
        ElseIf ReadRow(i) Then
            cntTrue = cntTrue + 1
        End If
    Next i
    
    If cntEmpty = 3 And db.Cells(RowIndex, 70).Value <> vbNullString Then
        db.Cells(RowIndex, 140) = False
    ElseIf cntEmpty = 3 Then
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
    For i = 9 To 11
        If ReadRow(i) = vbNullString Then
            cntEmpty = cntEmpty + 1
        ElseIf ReadRow(i) Then
            cntTrue = cntTrue + 1
        End If
    Next i
    
    If cntEmpty = 3 And db.Cells(RowIndex, 74).Value <> vbNullString Then
        db.Cells(RowIndex, 141) = False
    ElseIf cntEmpty = 3 Then
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
    For i = 13 To 15
        If ReadRow(i) = vbNullString Then
            cntEmpty = cntEmpty + 1
        ElseIf ReadRow(i) Then
            cntTrue = cntTrue + 1
        End If
    Next i
    
    If cntEmpty = 3 And db.Cells(RowIndex, 78).Value <> vbNullString Then
        db.Cells(RowIndex, 142) = False
    ElseIf cntEmpty = 3 Then
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
    For i = 17 To 19
        If ReadRow(i) = vbNullString Then
            cntEmpty = cntEmpty + 1
        ElseIf ReadRow(i) Then
            cntTrue = cntTrue + 1
        End If
    Next i
    
    If cntEmpty = 3 And db.Cells(RowIndex, 82).Value <> vbNullString Then
        db.Cells(RowIndex, 143) = False
    ElseIf cntEmpty = 3 Then
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
    For i = 21 To 23
        If ReadRow(i) = vbNullString Then
            cntEmpty = cntEmpty + 1
        ElseIf ReadRow(i) Then
            cntTrue = cntTrue + 1
        End If
    Next i
    
    If cntEmpty = 3 And db.Cells(RowIndex, 86).Value <> vbNullString Then
        db.Cells(RowIndex, 144) = False
    ElseIf cntEmpty = 3 Then
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
    For i = 25 To 28
        If ReadRow(i) = vbNullString Then
            cntEmpty = cntEmpty + 1
        ElseIf ReadRow(i) Then
            cntTrue = cntTrue + 1
        End If
    Next i
    
    If cntEmpty = 4 And db.Cells(RowIndex, 91).Value <> vbNullString Then
        db.Cells(RowIndex, 145) = False
    ElseIf cntEmpty = 4 Then
        db.Cells(RowIndex, 145) = vbNullString
    ElseIf cntTrue = 4 Then
        db.Cells(RowIndex, 145) = True
    Else
        db.Cells(RowIndex, 145) = False
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
    Unload form07_Governance
    
    form00_Nav.show False
End Sub

Private Sub tglCDA_Click()
    'PURPOSE: Closes current form and open CDA form
    Unload form07_Governance
    
    form02_CDA.show False
End Sub

Private Sub tglFS_Click()
    'PURPOSE: Closes current form and open Feasibility form
    Unload form07_Governance
    
    form03_FS.show False
End Sub

Private Sub tglSiteSelect_Click()
    'PURPOSE: Closes current form and open Site Select form
    Unload form07_Governance
    
    form04_SiteSelect.show False
End Sub

Private Sub tglRecruit_Click()
    'PURPOSE: Closes current form and open Recruitment form
    Unload form07_Governance
    
    form05_Recruitment.show False
End Sub

Private Sub tglEthics_Click()
    'PURPOSE: Closes current form and open Ethics form
    Unload form07_Governance
    
    form06_Ethics.show False
    form06_Ethics.multiEthics.Value = 0
End Sub

Private Sub tglStudyDetail_Click()
    'PURPOSE: Closes current form and open Study Detail form
    Unload form07_Governance
    
    form01_StudyDetail.show False
End Sub

Private Sub tglBudget_Click()
    'PURPOSE: Closes current form and open Budget form
    Unload form07_Governance
    
    form08_Budget.show False
    form08_Budget.multiBudget.Value = 0
End Sub

Private Sub tglIndemnity_Click()
    'PURPOSE: Closes current form and open Indemnity form
    Unload form07_Governance
    
    form09_Indemnity.show False
End Sub

Private Sub tglCTRA_Click()
    'PURPOSE: Closes current form and open CTRA form
    Unload form07_Governance
    
    form10_CTRA.show False
End Sub

Private Sub tglFinDisc_Click()
    'PURPOSE: Closes current form and open Fin. Disc. form
    Unload form07_Governance
    
    form11_FinDisc.show False
End Sub

Private Sub tglSIV_Click()
    'PURPOSE: Closes current form and open SIV form
    Unload form07_Governance
    
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
