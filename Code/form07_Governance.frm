VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form07_Governance 
   Caption         =   "Governance Review"
   ClientHeight    =   10236
   ClientLeft      =   -516
   ClientTop       =   -2316
   ClientWidth     =   13536
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
       
    'Clear user form
    'SOURCE: https://www.mrexcel.com/board/threads/loop-through-controls-on-a-userform.427103/
    For Each ctrl In Me.Controls
        Select Case True
                Case TypeOf ctrl Is MSForms.CheckBox
                    ctrl.Value = False
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
        
        Me.txtPCH_Date_Submitted.Value = Format(.Range(63).Value, "dd-mmm-yyyy")
        Me.txtPCH_Date_Responded.Value = Format(.Range(64).Value, "dd-mmm-yyyy")
        Me.txtPCH_Date_Approved.Value = Format(.Range(65).Value, "dd-mmm-yyyy")
        Me.txtPCH_Reminder.Value = .Range(66).Value
        
        Me.txtTKI_Date_Submitted.Value = Format(.Range(67).Value, "dd-mmm-yyyy")
        Me.txtTKI_Date_Responded.Value = Format(.Range(68).Value, "dd-mmm-yyyy")
        Me.txtTKI_Date_Approved.Value = Format(.Range(69).Value, "dd-mmm-yyyy")
        Me.txtTKI_Reminder.Value = .Range(70).Value
        
        Me.txtKEMH_Date_Submitted.Value = Format(.Range(71).Value, "dd-mmm-yyyy")
        Me.txtKEMH_Date_Responded.Value = Format(.Range(72).Value, "dd-mmm-yyyy")
        Me.txtKEMH_Date_Approved.Value = Format(.Range(73).Value, "dd-mmm-yyyy")
        Me.txtKEMH_Reminder.Value = .Range(74).Value
        
        Me.txtSJOG_S_Date_Submitted.Value = Format(.Range(75).Value, "dd-mmm-yyyy")
        Me.txtSJOG_S_Date_Responded.Value = Format(.Range(76).Value, "dd-mmm-yyyy")
        Me.txtSJOG_S_Date_Approved.Value = Format(.Range(77).Value, "dd-mmm-yyyy")
        Me.txtSJOG_S_Reminder.Value = .Range(78).Value
        
        Me.txtSJOG_L_Date_Submitted.Value = Format(.Range(79).Value, "dd-mmm-yyyy")
        Me.txtSJOG_L_Date_Responded.Value = Format(.Range(80).Value, "dd-mmm-yyyy")
        Me.txtSJOG_L_Date_Approved.Value = Format(.Range(81).Value, "dd-mmm-yyyy")
        Me.txtSJOG_L_Reminder.Value = .Range(82).Value
        
        Me.txtSJOG_M_Date_Submitted.Value = Format(.Range(83).Value, "dd-mmm-yyyy")
        Me.txtSJOG_M_Date_Responded.Value = Format(.Range(84).Value, "dd-mmm-yyyy")
        Me.txtSJOG_M_Date_Approved.Value = Format(.Range(85).Value, "dd-mmm-yyyy")
        Me.txtSJOG_M_Reminder.Value = .Range(86).Value
        
        Me.txtOthers_Committee.Value = .Range(87).Value
        Me.txtOthers_Date_Submitted.Value = Format(.Range(88).Value, "dd-mmm-yyyy")
        Me.txtOthers_Date_Responded.Value = Format(.Range(89).Value, "dd-mmm-yyyy")
        Me.txtOthers_Date_Approved.Value = Format(.Range(90).Value, "dd-mmm-yyyy")
        Me.txtOthers_Reminder.Value = .Range(91).Value
        
    End With
    
    'Access version control
    Call LogLastAccess
    
    'Depress and make toggle green on nav bar
    Me.tglGov.Value = True
    Me.tglGov.BackColor = vbGreen
    
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

Private Sub cmdClose_Click()
    'PURPOSE: Closes current form
    
    'Access version control
    Call LogLastAccess
    
    Unload Me
    
End Sub

Private Sub cmdEdit_Click()
    'PURPOSE: Apply changes into Register table
    With RegTable.ListRows(RowIndex)
        
        .Range(63) = String_to_Date(Me.txtPCH_Date_Submitted.Value)
        .Range(64) = String_to_Date(Me.txtPCH_Date_Responded.Value)
        .Range(65) = String_to_Date(Me.txtPCH_Date_Approved.Value)
        .Range(66) = Me.txtPCH_Reminder.Value
        
        .Range(67) = String_to_Date(Me.txtTKI_Date_Submitted.Value)
        .Range(68) = String_to_Date(Me.txtTKI_Date_Responded.Value)
        .Range(69) = String_to_Date(Me.txtTKI_Date_Approved.Value)
        .Range(70) = Me.txtTKI_Reminder.Value
        
        .Range(71) = String_to_Date(Me.txtKEMH_Date_Submitted.Value)
        .Range(72) = String_to_Date(Me.txtKEMH_Date_Responded.Value)
        .Range(73) = String_to_Date(Me.txtKEMH_Date_Approved.Value)
        .Range(74) = Me.txtKEMH_Reminder.Value
        
        .Range(75) = String_to_Date(Me.txtSJOG_S_Date_Submitted.Value)
        .Range(76) = String_to_Date(Me.txtSJOG_S_Date_Responded.Value)
        .Range(77) = String_to_Date(Me.txtSJOG_S_Date_Approved.Value)
        .Range(78) = Me.txtSJOG_S_Reminder.Value
        
        .Range(79) = String_to_Date(Me.txtSJOG_L_Date_Submitted.Value)
        .Range(80) = String_to_Date(Me.txtSJOG_L_Date_Responded.Value)
        .Range(81) = String_to_Date(Me.txtSJOG_L_Date_Approved.Value)
        .Range(82) = Me.txtSJOG_L_Reminder.Value
        
        .Range(83) = String_to_Date(Me.txtSJOG_M_Date_Submitted.Value)
        .Range(84) = String_to_Date(Me.txtSJOG_M_Date_Responded.Value)
        .Range(85) = String_to_Date(Me.txtSJOG_M_Date_Approved.Value)
        .Range(86) = Me.txtSJOG_M_Reminder.Value
    
        .Range(87) = Me.txtOthers_Committee.Value
        .Range(88) = String_to_Date(Me.txtOthers_Date_Submitted.Value)
        .Range(89) = String_to_Date(Me.txtOthers_Date_Responded.Value)
        .Range(90) = String_to_Date(Me.txtOthers_Date_Approved.Value)
        .Range(91) = Me.txtOthers_Reminder.Value
        
        'Apply completion status
        'PCH Governance
        If .Range(65).Value = vbNullString Then
            .Range(139).Value = vbNullString
        ElseIf IsDate(.Range(65).Value) Then
            .Range(139).Value = True
        End If
        
        'TKI Governance
        If .Range(69).Value = vbNullString Then
            .Range(140).Value = vbNullString
        ElseIf IsDate(.Range(69).Value) Then
            .Range(140).Value = True
        End If
        
        'KEMH Governance
        If .Range(73).Value = vbNullString Then
            .Range(141).Value = vbNullString
        ElseIf IsDate(.Range(73).Value) Then
            .Range(141).Value = True
        End If
        
        'SJOG Subiaco Governance
        If .Range(77).Value = vbNullString Then
            .Range(142).Value = vbNullString
        ElseIf IsDate(.Range(77).Value) Then
            .Range(142).Value = True
        End If
        
        'SJOG Mt Lawley Governance
        If .Range(81).Value = vbNullString Then
            .Range(143).Value = vbNullString
        ElseIf IsDate(.Range(81).Value) Then
            .Range(143).Value = True
        End If
        
        'SJOG Murdoch Governance
        If .Range(85).Value = vbNullString Then
            .Range(144).Value = vbNullString
        ElseIf IsDate(.Range(85).Value) Then
            .Range(144).Value = True
        End If
        
        'Others Governance
        If .Range(87).Value = vbNullString Or .Range(90).Value = vbNullString Then
            .Range(145).Value = vbNullString
        ElseIf IsDate(.Range(90).Value) Then
            .Range(145).Value = True
        End If

        'Update version control
        .Range(92) = Now
        .Range(93) = Username
        
    End With
    
    'Access version control
    Call LogLastAccess
    
    Call UserForm_Initialize

End Sub


'----------------- Navigation section Toggles ----------------

Private Sub tglNav_Click()
    'PURPOSE: Closes current form and open Nav form
    Unload form07_Governance
    
    form00_Nav.Show False
End Sub

Private Sub tglCDA_Click()
    'PURPOSE: Closes current form and open CDA form
    Unload form07_Governance
    
    form02_CDA.Show False
End Sub

Private Sub tglFS_Click()
    'PURPOSE: Closes current form and open Feasibility form
    Unload form07_Governance
    
    form03_FS.Show False
End Sub

Private Sub tglSiteSelect_Click()
    'PURPOSE: Closes current form and open Site Select form
    Unload form07_Governance
    
    form04_SiteSelect.Show False
End Sub

Private Sub tglRecruit_Click()
    'PURPOSE: Closes current form and open Recruitment form
    Unload form07_Governance
    
    form05_Recruitment.Show False
End Sub

Private Sub tglEthics_Click()
    'PURPOSE: Closes current form and open Ethics form
    Unload form07_Governance
    
    form06_Ethics.Show False
End Sub

Private Sub tglStudyDetail_Click()
    'PURPOSE: Closes current form and open Study Detail form
    Unload form07_Governance
    
    form01_StudyDetail.Show False
End Sub

Private Sub tglBudget_Click()
    'PURPOSE: Closes current form and open Budget form
    Unload form07_Governance
    
    form08_Budget.Show False
End Sub

Private Sub tglIndemnity_Click()
    'PURPOSE: Closes current form and open Indemnity form
    Unload form07_Governance
    
    form09_Indemnity.Show False
End Sub

Private Sub tglCTRA_Click()
    'PURPOSE: Closes current form and open CTRA form
    Unload form07_Governance
    
    form10_CTRA.Show False
End Sub

Private Sub tglFinDisc_Click()
    'PURPOSE: Closes current form and open Fin. Disc. form
    Unload form07_Governance
    
    form11_FinDisc.Show False
End Sub

Private Sub tglSIV_Click()
    'PURPOSE: Closes current form and open SIV form
    Unload form07_Governance
    
    form12_SIV.Show False
End Sub
