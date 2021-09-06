VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form043_Governance 
   Caption         =   "Governance Review"
   ClientHeight    =   4872
   ClientLeft      =   -504
   ClientTop       =   -2328
   ClientWidth     =   5808
   OleObjectBlob   =   "form043_Governance.frx":0000
End
Attribute VB_Name = "form043_Governance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Activate()
    'PURPOSE: Reposition userform to Top Left of application Window and fix size
    'source: https://www.mrexcel.com/board/threads/userform-startup-position.671108/
    'Me.StartUpPosition = 0
    'Me.Top = Application.Top + 25
    'Me.Left = Application.Left + 25
    Me.Top = UserFormTopPos
    Me.Left = UserFormLeftPos
    Me.Height = UHeight
    Me.Width = UWidth

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
                    ctrl.value = False
                Case TypeOf ctrl Is MSForms.TextBox
                    ctrl.value = ""
                Case TypeOf ctrl Is MSForms.Label
                    'Empty error captions
                    If Left(ctrl.Name, 3) = "err" Then
                        ctrl.Caption = ""
                    End If
                Case TypeOf ctrl Is MSForms.ComboBox
                    ctrl.value = ""
                    ctrl.Clear
                Case TypeOf ctrl Is MSForms.ListBox
                    ctrl.value = ""
                    ctrl.Clear
            End Select
    Next ctrl
    
    For Each pPage In Me.multiGov.Pages
        For Each ctrl In pPage.Controls
            Select Case True
                Case TypeOf ctrl Is MSForms.CheckBox
                    ctrl.value = False
                Case TypeOf ctrl Is MSForms.TextBox
                    ctrl.value = ""
                Case TypeOf ctrl Is MSForms.ComboBox
                    ctrl.value = ""
                Case TypeOf ctrl Is MSForms.ListBox
                    ctrl.value = ""
            End Select
                
        Next ctrl
    Next pPage
    
    'Read information from register table
    With RegTable.ListRows(RowIndex)

        Me.txtStudyName.value = .Range(10).value
        
        Me.txtPCH_Date_Submitted.value = Format(.Range(58).value, "dd-mmm-yyyy")
        Me.txtPCH_Date_Responded.value = Format(.Range(59).value, "dd-mmm-yyyy")
        Me.txtPCH_Date_Approved.value = Format(.Range(60).value, "dd-mmm-yyyy")
        
        Me.txtTKI_Date_Submitted.value = Format(.Range(61).value, "dd-mmm-yyyy")
        Me.txtTKI_Date_Responded.value = Format(.Range(62).value, "dd-mmm-yyyy")
        Me.txtTKI_Date_Approved.value = Format(.Range(63).value, "dd-mmm-yyyy")
        
        Me.txtKEMH_Date_Submitted.value = Format(.Range(64).value, "dd-mmm-yyyy")
        Me.txtKEMH_Date_Responded.value = Format(.Range(65).value, "dd-mmm-yyyy")
        Me.txtKEMH_Date_Approved.value = Format(.Range(66).value, "dd-mmm-yyyy")
        
        Me.txtSJOG_S_Date_Submitted.value = Format(.Range(67).value, "dd-mmm-yyyy")
        Me.txtSJOG_S_Date_Responded.value = Format(.Range(68).value, "dd-mmm-yyyy")
        Me.txtSJOG_S_Date_Approved.value = Format(.Range(69).value, "dd-mmm-yyyy")
        
        Me.txtSJOG_L_Date_Submitted.value = Format(.Range(70).value, "dd-mmm-yyyy")
        Me.txtSJOG_L_Date_Responded.value = Format(.Range(71).value, "dd-mmm-yyyy")
        Me.txtSJOG_L_Date_Approved.value = Format(.Range(72).value, "dd-mmm-yyyy")
        
        Me.txtSJOG_M_Date_Submitted.value = Format(.Range(73).value, "dd-mmm-yyyy")
        Me.txtSJOG_M_Date_Responded.value = Format(.Range(74).value, "dd-mmm-yyyy")
        Me.txtSJOG_M_Date_Approved.value = Format(.Range(75).value, "dd-mmm-yyyy")
        
        Me.txtOthers_Committee.value = .Range(76).value
        Me.txtOthers_Date_Submitted.value = Format(.Range(77).value, "dd-mmm-yyyy")
        Me.txtOthers_Date_Responded.value = Format(.Range(78).value, "dd-mmm-yyyy")
        Me.txtOthers_Date_Approved.value = Format(.Range(79).value, "dd-mmm-yyyy")
        
        Me.txtReminder.value = .Range(80).value
        
    End With
    
    'Access version control
    Call LogLastAccess
    
    'Depress and make toggle green on nav bar
    Me.tglReviews.value = True
    Me.tglReviews.BackColor = vbGreen
    Me.tglGovernance.value = True
    Me.tglGovernance.BackColor = vbGreen
    
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
    
    err = Date_Validation(Me.txtPCH_Date_Submitted.value)
    
    'Display error message
    Me.errPCH_Date_Submitted.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtPCH_Date_Submitted.value) Then
        Me.txtPCH_Date_Submitted.value = Format(Me.txtPCH_Date_Submitted.value, "dd-mmm-yyyy")
    End If
    
End Sub

Private Sub txtPCH_Date_Responded_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtPCH_Date_Responded.value, Me.txtPCH_Date_Submitted.value, _
            "Date entered earlier than date Submitted")

    'Display error message
    Me.errPCH_Date_Responded.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtPCH_Date_Responded.value) Then
        Me.txtPCH_Date_Responded.value = Format(Me.txtPCH_Date_Responded.value, "dd-mmm-yyyy")
    End If
     
End Sub

Private Sub txtPCH_Date_Approved_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtPCH_Date_Approved.value, Me.txtPCH_Date_Submitted.value, _
            "Date entered earlier than date Submitted")

    'Display error message
    Me.errPCH_Date_Approved.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtPCH_Date_Approved.value) Then
        Me.txtPCH_Date_Approved.value = Format(Me.txtPCH_Date_Approved.value, "dd-mmm-yyyy")
    End If
     
End Sub

Private Sub txtTKI_Date_Submitted_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtTKI_Date_Submitted.value)
    
    'Display error message
    Me.errTKI_Date_Submitted.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtTKI_Date_Submitted.value) Then
        Me.txtTKI_Date_Submitted.value = Format(Me.txtTKI_Date_Submitted.value, "dd-mmm-yyyy")
    End If
    
End Sub

Private Sub txtTKI_Date_Responded_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtTKI_Date_Responded.value, Me.txtTKI_Date_Submitted.value, _
            "Date entered earlier than date Submitted")

    'Display error message
    Me.errTKI_Date_Responded.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtTKI_Date_Responded.value) Then
        Me.txtTKI_Date_Responded.value = Format(Me.txtTKI_Date_Responded.value, "dd-mmm-yyyy")
    End If
     
End Sub

Private Sub txtTKI_Date_Approved_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtTKI_Date_Approved.value, Me.txtTKI_Date_Submitted.value, _
            "Date entered earlier than date Submitted")

    'Display error message
    Me.errTKI_Date_Approved.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtTKI_Date_Approved.value) Then
        Me.txtTKI_Date_Approved.value = Format(Me.txtTKI_Date_Approved.value, "dd-mmm-yyyy")
    End If
     
End Sub

Private Sub txtKEMH_Date_Submitted_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtKEMH_Date_Submitted.value)
    
    'Display error message
    Me.errKEMH_Date_Submitted.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtKEMH_Date_Submitted.value) Then
        Me.txtKEMH_Date_Submitted.value = Format(Me.txtKEMH_Date_Submitted.value, "dd-mmm-yyyy")
    End If
    
End Sub

Private Sub txtKEMH_Date_Responded_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtKEMH_Date_Responded.value, Me.txtKEMH_Date_Submitted.value, _
            "Date entered earlier than date Submitted")

    'Display error message
    Me.errKEMH_Date_Responded.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtKEMH_Date_Responded.value) Then
        Me.txtKEMH_Date_Responded.value = Format(Me.txtKEMH_Date_Responded.value, "dd-mmm-yyyy")
    End If
     
End Sub

Private Sub txtKEMH_Date_Approved_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtKEMH_Date_Approved.value, Me.txtKEMH_Date_Submitted.value, _
            "Date entered earlier than date Submitted")

    'Display error message
    Me.errKEMH_Date_Approved.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtKEMH_Date_Approved.value) Then
        Me.txtKEMH_Date_Approved.value = Format(Me.txtKEMH_Date_Approved.value, "dd-mmm-yyyy")
    End If
     
End Sub

Private Sub txtSJOG_S_Date_Submitted_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtSJOG_S_Date_Submitted.value)
    
    'Display error message
    Me.errSJOG_S_Date_Submitted.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtSJOG_S_Date_Submitted.value) Then
        Me.txtSJOG_S_Date_Submitted.value = Format(Me.txtSJOG_S_Date_Submitted.value, "dd-mmm-yyyy")
    End If
    
End Sub

Private Sub txtSJOG_S_Date_Responded_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtSJOG_S_Date_Responded.value, Me.txtSJOG_S_Date_Submitted.value, _
            "Date entered earlier than date Submitted")

    'Display error message
    Me.errSJOG_S_Date_Responded.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtSJOG_S_Date_Responded.value) Then
        Me.txtSJOG_S_Date_Responded.value = Format(Me.txtSJOG_S_Date_Responded.value, "dd-mmm-yyyy")
    End If
     
End Sub

Private Sub txtSJOG_S_Date_Approved_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtSJOG_S_Date_Approved.value, Me.txtSJOG_S_Date_Submitted.value, _
            "Date entered earlier than date Submitted")

    'Display error message
    Me.errSJOG_S_Date_Approved.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtSJOG_S_Date_Approved.value) Then
        Me.txtSJOG_S_Date_Approved.value = Format(Me.txtSJOG_S_Date_Approved.value, "dd-mmm-yyyy")
    End If
     
End Sub

Private Sub txtSJOG_L_Date_Submitted_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtSJOG_L_Date_Submitted.value)
    
    'Display error message
    Me.errSJOG_L_Date_Submitted.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtSJOG_L_Date_Submitted.value) Then
        Me.txtSJOG_L_Date_Submitted.value = Format(Me.txtSJOG_L_Date_Submitted.value, "dd-mmm-yyyy")
    End If
    
End Sub

Private Sub txtSJOG_L_Date_Responded_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtSJOG_L_Date_Responded.value, Me.txtSJOG_L_Date_Submitted.value, _
            "Date entered earlier than date Submitted")

    'Display error message
    Me.errSJOG_L_Date_Responded.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtSJOG_L_Date_Responded.value) Then
        Me.txtSJOG_L_Date_Responded.value = Format(Me.txtSJOG_L_Date_Responded.value, "dd-mmm-yyyy")
    End If
     
End Sub

Private Sub txtSJOG_L_Date_Approved_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtSJOG_L_Date_Approved.value, Me.txtSJOG_L_Date_Submitted.value, _
            "Date entered earlier than date Submitted")

    'Display error message
    Me.errSJOG_L_Date_Approved.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtSJOG_L_Date_Approved.value) Then
        Me.txtSJOG_L_Date_Approved.value = Format(Me.txtSJOG_L_Date_Approved.value, "dd-mmm-yyyy")
    End If
     
End Sub

Private Sub txtSJOG_M_Date_Submitted_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtSJOG_M_Date_Submitted.value)
    
    'Display error message
    Me.errSJOG_M_Date_Submitted.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtSJOG_M_Date_Submitted.value) Then
        Me.txtSJOG_M_Date_Submitted.value = Format(Me.txtSJOG_M_Date_Submitted.value, "dd-mmm-yyyy")
    End If
    
End Sub

Private Sub txtSJOG_M_Date_Responded_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtSJOG_M_Date_Responded.value, Me.txtSJOG_M_Date_Submitted.value, _
            "Date entered earlier than date Submitted")

    'Display error message
    Me.errSJOG_M_Date_Responded.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtSJOG_M_Date_Responded.value) Then
        Me.txtSJOG_M_Date_Responded.value = Format(Me.txtSJOG_M_Date_Responded.value, "dd-mmm-yyyy")
    End If
     
End Sub

Private Sub txtSJOG_M_Date_Approved_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtSJOG_M_Date_Approved.value, Me.txtSJOG_M_Date_Submitted.value, _
            "Date entered earlier than date Submitted")

    'Display error message
    Me.errSJOG_M_Date_Approved.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtSJOG_M_Date_Approved.value) Then
        Me.txtSJOG_M_Date_Approved.value = Format(Me.txtSJOG_M_Date_Approved.value, "dd-mmm-yyyy")
    End If
     
End Sub
Private Sub txtOthers_Date_Submitted_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtOthers_Date_Submitted.value)
    
    'Display error message
    Me.errOthers_Date_Submitted.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtOthers_Date_Submitted.value) Then
        Me.txtOthers_Date_Submitted.value = Format(Me.txtOthers_Date_Submitted.value, "dd-mmm-yyyy")
    End If
    
End Sub

Private Sub txtOthers_Date_Responded_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtOthers_Date_Responded.value, Me.txtOthers_Date_Submitted.value, _
            "Date entered earlier than date Submitted")

    'Display error message
    Me.errOthers_Date_Responded.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtOthers_Date_Responded.value) Then
        Me.txtOthers_Date_Responded.value = Format(Me.txtOthers_Date_Responded.value, "dd-mmm-yyyy")
    End If
     
End Sub

Private Sub txtOthers_Date_Approved_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtOthers_Date_Approved.value, Me.txtOthers_Date_Submitted.value, _
            "Date entered earlier than date Submitted")

    'Display error message
    Me.errOthers_Date_Approved.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtOthers_Date_Approved.value) Then
        Me.txtOthers_Date_Approved.value = Format(Me.txtOthers_Date_Approved.value, "dd-mmm-yyyy")
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
        
        .Range(58) = String_to_Date(Me.txtPCH_Date_Submitted.value)
        .Range(59) = String_to_Date(Me.txtPCH_Date_Responded.value)
        .Range(60) = String_to_Date(Me.txtPCH_Date_Approved.value)
        
        .Range(61) = String_to_Date(Me.txtTKI_Date_Submitted.value)
        .Range(62) = String_to_Date(Me.txtTKI_Date_Responded.value)
        .Range(63) = String_to_Date(Me.txtTKI_Date_Approved.value)
        
        .Range(64) = String_to_Date(Me.txtKEMH_Date_Submitted.value)
        .Range(65) = String_to_Date(Me.txtKEMH_Date_Responded.value)
        .Range(66) = String_to_Date(Me.txtKEMH_Date_Approved.value)
        
        .Range(67) = String_to_Date(Me.txtSJOG_S_Date_Submitted.value)
        .Range(68) = String_to_Date(Me.txtSJOG_S_Date_Responded.value)
        .Range(69) = String_to_Date(Me.txtSJOG_S_Date_Approved.value)
        
        .Range(70) = String_to_Date(Me.txtSJOG_L_Date_Submitted.value)
        .Range(71) = String_to_Date(Me.txtSJOG_L_Date_Responded.value)
        .Range(72) = String_to_Date(Me.txtSJOG_L_Date_Approved.value)
        
        .Range(73) = String_to_Date(Me.txtSJOG_M_Date_Submitted.value)
        .Range(74) = String_to_Date(Me.txtSJOG_M_Date_Responded.value)
        .Range(75) = String_to_Date(Me.txtSJOG_M_Date_Approved.value)
    
        .Range(76) = Me.txtOthers_Committee.value
        .Range(77) = String_to_Date(Me.txtOthers_Date_Submitted.value)
        .Range(78) = String_to_Date(Me.txtOthers_Date_Responded.value)
        .Range(79) = String_to_Date(Me.txtOthers_Date_Approved.value)
        
        .Range(80) = Me.txtReminder.value
        
        'Update version control
        .Range(81) = Now
        .Range(82) = Username
        
    End With
    
    'Access version control
    Call LogLastAccess
    
    Call UserForm_Initialize

End Sub



'----------------- Navigation section Toggles ----------------

Private Sub tglNav_Click()
    'PURPOSE: Closes current form and open Nav form
    Unload form043_Governance
    
    form00_Nav.Show False
End Sub

Private Sub tglStudyDetail_Click()
    'PURPOSE: Closes current form and open Study Details form
    Unload form043_Governance
    
    form01_StudyDetail.Show False
End Sub

Private Sub tglCDA_FS_Click()
    'PURPOSE: Closes current form and open CDA / FS form
    Unload form043_Governance
    
    form02_CDA_FS.Show False
End Sub

Private Sub tglSiteSelect_Click()
    'PURPOSE: Closes current form and open Site Select form
    Unload form043_Governance
    
    form03_SiteSelect.Show False
End Sub

Private Sub tglRecruitment_Click()
    'PURPOSE: Closes current form and open Recruitment form
    Unload form043_Governance
    
    form041_Recruitment.Show False
End Sub

Private Sub tglEthics_Click()
    'PURPOSE: Closes current form and open Ethics form
    Unload form043_Governance
    
    form042_Ethics.Show False
End Sub

Private Sub tglBudget_Click()
    'PURPOSE: Closes current form and open Budget form
    Unload form043_Governance
    
    form044_Budget.Show False
End Sub

Private Sub tglIndemnity_Click()
    'PURPOSE: Closes current form and open Indemnity form
    Unload form043_Governance
    
    form045_Indemnity.Show False
End Sub

Private Sub tglCTRA_Click()
    'PURPOSE: Closes current form and open CTRA form
    Unload form043_Governance
    
    form05_CTRA.Show False
End Sub

Private Sub tglFinDisc_Click()
    'PURPOSE: Closes current form and open Fin. Disc. form
    Unload form043_Governance
    
    form06_FinDisc.Show False
End Sub

Private Sub tglSIV_Click()
    'PURPOSE: Closes current form and open SIV form
    Unload form043_Governance
    
    form07_SIV.Show False
End Sub
