VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form06_Ethics 
   Caption         =   "Ethics Review"
   ClientHeight    =   6720
   ClientLeft      =   -420
   ClientTop       =   -1995
   ClientWidth     =   10320
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
        Me.txtStudyName.Value = .Range(10).Value
        
        Me.txtCAHS_Date_Submitted.Value = Format(.Range(41).Value, "dd-mmm-yyyy")
        Me.txtCAHS_Date_Responded.Value = Format(.Range(42).Value, "dd-mmm-yyyy")
        Me.txtCAHS_Date_Resubmitted.Value = Format(.Range(43).Value, "dd-mmm-yyyy")
        Me.txtCAHS_Date_Approved.Value = Format(.Range(44).Value, "dd-mmm-yyyy")
        
        Me.txtNMA_Committee.Value = .Range(45).Value
        Me.txtNMA_Date_Submitted.Value = Format(.Range(46).Value, "dd-mmm-yyyy")
        Me.txtNMA_Date_Approved.Value = Format(.Range(47).Value, "dd-mmm-yyyy")
        
        Me.txtWNHS_Date_Submitted.Value = Format(.Range(48).Value, "dd-mmm-yyyy")
        Me.txtWNHS_Date_Approved.Value = Format(.Range(49).Value, "dd-mmm-yyyy")
        
        Me.txtSJOG_Date_Submitted.Value = Format(.Range(50).Value, "dd-mmm-yyyy")
        Me.txtSJOG_Date_Approved.Value = Format(.Range(51).Value, "dd-mmm-yyyy")
        
        Me.txtOthers_Committee.Value = .Range(52).Value
        Me.txtOthers_Date_Submitted.Value = Format(.Range(53).Value, "dd-mmm-yyyy")
        Me.txtOthers_Date_Approved.Value = Format(.Range(54).Value, "dd-mmm-yyyy")
        
        Me.txtReminder.Value = .Range(55).Value
        
    End With
    
    'Access version control
    Call LogLastAccess
    
    'Depress and make toggle green on nav bar
    Me.tglReviews.Value = True
    Me.tglReviews.BackColor = vbGreen
    Me.tglEthics.Value = True
    Me.tglEthics.BackColor = vbGreen
    
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

Private Sub cmdClose_Click()
    'PURPOSE: Closes current form
    
    'Access version control
    Call LogLastAccess
    
    Unload Me
    
End Sub

Private Sub cmdEdit_Click()
    'PURPOSE: Apply changes into Register table
    With RegTable.ListRows(RowIndex)
        
        .Range(41) = String_to_Date(Me.txtCAHS_Date_Submitted.Value)
        .Range(42) = String_to_Date(Me.txtCAHS_Date_Responded.Value)
        .Range(43) = String_to_Date(Me.txtCAHS_Date_Resubmitted.Value)
        .Range(44) = String_to_Date(Me.txtCAHS_Date_Approved.Value)
        
        .Range(45) = Me.txtNMA_Committee.Value
        .Range(46) = String_to_Date(Me.txtNMA_Date_Submitted.Value)
        .Range(47) = String_to_Date(Me.txtNMA_Date_Approved.Value)
        
        .Range(48) = String_to_Date(Me.txtWNHS_Date_Submitted.Value)
        .Range(49) = String_to_Date(Me.txtWNHS_Date_Approved.Value)
        
        .Range(50) = String_to_Date(Me.txtSJOG_Date_Submitted.Value)
        .Range(51) = String_to_Date(Me.txtSJOG_Date_Approved.Value)
        
        .Range(52) = Me.txtOthers_Committee.Value
        .Range(53) = String_to_Date(Me.txtOthers_Date_Submitted.Value)
        .Range(54) = String_to_Date(Me.txtOthers_Date_Approved.Value)
        
        .Range(55) = Me.txtReminder.Value
        
        'Apply completion status
        If Application.CountA(Range(RegTable.DataBodyRange.Cells(RowIndex, 41), _
            RegTable.DataBodyRange.Cells(RowIndex, 54))) = 0 Then
            .Range(121).Value = False
        ElseIf .Range(41).Value = vbNullString Then
            .Range(121).Value = vbNullString
        ElseIf IsDate(.Range(44).Value) Then
            .Range(121).Value = True
        Else
            .Range(121).Value = False
        End If
        
        If .Range(45).Value = vbNullString And .Range(46).Value = vbNullString Then
            .Range(122).Value = vbNullString
        ElseIf IsDate(.Range(47).Value) Then
            .Range(122).Value = True
        Else
            .Range(122).Value = False
        End If
        
        If .Range(48).Value = vbNullString Then
            .Range(123).Value = vbNullString
        ElseIf IsDate(.Range(49).Value) Then
            .Range(123).Value = True
        Else
            .Range(123).Value = False
        End If
        
        If .Range(50).Value = vbNullString Then
            .Range(124).Value = vbNullString
        ElseIf IsDate(.Range(51).Value) Then
            .Range(124).Value = True
        Else
            .Range(124).Value = False
        End If
        
        If .Range(52).Value = vbNullString And .Range(53).Value = vbNullString Then
            .Range(125).Value = vbNullString
        ElseIf IsDate(.Range(54).Value) Then
            .Range(125).Value = True
        Else
            .Range(125).Value = False
        End If
       
        'Update version control
        .Range(56) = Now
        .Range(57) = Username
    End With
    
    'Access version control
    Call LogLastAccess
    
    Call UserForm_Initialize

End Sub


'----------------- Navigation section Toggles ----------------

Private Sub tglNav_Click()
    'PURPOSE: Closes current form and open Nav form
    Unload form042_Ethics
    
    form00_Nav.Show False
End Sub

Private Sub tglStudyDetail_Click()
    'PURPOSE: Closes current form and open Study Details form
    Unload form042_Ethics
    
    form01_StudyDetail.Show False
End Sub

Private Sub tglCDA_FS_Click()
    'PURPOSE: Closes current form and open CDA / FS form
    Unload form042_Ethics
    
    form02_CDA_FS.Show False
    form02_CDA_FS.multiCDA_FS.Value = 0
End Sub

Private Sub tglSiteSelect_Click()
    'PURPOSE: Closes current form and open Site Select form
    Unload form042_Ethics
    
    form03_SiteSelect.Show False
End Sub

Private Sub tglRecruitment_Click()
    'PURPOSE: Closes current form and open Recruitment form
    Unload form042_Ethics
    
    form041_Recruitment.Show False
End Sub

Private Sub tglGovernance_Click()
    'PURPOSE: Closes current form and open Governance form
    Unload form042_Ethics
    
    form043_Governance.Show False
    form043_Governance.multiGov.Value = 0
End Sub

Private Sub tglBudget_Click()
    'PURPOSE: Closes current form and open Budget form
    Unload form042_Ethics
    
    form044_Budget.Show False
End Sub

Private Sub tglIndemnity_Click()
    'PURPOSE: Closes current form and open Indemnity form
    Unload form042_Ethics
    
    form045_Indemnity.Show False
End Sub

Private Sub tglCTRA_Click()
    'PURPOSE: Closes current form and open CTRA form
    Unload form042_Ethics
    
    form05_CTRA.Show False
End Sub

Private Sub tglFinDisc_Click()
    'PURPOSE: Closes current form and open Fin. Disc. form
    Unload form042_Ethics
    
    form06_FinDisc.Show False
End Sub

Private Sub tglSIV_Click()
    'PURPOSE: Closes current form and open SIV form
    Unload form042_Ethics
    
    form07_SIV.Show False
End Sub



