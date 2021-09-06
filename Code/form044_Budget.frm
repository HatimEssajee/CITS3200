VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form044_Budget 
   Caption         =   "Budget Review"
   ClientHeight    =   5364
   ClientLeft      =   -408
   ClientTop       =   -1992
   ClientWidth     =   6384
   OleObjectBlob   =   "form044_Budget.frx":0000
End
Attribute VB_Name = "form044_Budget"
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
    
    For Each pPage In Me.multiBudget.Pages
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
        Me.txtVTG_Date_Submitted.Value = Format(.Range(83).Value, "dd-mmm-yyyy")
        Me.txtVTG_Date_Finalised.Value = Format(.Range(84).Value, "dd-mmm-yyyy")
        Me.txtVTG_Date_Approved.Value = Format(.Range(85).Value, "dd-mmm-yyyy")
        
        Me.txtTKI_Date_Approved.Value = Format(.Range(86).Value, "dd-mmm-yyyy")
        
        Me.txtPharm_Date_Quote.Value = Format(.Range(87).Value, "dd-mmm-yyyy")
        Me.txtPharm_Date_Finalised.Value = Format(.Range(88).Value, "dd-mmm-yyyy")
        
        Me.txtReminder.Value = .Range(89).Value
    End With
    
    'Access version control
    Call LogLastAccess
    
    'Depress and make toggle green on nav bar
    Me.tglReviews.Value = True
    Me.tglReviews.BackColor = vbGreen
    Me.tglBudget.Value = True
    Me.tglBudget.BackColor = vbGreen
    
    'Run date validation on data entered
    Call txtVTG_Date_Submitted_AfterUpdate
    Call txtVTG_Date_Finalised_AfterUpdate
    Call txtVTG_Date_Approved_AfterUpdate
    
    Call txtTKI_Date_Approved_AfterUpdate
    
    Call txtPharm_Date_Quote_AfterUpdate
    Call txtPharm_Date_Finalised_AfterUpdate
    
End Sub

Private Sub txtVTG_Date_Finalised_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtVTG_Date_Finalised.Value)
    
    'Display error message
    Me.errVTG_Date_Finalised.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtVTG_Date_Finalised.Value) Then
        Me.txtVTG_Date_Finalised.Value = Format(Me.txtVTG_Date_Finalised.Value, "dd-mmm-yyyy")
    End If
    
End Sub

Private Sub txtVTG_Date_Submitted_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtVTG_Date_Submitted.Value, Me.txtVTG_Date_Finalised.Value, _
            "Date entered earlier than date Finalised")

    'Display error message
    Me.errVTG_Date_Submitted.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtVTG_Date_Submitted.Value) Then
        Me.txtVTG_Date_Submitted.Value = Format(Me.txtVTG_Date_Submitted.Value, "dd-mmm-yyyy")
    End If
     
End Sub

Private Sub txtVTG_Date_Approved_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtVTG_Date_Approved.Value, Me.txtVTG_Date_Submitted.Value, _
            "Date entered earlier than date Submitted")

    'Display error message
    Me.errVTG_Date_Approved.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtVTG_Date_Approved.Value) Then
        Me.txtVTG_Date_Approved.Value = Format(Me.txtVTG_Date_Approved.Value, "dd-mmm-yyyy")
    End If
     
End Sub

Private Sub txtTKI_Date_Approved_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtTKI_Date_Approved.Value)
    
    'Display error message
    Me.errTKI_Date_Approved.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtTKI_Date_Approved.Value) Then
        Me.txtTKI_Date_Approved.Value = Format(Me.txtTKI_Date_Approved.Value, "dd-mmm-yyyy")
    End If
    
End Sub

Private Sub txtPharm_Date_Quote_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtPharm_Date_Quote.Value)
    
    'Display error message
    Me.errPharm_Date_Quote.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtPharm_Date_Quote.Value) Then
        Me.txtPharm_Date_Quote.Value = Format(Me.txtPharm_Date_Quote.Value, "dd-mmm-yyyy")
    End If
    
End Sub

Private Sub txtPharm_Date_Finalised_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtPharm_Date_Finalised.Value, Me.txtPharm_Date_Quote.Value, _
            "Date entered earlier than date" & Chr(10) & "Quote was received")

    'Display error message
    Me.errPharm_Date_Finalised.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtPharm_Date_Finalised.Value) Then
        Me.txtPharm_Date_Finalised.Value = Format(Me.txtPharm_Date_Finalised.Value, "dd-mmm-yyyy")
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
        
        .Range(83) = String_to_Date(Me.txtVTG_Date_Submitted.Value)
        .Range(84) = String_to_Date(Me.txtVTG_Date_Finalised.Value)
        .Range(85) = String_to_Date(Me.txtVTG_Date_Approved.Value)
        .Range(86) = String_to_Date(Me.txtTKI_Date_Approved.Value)
        .Range(87) = String_to_Date(Me.txtPharm_Date_Quote.Value)
        .Range(88) = String_to_Date(Me.txtPharm_Date_Finalised.Value)
        .Range(89) = Me.txtReminder.Value
        
        'Update version control
        .Range(90) = Now
        .Range(91) = Username
    End With
    
    'Access version control
    Call LogLastAccess
    
    Call UserForm_Initialize

End Sub


'----------------- Navigation section Toggles ----------------

Private Sub tglNav_Click()
    'PURPOSE: Closes current form and open Nav form
    Unload form044_Budget
    
    form00_Nav.Show False
End Sub

Private Sub tglStudyDetail_Click()
    'PURPOSE: Closes current form and open Study Details form
    Unload form044_Budget
    
    form01_StudyDetail.Show False
End Sub

Private Sub tglCDA_FS_Click()
    'PURPOSE: Closes current form and open CDA / FS form
    Unload form044_Budget
    
    form02_CDA_FS.Show False
End Sub

Private Sub tglSiteSelect_Click()
    'PURPOSE: Closes current form and open Site Select form
    Unload form044_Budget
    
    form03_SiteSelect.Show False
End Sub

Private Sub tglRecruitment_Click()
    'PURPOSE: Closes current form and open Recruitment form
    Unload form044_Budget
    
    form041_Recruitment.Show False
End Sub

Private Sub tglEthics_Click()
    'PURPOSE: Closes current form and open Ethics form
    Unload form044_Budget
    
    form042_Ethics.Show False
End Sub

Private Sub tglGovernance_Click()
    'PURPOSE: Closes current form and open Budget form
    Unload form044_Budget
    
    form043_Governance.Show False
End Sub

Private Sub tglIndemnity_Click()
    'PURPOSE: Closes current form and open Indemnity form
    Unload form044_Budget
    
    form045_Indemnity.Show False
End Sub

Private Sub tglCTRA_Click()
    'PURPOSE: Closes current form and open CTRA form
    Unload form044_Budget
    
    form05_CTRA.Show False
End Sub

Private Sub tglFinDisc_Click()
    'PURPOSE: Closes current form and open Fin. Disc. form
    Unload form044_Budget
    
    form06_FinDisc.Show False
End Sub

Private Sub tglSIV_Click()
    'PURPOSE: Closes current form and open SIV form
    Unload form044_Budget
    
    form07_SIV.Show False
End Sub



