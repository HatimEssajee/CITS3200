VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form041_Recruitment 
   Caption         =   "Recruitment Plan"
   ClientHeight    =   6408
   ClientLeft      =   -504
   ClientTop       =   -2208
   ClientWidth     =   9036.001
   OleObjectBlob   =   "form041_Recruitment.frx":0000
End
Attribute VB_Name = "form041_Recruitment"
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
    Dim cboList_RecruitStatus As Variant, item As Variant
    
    cboList_RecruitStatus = Array("In-progress", "Complete")
    
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
    
    'Fill combo box for study status
    For Each item In cboList_RecruitStatus
        cboRecruitStatus.AddItem item
    Next item
    
    'Read information from register table
    With RegTable.ListRows(RowIndex)
        Me.txtStudyName.value = .Range(10).value
        Me.txtDate_Plan.value = Format(.Range(36).value, "dd-mmm-yyyy")
        Me.cboRecruitStatus.value = .Range(37).value
        Me.txtReminder.value = .Range(38).value
    End With
    
    'Access version control
    Call LogLastAccess
    
    'Depress and make toggle green on nav bar
    Me.tglReviews.value = True
    Me.tglReviews.BackColor = vbGreen
    Me.tglRecruitment.value = True
    Me.tglRecruitment.BackColor = vbGreen
    
    'Run date validation on data entered
    Call txtDate_Plan_AfterUpdate
    
End Sub

Private Sub txtDate_Plan_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtDate_Plan.value)
    
    'Display error message
    Me.errDate_Plan.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtDate_Plan.value) Then
        Me.txtDate_Plan.value = Format(Me.txtDate_Plan.value, "dd-mmm-yyyy")
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
        
        .Range(36) = String_to_Date(Me.txtDate_Plan)
        .Range(37) = Me.cboRecruitStatus
        .Range(38) = Me.txtReminder
        
        'Update version control
        .Range(39) = Now
        .Range(40) = Username
    End With
    
    'Access version control
    Call LogLastAccess
    
    Call UserForm_Initialize

End Sub


'----------------- Navigation section Toggles ----------------

Private Sub tglNav_Click()
    'PURPOSE: Closes current form and open Nav form
    Unload form041_Recruitment
    
    form00_Nav.Show False
End Sub

Private Sub tglStudyDetail_Click()
    'PURPOSE: Closes current form and open Study Details form
    Unload form041_Recruitment
    
    form01_StudyDetail.Show False
End Sub

Private Sub tglCDA_FS_Click()
    'PURPOSE: Closes current form and open CDA / FS form
    Unload form041_Recruitment
    
    form02_CDA_FS.Show False
End Sub

Private Sub tglSiteSelect_Click()
    'PURPOSE: Closes current form and open Site Select form
    Unload form041_Recruitment
    
    form03_SiteSelect.Show False
End Sub

Private Sub tglEthics_Click()
    'PURPOSE: Closes current form and open Ethics form
    Unload form041_Recruitment
    
    form042_Ethics.Show False
End Sub

Private Sub tglGovernance_Click()
    'PURPOSE: Closes current form and open Governance form
    Unload form041_Recruitment
    
    form043_Governance.Show False
End Sub

Private Sub tglBudget_Click()
    'PURPOSE: Closes current form and open Budget form
    Unload form041_Recruitment
    
    form044_Budget.Show False
End Sub

Private Sub tglIndemnity_Click()
    'PURPOSE: Closes current form and open Indemnity form
    Unload form041_Recruitment
    
    form045_Indemnity.Show False
End Sub


Private Sub tglCTRA_Click()
    'PURPOSE: Closes current form and open CTRA form
    Unload form041_Recruitment
    
    form05_CTRA.Show False
End Sub

Private Sub tglFinDisc_Click()
    'PURPOSE: Closes current form and open Fin. Disc. form
    Unload form041_Recruitment
    
    form06_FinDisc.Show False
End Sub

Private Sub tglSIV_Click()
    'PURPOSE: Closes current form and open SIV form
    Unload form041_Recruitment
    
    form07_SIV.Show False
End Sub

