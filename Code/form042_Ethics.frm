VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form042_Ethics 
   Caption         =   "Ethics Review"
   ClientHeight    =   6744
   ClientLeft      =   -456
   ClientTop       =   -2100
   ClientWidth     =   8268.001
   OleObjectBlob   =   "form042_Ethics.frx":0000
End
Attribute VB_Name = "form042_Ethics"
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
    
    For Each pPage In Me.multiEthics.Pages
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
        
        Me.txtCAHS_Date_Submitted.value = Format(.Range(41).value, "dd-mmm-yyyy")
        Me.txtCAHS_Date_Responded.value = Format(.Range(42).value, "dd-mmm-yyyy")
        Me.txtCAHS_Date_Resubmitted.value = Format(.Range(43).value, "dd-mmm-yyyy")
        Me.txtCAHS_Date_Approved.value = Format(.Range(44).value, "dd-mmm-yyyy")
        
        Me.txtNMA_Committee.value = .Range(45).value
        Me.txtNMA_Date_Submitted.value = Format(.Range(46).value, "dd-mmm-yyyy")
        Me.txtNMA_Date_Approved.value = Format(.Range(47).value, "dd-mmm-yyyy")
        
        Me.txtWNHS_Date_Submitted.value = Format(.Range(48).value, "dd-mmm-yyyy")
        Me.txtWNHS_Date_Approved.value = Format(.Range(49).value, "dd-mmm-yyyy")
        
        Me.txtSJOG_Date_Submitted.value = Format(.Range(50).value, "dd-mmm-yyyy")
        Me.txtSJOG_Date_Approved.value = Format(.Range(51).value, "dd-mmm-yyyy")
        
        Me.txtOthers_Committee.value = .Range(52).value
        Me.txtOthers_Date_Submitted.value = Format(.Range(53).value, "dd-mmm-yyyy")
        Me.txtOthers_Date_Approved.value = Format(.Range(54).value, "dd-mmm-yyyy")
        
        Me.txtReminder.value = .Range(55).value
        
    End With
    
    'Access version control
    Call LogLastAccess
    
    'Depress and make toggle green on nav bar
    Me.tglReviews.value = True
    Me.tglReviews.BackColor = vbGreen
    Me.tglEthics.value = True
    Me.tglEthics.BackColor = vbGreen
    
    
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
        
        .Range(41) = String_to_Date(Me.txtCAHS_Date_Submitted.value)
        .Range(42) = String_to_Date(Me.txtCAHS_Date_Responded.value)
        .Range(43) = String_to_Date(Me.txtCAHS_Date_Resubmitted.value)
        .Range(44) = String_to_Date(Me.txtCAHS_Date_Approved.value)
        
        .Range(45) = Me.txtNMA_Committee.value
        .Range(46) = String_to_Date(Me.txtNMA_Date_Submitted.value)
        .Range(47) = String_to_Date(Me.txtNMA_Date_Approved.value)
        
        .Range(48) = String_to_Date(Me.txtWNHS_Date_Submitted.value)
        .Range(49) = String_to_Date(Me.txtWNHS_Date_Approved.value)
        
        .Range(50) = String_to_Date(Me.txtSJOG_Date_Submitted.value)
        .Range(51) = String_to_Date(Me.txtSJOG_Date_Approved.value)
        
        .Range(52) = Me.txtOthers_Committee.value
        .Range(53) = String_to_Date(Me.txtOthers_Date_Submitted.value)
        .Range(54) = String_to_Date(Me.txtOthers_Date_Approved.value)
        
        .Range(55) = Me.txtReminder.value
        
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



