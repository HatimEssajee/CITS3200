VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form043_Governance 
   Caption         =   "Governance Review"
   ClientHeight    =   1728
   ClientLeft      =   -384
   ClientTop       =   -1860
   ClientWidth     =   3372
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
        Me.txtDate_RGC.value = Format(.Range(98).value, "dd-mmm-yyyy")
        Me.txtDate_UWA.value = Format(.Range(99).value, "dd-mmm-yyyy")
        Me.txtDate_Finance.value = Format(.Range(100).value, "dd-mmm-yyyy")
        Me.txtDate_COO.value = Format(.Range(101).value, "dd-mmm-yyyy")
        Me.txtDate_VTG.value = Format(.Range(102).value, "dd-mmm-yyyy")
        Me.txtDate_Company.value = Format(.Range(103).value, "dd-mmm-yyyy")
        Me.txtDate_Finalised.value = Format(.Range(104).value, "dd-mmm-yyyy")
        Me.txtReminder = .Range(105).value
    End With
    
    'Access version control
    Call LogLastAccess
    
    'Depress and make toggle green on nav bar
    Me.tglReviews.value = True
    Me.tglReviews.BackColor = vbGreen
    Me.tglGovernance.value = True
    Me.tglGovernance.BackColor = vbGreen
    
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
        
        .Range(98) = Me.txtDate_RGC.value
        .Range(99) = Me.txtDate_UWA.value
        .Range(100) = Me.txtDate_Finance.value
        .Range(101) = Me.txtDate_COO.value
        .Range(102) = Me.txtDate_VTG.value
        .Range(103) = Me.txtDate_Company.value
        .Range(104) = Me.txtDate_Finalised.value
        .Range(105) = Me.txtReminder.value
        
        'Update version control
        .Range(106) = Now
        .Range(107) = Username
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
