VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form06_FinDisc 
   Caption         =   "Financial Disclosure"
   ClientHeight    =   6096
   ClientLeft      =   -384
   ClientTop       =   -1632
   ClientWidth     =   11052
   OleObjectBlob   =   "form06_FinDisc.frx":0000
End
Attribute VB_Name = "form06_FinDisc"
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
    
    'Clear user form
    'source: https://www.mrexcel.com/board/threads/loop-through-controls-on-a-userform.427103/
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
    
    'Read information from register table
    With RegTable.ListRows(RowIndex)
        Me.txtStudyName.value = .Range(10).value
        Me.txtFinDisc_Complete.value = Format(.Range(108).value, "dd-mmm-yyyy")
        Me.txtReminder.value = .Range(109).value
    End With
    
    'Access version control
    Call LogLastAccess
    
    'Depress and make toggle green on nav bar
    Me.tglFinDisc.value = True
    Me.tglFinDisc.BackColor = vbGreen
    
End Sub

Private Sub txtFinDisc_Complete_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtFinDisc_Complete)
    
    Me.errFinDisc_Complete.Caption = err
    
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
        
        .Range(108) = String_to_Date(Me.txtFinDisc_Complete.value)
        .Range(109) = Me.txtReminder.value
        
        'Update version control
        .Range(110) = Now
        .Range(111) = Username
    End With
    
    'Access version control
    Call LogLastAccess
    
    Call UserForm_Initialize

End Sub


'----------------- Navigation section Toggles ----------------

Private Sub tglNav_Click()
    'PURPOSE: Closes current form and open Nav form
    Unload form06_FinDisc
    
    form00_Nav.Show False
End Sub

Private Sub tglStudyDetail_Click()
    'PURPOSE: Closes current form and open Study Details form
    Unload form06_FinDisc
    
    form01_StudyDetail.Show False
End Sub

Private Sub tglCDA_FS_Click()
    'PURPOSE: Closes current form and open CDA / FS form
    Unload form06_FinDisc
    
    form02_CDA_FS.Show False
End Sub

Private Sub tglSiteSelect_Click()
    'PURPOSE: Closes current form and open Site Selection form
    Unload form06_FinDisc
    
    form03_SiteSelect.Show False
End Sub

Private Sub tglReviews_Click()
    'PURPOSE: Closes current form and open Reviews form - Recruitment tab
    Unload form06_FinDisc
    
    form041_Recruitment.Show False
End Sub

Private Sub tglCTRA_Click()
    'PURPOSE: Closes current form and open CTRA form
    Unload form06_FinDisc
    
    form05_CTRA.Show False
End Sub

Private Sub tglSIV_Click()
    'PURPOSE: Closes current form and open SIV form
    Unload form06_FinDisc
    
    form07_SIV.Show False
End Sub

