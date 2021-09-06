VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form01_StudyDetail 
   Caption         =   "Study Details"
   ClientHeight    =   5436
   ClientLeft      =   -435
   ClientTop       =   -1845
   ClientWidth     =   6975
   OleObjectBlob   =   "form01_StudyDetail.frx":0000
End
Attribute VB_Name = "form01_StudyDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





<<<<<<< HEAD
=======




>>>>>>> AnferneeAlviar
Option Explicit

Private Sub UserForm_Activate()
    'PURPOSE: Reposition userform to Top Left of application Window and fix size
    'source: https://www.mrexcel.com/board/threads/userform-startup-position.671108/
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
    
    'Highlight tab selected
    Me.tglStudyDetail.Value = True
    Me.tglStudyDetail.BackColor = vbGreen
    
    'Read information from register table
    With RegTable.ListRows(RowIndex)
        Me.txtProtocolNum.Value = .Range(9).Value
        Me.txtStudyName.Value = .Range(10).Value
        Me.txtSponsor.Value = .Range(11).Value
        Me.txtCRO.Value = .Range(12).Value
        Me.txtAgeRange.Value = .Range(13).Value
        Me.txtReminder.Value = .Range(14).Value
        
    End With
    
    'Edit LastAccess log
    Call LogLastAccess
    
    'Depress and make toggle green on nav bar
    Me.tglStudyDetail.Value = True
    Me.tglStudyDetail.BackColor = vbGreen
    
End Sub

Private Sub cmdClose_Click()
    'PURPOSE: Closes current form
    
    'Edit LastAccess log
    Call LogLastAccess
    
    Unload Me
    
End Sub

Private Sub cmdEdit_Click()
    'PURPOSE: Apply changes into Register table
    With RegTable.ListRows(RowIndex)
        .Range(9) = Me.txtProtocolNum.Value
        .Range(10) = Me.txtStudyName.Value
        .Range(11) = Me.txtSponsor.Value
        .Range(12) = Me.txtCRO.Value
        .Range(13) = Me.txtAgeRange.Value
        .Range(14) = Me.txtReminder.Value
        
        'Update version control
        .Range(15) = Now
        .Range(16) = Username
    End With
    
    'Access version control
    Call LogLastAccess
    
    Call UserForm_Initialize
End Sub


'----------------- Navigation section Toggles ----------------

Private Sub tglNav_Click()
    'PURPOSE: Closes current form and open Nav form
    Unload form01_StudyDetail
    
    form00_Nav.Show False
End Sub

Private Sub tglCDA_FS_Click()
    'PURPOSE: Closes current form and open CDA / FS form
    Unload form01_StudyDetail
    
    form02_CDA_FS.Show False
End Sub

Private Sub tglSiteSelect_Click()
    'PURPOSE: Closes current form and open Site Select form
    Unload form01_StudyDetail
    
    form03_SiteSelect.Show False
End Sub

Private Sub tglReviews_Click()
    'PURPOSE: Closes current form and open Reviews form - Recruitment tab
    Unload form01_StudyDetail
    
    form041_Recruitment.Show False
End Sub

Private Sub tglCTRA_Click()
    'PURPOSE: Closes current form and open CTRA form
    Unload form01_StudyDetail
    
    form05_CTRA.Show False
End Sub

Private Sub tglFinDisc_Click()
    'PURPOSE: Closes current form and open Fin. Disc. form
    Unload form01_StudyDetail
    
    form06_FinDisc.Show False
End Sub

Private Sub tglSIV_Click()
    'PURPOSE: Closes current form and open SIV form
    Unload form01_StudyDetail
    
    form07_SIV.Show False
End Sub



