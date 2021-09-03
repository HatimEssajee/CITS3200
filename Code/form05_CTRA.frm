VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form05_CTRA 
   Caption         =   "Project Form"
   ClientHeight    =   3312
   ClientLeft      =   -264
   ClientTop       =   -1284
   ClientWidth     =   6384
   OleObjectBlob   =   "form05_CTRA.frx":0000
End
Attribute VB_Name = "form05_CTRA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False













Option Explicit

Private Sub UserForm_Activate()
    'PURPOSE: Reposition userform to Top Left of application Window and fix size
    'source: https://www.mrexcel.com/board/threads/userform-startup-position.671108/
    Me.StartUpPosition = 0
    Me.Top = Application.Top + 25
    Me.Left = Application.Left + 25
    Me.Height = UHeight
    Me.Width = UWidth

End Sub

Private Sub UserForm_Initialize()
    'PURPOSE: Clear form on initialization
    'Source: https://www.contextures.com/xlUserForm02.html
    'Source: https://www.contextures.com/Excel-VBA-ComboBox-Lists.html
    Dim ws As Worksheet
    Dim ctrl As MSForms.Control

    Set ws = Worksheets("Lookup Lists")
    
    'Clear user form
    'source: https://www.mrexcel.com/board/threads/loop-through-controls-on-a-userform.427103/
    For Each ctrl In Me.Controls
        Select Case True
                Case TypeOf ctrl Is MSForms.CheckBox
                    ctrl.Value = False
                Case TypeOf ctrl Is MSForms.TextBox
                    ctrl.Value = ""
                Case TypeOf ctrl Is MSForms.ComboBox
                    ctrl.Value = ""
                    ctrl.Clear
                Case TypeOf ctrl Is MSForms.ListBox
                    ctrl.Value = ""
                    ctrl.Clear
            End Select
    Next ctrl
    
    Me.tglCTRA.Value = True
    Me.tglCTRA.BackColor = vbGreen
    
End Sub

Private Sub cmdClose_Click()
    'PURPOSE: Closes current form
    Unload Me
End Sub

Private Sub cmdEdit_Click()
    'PURPOSE: Apply changes into Register table

End Sub


'----------------- Navigation section Toggles ----------------

Private Sub tglNav_Click()
    'PURPOSE: Closes current form and open Nav form
    Unload form05_CTRA
    
    form00_Nav.Show False
End Sub

Private Sub tglStudyDetail_Click()
    'PURPOSE: Closes current form and open Study Details form
    Unload form05_CTRA
    
    form01_StudyDetail.Show False
End Sub

Private Sub tglCDA_FS_Click()
    'PURPOSE: Closes current form and open CDA / FS form
    Unload form05_CTRA
    
    form02_CDA_FS.Show False
End Sub

Private Sub tglSiteSelect_Click()
    'PURPOSE: Closes current form and open Site Select form
    Unload form05_CTRA
    
    form03_SiteSelect.Show False
End Sub

Private Sub tglReviews_Click()
    'PURPOSE: Closes current form and open Reviews form - Recruitment tab
    Unload form05_CTRA
    
    form041_Recruitment.Show False
End Sub


Private Sub tglFinDisc_Click()
    'PURPOSE: Closes current form and open Fin. Disc. form
    Unload form05_CTRA
    
    form06_FinDisc.Show False
End Sub

Private Sub tglSIV_Click()
    'PURPOSE: Closes current form and open SIV form
    Unload Me
    
    form07_SIV.Show False
End Sub

