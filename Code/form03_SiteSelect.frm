VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form03_SiteSelect 
   Caption         =   "Site Selection"
   ClientHeight    =   6108
   ClientLeft      =   -408
   ClientTop       =   -1752
   ClientWidth     =   10872
   OleObjectBlob   =   "form03_SiteSelect.frx":0000
End
Attribute VB_Name = "form03_SiteSelect"
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
    Dim cboList_StudyStatus As Variant, item As Variant
    
    cboList_StudyStatus = Array("On-site", "Virtual")
    
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
    
    'Fill combo box for study status
    For Each item In cboList_StudyStatus
        cboPrestudy_Type.AddItem item
        cboValidation_Type.AddItem item
    Next item
    
    'Read information from register table
    With RegTable.ListRows(RowIndex)
        Me.txtStudyName = .Range(10).value
        Me.txtPrestudy_Date.value = Format(.Range(28).value, "dd-mmm-yyyy")
        Me.cboPrestudy_Type.value = .Range(29).value
        Me.txtValidation_Date.value = Format(.Range(30).value, "dd-mmm-yyyy")
        Me.cboValidation_Type.value = .Range(31).value
        Me.txtSiteSelect.value = Format(.Range(32).value, "dd-mmm-yyyy")
        
        Me.txtReminder.value = .Range(25).value
    End With
    
    'Access version control
    Call LogLastAccess
    
    'Depress and make toggle green on nav bar
    Me.tglSiteSelect.value = True
    Me.tglSiteSelect.BackColor = vbGreen
    
    'Run date validation on data entered
    Call txtPrestudy_Date_AfterUpdate
    Call txtValidation_Date_AfterUpdate
    Call txtSiteSelect_AfterUpdate
    
End Sub

Private Sub txtPrestudy_Date_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtPrestudy_Date.value)
    
    'Display error message
    Me.errPrestudy_Date.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtPrestudy_Date.value) Then
        Me.txtPrestudy_Date.value = Format(Me.txtPrestudy_Date.value, "dd-mmm-yyyy")
    End If
    
End Sub

Private Sub txtValidation_Date_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtValidation_Date.value, Me.txtPrestudy_Date.value, _
            "Date entered earlier than date of" & Chr(10) & "Pre-study visit")

    'Display error message
    Me.errValidation_Date.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtValidation_Date.value) Then
        Me.txtValidation_Date.value = Format(Me.txtValidation_Date.value, "dd-mmm-yyyy")
    End If
     
End Sub

Private Sub txtSiteSelect_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtSiteSelect.value, Me.txtCDA_Recv_Sponsor.value, _
            "Date entered earlier than date of" & Chr(10) & "Validation visit")

    'Display error message
    Me.errSiteSelect.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtSiteSelect.value) Then
        Me.txtSiteSelect.value = Format(Me.txtSiteSelect.value, "dd-mmm-yyyy")
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
        
        .Range(28) = String_to_Date(Me.txtPrestudy_Date.value)
        .Range(29) = Me.cboPrestudy_Type.value
        .Range(30) = String_to_Date(Me.txtValidation_Date.value)
        .Range(31) = Me.cboValidation_Type.value
        .Range(32) = String_to_Date(Me.txtSiteSelect.value)
        .Range(33) = Me.txtReminder.value
        
        'Update version control
        .Range(34) = Now
        .Range(35) = Username
    End With
    
    'Access version control
    Call LogLastAccess
    
    Call UserForm_Initialize
End Sub


'----------------- Navigation section Toggles ----------------

Private Sub tglNav_Click()
    'PURPOSE: Closes current form and open Nav form
    Unload form03_SiteSelect
    
    form00_Nav.Show False
End Sub

Private Sub tglStudyDetail_Click()
    'PURPOSE: Closes current form and open Study Details form
    Unload form03_SiteSelect
    
    form01_StudyDetail.Show False
End Sub

Private Sub tglCDA_FS_Click()
    'PURPOSE: Closes current form and open CDA / FS form
    Unload form03_SiteSelect
    
    form02_CDA_FS.Show False
End Sub

Private Sub tglReviews_Click()
    'PURPOSE: Closes current form and open Reviews form - Recruitment tab
    Unload form03_SiteSelect
    
    form041_Recruitment.Show False
End Sub

Private Sub tglCTRA_Click()
    'PURPOSE: Closes current form and open CTRA form
    Unload form03_SiteSelect
    
    form05_CTRA.Show False
End Sub

Private Sub tglFinDisc_Click()
    'PURPOSE: Closes current form and open Fin. Disc. form
    Unload form03_SiteSelect
    
    form06_FinDisc.Show False
End Sub

Private Sub tglSIV_Click()
    'PURPOSE: Closes current form and open SIV form
    Unload form03_SiteSelect
    
    form07_SIV.Show False
End Sub

