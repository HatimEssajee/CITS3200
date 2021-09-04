VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form05_CTRA 
   Caption         =   "CTRA"
   ClientHeight    =   6960
   ClientLeft      =   -312
   ClientTop       =   -1620
   ClientWidth     =   8712.001
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
    Me.tglCTRA.value = True
    Me.tglCTRA.BackColor = vbGreen
    
End Sub

Private Sub txtDate_RGC_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtDate_RGC.value)
    
    Me.errDate_RGC.Caption = err
    
End Sub

Private Sub txtDate_UWA_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtDate_UWA.value)
    
    Me.errDate_UWA.Caption = err
    
End Sub

Private Sub txtDate_Finance_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtDate_Finance.value)
    
    Me.errFinance_RGC.Caption = err
    
End Sub

Private Sub txtDate_COO_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    Dim d1 As Variant
    Dim d2 As Variant
    
    err = Date_Validation(Me.txtDate_COO.value)
    d1 = String_to_Date(Me.txtDate_Finance.value)
    d2 = String_to_Date(Me.txtDate_COO.value)
    
    'If no date entry issue, check date for chronology
    If err = "" And d1 <> "" And d2 <> "" And d2 < d1 Then
        err = "Date entered earlier than finance signoff"
    End If
    
    Me.errDate_COO.Caption = err
    
End Sub

Private Sub txtDate_VTG_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    Dim d1 As Variant
    Dim d2 As Variant
    
    err = Date_Validation(Me.txtDate_VTG.value)
    d1 = String_to_Date(Me.txtDate_COO.value)
    d2 = String_to_Date(Me.txtDate_VTG.value)
    
    'If no date entry issue, check date for chronology
    If err = "" And d1 <> "" And d2 <> "" And d2 < d1 Then
        err = "Date entered earlier than COO signoff"
    End If
    
    Me.errDate_VTG.Caption = err
    
End Sub

Private Sub txtDate_Company_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    Dim d1 As Variant
    Dim d2 As Variant
    
    err = Date_Validation(Me.txtDate_Company.value)
    d1 = String_to_Date(Me.txtDate_VTG.value)
    d2 = String_to_Date(Me.txtDate_Company.value)
    
    'If no date entry issue, check date for chronology
    If err = "" And d1 <> "" And d2 <> "" And d2 < d1 Then
        err = "Date entered earlier than VTG sign-off"
    End If
    
    Me.errDate_Company.Caption = err
    
End Sub

Private Sub txtDate_Finalised_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    Dim d1 As Variant
    Dim d2 As Variant
    
    err = Date_Validation(Me.txtDate_Finalised.value)
    d1 = String_to_Date(Me.txtDate_Company.value)
    d2 = String_to_Date(Me.txtDate_Finalised.value)
    
    'If no date entry issue, check date for chronology
    If err = "" And d1 <> "" And d2 <> "" And d2 < d1 Then
        err = "Date entered earlier than company submission"
    End If
    
    Me.errDate_Finalised.Caption = err
    
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
        
        .Range(98) = String_to_Date(Me.txtDate_RGC.value)
        .Range(99) = String_to_Date(Me.txtDate_UWA.value)
        .Range(100) = String_to_Date(Me.txtDate_Finance.value)
        .Range(101) = String_to_Date(Me.txtDate_COO.value)
        .Range(102) = String_to_Date(Me.txtDate_VTG.value)
        .Range(103) = String_to_Date(Me.txtDate_Company.value)
        .Range(104) = String_to_Date(Me.txtDate_Finalised.value)
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

