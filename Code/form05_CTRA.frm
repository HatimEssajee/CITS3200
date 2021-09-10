VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form05_CTRA 
   Caption         =   "CTRA"
   ClientHeight    =   5148
   ClientLeft      =   -408
   ClientTop       =   -2088
   ClientWidth     =   6648
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
    
    'Read information from register table
    With RegTable.ListRows(RowIndex)
        Me.txtStudyName.Value = .Range(10).Value
        Me.txtDate_RGC.Value = Format(.Range(98).Value, "dd-mmm-yyyy")
        Me.txtDate_UWA.Value = Format(.Range(99).Value, "dd-mmm-yyyy")
        Me.txtDate_Finance.Value = Format(.Range(100).Value, "dd-mmm-yyyy")
        Me.txtDate_COO.Value = Format(.Range(101).Value, "dd-mmm-yyyy")
        Me.txtDate_VTG.Value = Format(.Range(102).Value, "dd-mmm-yyyy")
        Me.txtDate_Company.Value = Format(.Range(103).Value, "dd-mmm-yyyy")
        Me.txtDate_Finalised.Value = Format(.Range(104).Value, "dd-mmm-yyyy")
        Me.txtReminder.Value = .Range(105).Value
    End With
    
    'Access version control
    Call LogLastAccess
    
    'Depress and make toggle green on nav bar
    Me.tglCTRA.Value = True
    Me.tglCTRA.BackColor = vbGreen
    
    'Run date validation on data entered
    Call txtDate_RGC_AfterUpdate
    Call txtDate_UWA_AfterUpdate
    Call txtDate_Finance_AfterUpdate
    Call txtDate_COO_AfterUpdate
    Call txtDate_VTG_AfterUpdate
    Call txtDate_Company_AfterUpdate
    Call txtDate_Finalised_AfterUpdate
    
End Sub

Private Sub txtDate_RGC_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtDate_RGC.Value)
    
    'Display error message
    Me.errDate_RGC.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtDate_RGC.Value) Then
        Me.txtDate_RGC.Value = Format(Me.txtDate_RGC.Value, "dd-mmm-yyyy")
    End If
    
End Sub

Private Sub txtDate_UWA_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtDate_UWA.Value)
    
    'Display error message
    Me.errDate_UWA.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtDate_UWA.Value) Then
        Me.txtDate_UWA.Value = Format(Me.txtDate_UWA.Value, "dd-mmm-yyyy")
    End If
    
End Sub

Private Sub txtDate_Finance_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtDate_Finance.Value)
    
    'Display error message
    Me.errDate_Finance.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtDate_Finance.Value) Then
        Me.txtDate_Finance.Value = Format(Me.txtDate_Finance.Value, "dd-mmm-yyyy")
    End If
    
End Sub

Private Sub txtDate_COO_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtDate_COO.Value, Me.txtDate_Finance.Value, _
            "Date entered earlier than" & Chr(10) & "Finance Sign-off")
    
    'Display error message
    Me.errDate_COO.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtDate_COO.Value) Then
        Me.txtDate_COO.Value = Format(Me.txtDate_COO.Value, "dd-mmm-yyyy")
    End If
    
End Sub

Private Sub txtDate_VTG_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtDate_VTG.Value, Me.txtDate_COO.Value, _
            "Date entered earlier than" & Chr(10) & "COO sign-off")
    
    'Display error message
    Me.errDate_VTG.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtDate_VTG.Value) Then
        Me.txtDate_VTG.Value = Format(Me.txtDate_VTG.Value, "dd-mmm-yyyy")
    End If
    
End Sub

Private Sub txtDate_Company_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    Dim d1 As Variant
    Dim d2 As Variant
    
    err = Date_Validation(Me.txtDate_Company.Value, Me.txtDate_VTG.Value, _
            "Date entered earlier than" & Chr(10) & "VTG Sign-off")
    
    'Display error message
    Me.errDate_Company.Caption = err
    
    'Change date format displayed
    If err = vbNullString Then
        Me.txtDate_Company.Value = Format(Me.txtDate_Company.Value, "dd-mmm-yyyy")
    End If
    
End Sub

Private Sub txtDate_Finalised_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    Dim d1 As Variant
    Dim d2 As Variant
    
    err = Date_Validation(Me.txtDate_Finalised.Value, Me.txtDate_Company.Value, _
            "Date entered earlier than" & Chr(10) & "Company submission")
    
    'Display error message
    Me.errDate_Finalised.Caption = err
    
    'Change date format displayed
    If err = vbNullString Then
        Me.txtDate_Finalised.Value = Format(Me.txtDate_Finalised.Value, "dd-mmm-yyyy")
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
        
        .Range(98) = String_to_Date(Me.txtDate_RGC.Value)
        .Range(99) = String_to_Date(Me.txtDate_UWA.Value)
        .Range(100) = String_to_Date(Me.txtDate_Finance.Value)
        .Range(101) = String_to_Date(Me.txtDate_COO.Value)
        .Range(102) = String_to_Date(Me.txtDate_VTG.Value)
        .Range(103) = String_to_Date(Me.txtDate_Company.Value)
        .Range(104) = String_to_Date(Me.txtDate_Finalised.Value)
        .Range(105) = Me.txtReminder.Value
        
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

