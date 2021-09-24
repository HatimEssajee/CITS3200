VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form02_CDA 
   Caption         =   "CDA & Feasibility"
   ClientHeight    =   8424.001
   ClientLeft      =   -396
   ClientTop       =   -1752
   ClientWidth     =   15804
   OleObjectBlob   =   "form02_CDA.frx":0000
End
Attribute VB_Name = "form02_CDA"
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
    
    'Read information from register table
    With RegTable.ListRows(RowIndex)
        Me.txtStudyName.Value = .Range(9).Value
        Me.txtCDA_Recv_Sponsor.Value = Format(.Range(16).Value, "dd-mmm-yyyy")
        Me.txtCDA_Sent_Contracts.Value = Format(.Range(17).Value, "dd-mmm-yyyy")
        Me.txtCDA_Recv_Contracts.Value = Format(.Range(18).Value, "dd-mmm-yyyy")
        Me.txtCDA_Sent_Sponsor.Value = Format(.Range(19).Value, "dd-mmm-yyyy")
        Me.txtCDA_Finalised.Value = Format(.Range(20).Value, "dd-mmm-yyyy")
        
        Me.txtReminder.Value = .Range(21).Value
    End With
    
    'Access version control
    Call LogLastAccess
    
    'Depress and make toggle green on nav bar
    Me.tglCDA.Value = True
    Me.tglCDA.BackColor = vbGreen
    
    'Run date validation on data entered
    Call txtCDA_Recv_Sponsor_AfterUpdate
    Call txtCDA_Sent_Contracts_AfterUpdate
    Call txtCDA_Recv_Contracts_AfterUpdate
    Call txtCDA_Sent_Sponsor_AfterUpdate
    Call txtCDA_Finalised_AfterUpdate
    
End Sub

Private Sub txtCDA_Recv_Sponsor_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtCDA_Recv_Sponsor.Value)
    
    'Display error message
    Me.errCDA_Recv_Sponsor.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtCDA_Recv_Sponsor.Value) Then
        Me.txtCDA_Recv_Sponsor.Value = Format(Me.txtCDA_Recv_Sponsor.Value, "dd-mmm-yyyy")
    End If
    
End Sub

Private Sub txtCDA_Sent_Contracts_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtCDA_Sent_Contracts.Value, Me.txtCDA_Recv_Sponsor.Value, _
            "Date entered earlier than date" & Chr(10) & "received from Sponsor")

    'Display error message
    Me.errCDA_Sent_Contracts.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtCDA_Sent_Contracts.Value) Then
        Me.txtCDA_Sent_Contracts.Value = Format(Me.txtCDA_Sent_Contracts.Value, "dd-mmm-yyyy")
    End If
     
End Sub

Private Sub txtCDA_Recv_Contracts_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtCDA_Recv_Contracts.Value, Me.txtCDA_Sent_Contracts.Value, _
            "Date entered earlier than date" & Chr(10) & "sent to Contracts")

    'Display error message
    Me.errCDA_Recv_Contracts.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtCDA_Recv_Contracts.Value) Then
        Me.txtCDA_Recv_Contracts.Value = Format(Me.txtCDA_Recv_Contracts.Value, "dd-mmm-yyyy")
    End If
     
End Sub

Private Sub txtCDA_Sent_Sponsor_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtCDA_Sent_Sponsor.Value, Me.txtCDA_Recv_Contracts.Value, _
            "Date entered earlier than date" & Chr(10) & "received from Contracts")

    'Display error message
    Me.errCDA_Sent_Sponsor.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtCDA_Sent_Sponsor.Value) Then
        Me.txtCDA_Sent_Sponsor.Value = Format(Me.txtCDA_Sent_Sponsor.Value, "dd-mmm-yyyy")
    End If
    
End Sub

Private Sub txtCDA_Finalised_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtCDA_Finalised.Value, Me.txtCDA_Sent_Sponsor.Value, _
            "Date entered earlier than date" & Chr(10) & "sent to Sponsor")

    'Display error message
    Me.errCDA_Finalised.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtCDA_Finalised.Value) Then
        Me.txtCDA_Finalised.Value = Format(Me.txtCDA_Finalised.Value, "dd-mmm-yyyy")
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
        
        .Range(16) = String_to_Date(Me.txtCDA_Recv_Sponsor.Value)
        .Range(17) = String_to_Date(Me.txtCDA_Sent_Contracts.Value)
        .Range(18) = String_to_Date(Me.txtCDA_Recv_Contracts.Value)
        .Range(19) = String_to_Date(Me.txtCDA_Sent_Sponsor.Value)
        .Range(20) = String_to_Date(Me.txtCDA_Finalised.Value)
        
        .Range(21) = Me.txtReminder.Value
        
        'Update version control
        .Range(22) = Now
        .Range(23) = Username
        
        'Apply completion status
        If .Range(20).Value = vbNullString Then
            .Range(130).Value = vbNullString
        Else
            .Range(130).Value = IsDate(.Range(20).Value)
        End If
        
    End With
    
    'Access version control
    Call LogLastAccess
    
    Call UserForm_Initialize
    
End Sub

'----------------- Navigation section Toggles ----------------

Private Sub tglNav_Click()
    'PURPOSE: Closes current form and open Nav form
    Unload form02_CDA
    
    form00_Nav.Show False
End Sub

Private Sub tglStudyDetail_Click()
    'PURPOSE: Closes current form and open Study Details form
    Unload form02_CDA
    
    form01_StudyDetail.Show False
End Sub

Private Sub tglFS_Click()
    'PURPOSE: Closes current form and open Feasibility form
    Unload form02_CDA
    
    form03_FS.Show False
End Sub

Private Sub tglSiteSelect_Click()
    'PURPOSE: Closes current form and open Site Select form
    Unload form02_CDA
    
    form04_SiteSelect.Show False
End Sub

Private Sub tglRecruit_Click()
    'PURPOSE: Closes current form and open Recruitment form
    Unload form02_CDA
    
    form05_Recruitment.Show False
End Sub

Private Sub tglEthics_Click()
    'PURPOSE: Closes current form and open Ethics form
    Unload form02_CDA
    
    form06_Ethics.Show False
End Sub

Private Sub tglGov_Click()
    'PURPOSE: Closes current form and open Governance form
    Unload form02_CDA
    
    form07_Governance.Show False
End Sub

Private Sub tglBudget_Click()
    'PURPOSE: Closes current form and open Budget form
    Unload form02_CDA
    
    form08_Budget.Show False
End Sub

Private Sub tglIndemnity_Click()
    'PURPOSE: Closes current form and open Indemnity form
    Unload form02_CDA
    
    form09_Indemnity.Show False
End Sub

Private Sub tglCTRA_Click()
    'PURPOSE: Closes current form and open CTRA form
    Unload form02_CDA
    
    form10_CTRA.Show False
End Sub

Private Sub tglFinDisc_Click()
    'PURPOSE: Closes current form and open Fin. Disc. form
    Unload form02_CDA
    
    form11_FinDisc.Show False
End Sub

Private Sub tglSIV_Click()
    'PURPOSE: Closes current form and open SIV form
    Unload form02_CDA
    
    form12_SIV.Show False
End Sub

