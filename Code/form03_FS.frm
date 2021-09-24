VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form03_FS 
   Caption         =   "CDA & Feasibility"
   ClientHeight    =   8292.001
   ClientLeft      =   -396
   ClientTop       =   -1752
   ClientWidth     =   18624
   OleObjectBlob   =   "form03_FS.frx":0000
End
Attribute VB_Name = "form03_FS"
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
        Me.txtFS_Recv.Value = Format(.Range(24).Value, "dd-mmm-yyyy")
        Me.txtFS_Comp.Value = Format(.Range(25).Value, "dd-mmm-yyyy")
        Me.txtFS_Initials.Value = .Range(26).Value
        
        Me.txtReminder.Value = .Range(27).Value
    End With
    
    'Access version control
    Call LogLastAccess
    
    'Depress and make toggle green on nav bar
    Me.tglFS.Value = True
    Me.tglFS.BackColor = vbGreen
    
    'Run date validation on data entered
    Call txtFS_Recv_AfterUpdate
    Call txtFS_Comp_AfterUpdate
    
End Sub

Private Sub txtFS_Recv_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtFS_Recv.Value)
    
    'Display error message
    Me.errFS_Recv.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtFS_Recv.Value) Then
        Me.txtFS_Recv.Value = Format(Me.txtFS_Recv.Value, "dd-mmm-yyyy")
    End If
    
End Sub

Private Sub txtFS_Comp_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtFS_Comp.Value, Me.txtFS_Recv.Value, _
            "Date entered earlier than date" & Chr(10) & "received")

    'Display error message
    Me.errFS_Comp.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtFS_Comp.Value) Then
        Me.txtFS_Comp.Value = Format(Me.txtFS_Comp.Value, "dd-mmm-yyyy")
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
        
        .Range(24) = String_to_Date(Me.txtFS_Recv.Value)
        .Range(25) = String_to_Date(Me.txtFS_Comp.Value)
        .Range(26) = Me.txtFS_Initials.Value
        
        .Range(27) = Me.txtReminder.Value
        
        'Update version control
        .Range(28) = Now
        .Range(29) = Username
        
        'Apply completion status
        If .Range(25).Value = vbNullString Then
            .Range(131).Value = vbNullString
        Else
            .Range(131).Value = IsDate(.Range(25).Value)
        End If
        
    End With
    
    'Access version control
    Call LogLastAccess
    
    Call UserForm_Initialize
    
End Sub

'----------------- Navigation section Toggles ----------------

Private Sub tglNav_Click()
    'PURPOSE: Closes current form and open Nav form
    Unload form03_FS
    
    form00_Nav.Show False
End Sub

Private Sub tglStudyDetail_Click()
    'PURPOSE: Closes current form and open Study Details form
    Unload form03_FS
    
    form01_StudyDetail.Show False
End Sub

Private Sub tglCDA_Click()
    'PURPOSE: Closes current form and open CDA form
    Unload form03_FS
    
    form02_CDA.Show False
End Sub

Private Sub tglSiteSelect_Click()
    'PURPOSE: Closes current form and open Site Select form
    Unload form03_FS
    
    form04_SiteSelect.Show False
End Sub

Private Sub tglRecruit_Click()
    'PURPOSE: Closes current form and open Recruitment form
    Unload form03_FS
    
    form05_Recruitment.Show False
End Sub

Private Sub tglEthics_Click()
    'PURPOSE: Closes current form and open Ethics form
    Unload form03_FS
    
    form06_Ethics.Show False
End Sub

Private Sub tglGov_Click()
    'PURPOSE: Closes current form and open Governance form
    Unload form03_FS
    
    form07_Governance.Show False
End Sub

Private Sub tglBudget_Click()
    'PURPOSE: Closes current form and open Budget form
    Unload form03_FS
    
    form08_Budget.Show False
End Sub

Private Sub tglIndemnity_Click()
    'PURPOSE: Closes current form and open Indemnity form
    Unload form03_FS
    
    form09_Indemnity.Show False
End Sub

Private Sub tglCTRA_Click()
    'PURPOSE: Closes current form and open CTRA form
    Unload form03_FS
    
    form10_CTRA.Show False
End Sub

Private Sub tglFinDisc_Click()
    'PURPOSE: Closes current form and open Fin. Disc. form
    Unload form03_FS
    
    form11_FinDisc.Show False
End Sub

Private Sub tglSIV_Click()
    'PURPOSE: Closes current form and open SIV form
    Unload form03_FS
    
    form12_SIV.Show False
End Sub


