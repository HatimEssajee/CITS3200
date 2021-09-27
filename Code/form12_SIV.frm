VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form12_SIV 
   Caption         =   "Financial Disclosure"
   ClientHeight    =   6684
   ClientLeft      =   -444
   ClientTop       =   -1992
   ClientWidth     =   10980
   OleObjectBlob   =   "form12_SIV.frx":0000
End
Attribute VB_Name = "form12_SIV"
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
        Me.txtStudyName.Value = .Range(9).Value
        Me.txtSIV_Date.Value = Format(.Range(125).Value, "dd-mmm-yyyy")
        Me.txtReminder.Value = .Range(126).Value
    End With
    
    'Access version control
    Call LogLastAccess
    
    'Depress and make toggle green on nav bar
    Me.tglSIV.Value = True
    Me.tglSIV.BackColor = vbGreen
    
    'Run date validation on data entered
    Call txtSIV_date_AfterUpdate
    
End Sub
Private Sub txtSIV_date_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtSIV_Date.Value)
    
    'Display error message
    Me.errSIV_Date.Caption = err
    
    'Change date format displayed
    If err = vbNullString Then
        Me.txtSIV_Date.Value = Format(Me.txtSIV_Date.Value, "dd-mmm-yyyy")
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
        
        .Range(125) = String_to_Date(Me.txtSIV_Date.Value)
        .Range(126) = Me.txtReminder.Value
        
        'Update version control
        .Range(127) = Now
        .Range(128) = Username
        
        'Change study status based on SIV date saved
        
        'Check if all states other than SIV complete
        If .Range(156) And Me.txtSIV_Date.Value <> vbNullString Then
            
            If String_to_Date(Me.txtSIV_Date.Value) > Now And .Range(7).Value = "Commenced" Then
                
                .Range(7) = "Pre-commencement"
                
                'Update version control
                .Range(14) = .Range(127).Value
                .Range(15) = .Range(128).Value
            
            ElseIf String_to_Date(Me.txtSIV_Date.Value) < Now And .Range(7).Value = "Pre-commencement" Then
                
                .Range(7) = "Commenced"
                
                'Update version control
                .Range(14) = .Range(127).Value
                .Range(15) = .Range(128).Value
            End If
        End If
        
        'Apply completion status
        Call Fill_Completion_Status
        DoEvents
        
    End With
    
    'Access version control
    Call LogLastAccess
    
    Call UserForm_Initialize

End Sub


'----------------- Navigation section Toggles ----------------

Private Sub tglNav_Click()
    'PURPOSE: Closes current form and open Nav form
    Unload form12_SIV
    
    form00_Nav.Show False
End Sub

Private Sub tglCDA_Click()
    'PURPOSE: Closes current form and open CDA form
    Unload form12_SIV
    
    form02_CDA.Show False
End Sub

Private Sub tglFS_Click()
    'PURPOSE: Closes current form and open Feasibility form
    Unload form12_SIV
    
    form03_FS.Show False
End Sub

Private Sub tglSiteSelect_Click()
    'PURPOSE: Closes current form and open Site Select form
    Unload form12_SIV
    
    form04_SiteSelect.Show False
End Sub

Private Sub tglRecruit_Click()
    'PURPOSE: Closes current form and open Recruitment form
    Unload form12_SIV
    
    form05_Recruitment.Show False
End Sub

Private Sub tglEthics_Click()
    'PURPOSE: Closes current form and open Ethics form
    Unload form12_SIV
    
    form06_Ethics.Show False
End Sub

Private Sub tglGov_Click()
    'PURPOSE: Closes current form and open Governance form
    Unload form12_SIV
    
    form07_Governance.Show False
End Sub

Private Sub tglBudget_Click()
    'PURPOSE: Closes current form and open Budget form
    Unload form12_SIV
    
    form08_Budget.Show False
End Sub

Private Sub tglIndemnity_Click()
    'PURPOSE: Closes current form and open Indemnity form
    Unload form12_SIV
    
    form09_Indemnity.Show False
End Sub

Private Sub tglCTRA_Click()
    'PURPOSE: Closes current form and open CTRA form
    Unload form12_SIV
    
    form10_CTRA.Show False
End Sub

Private Sub tglFinDisc_Click()
    'PURPOSE: Closes current form and open Fin. Disc. form
    Unload form12_SIV
    
    form11_FinDisc.Show False
End Sub

Private Sub tglStudyDetail_Click()
    'PURPOSE: Closes current form and open Study Detail form
    Unload form12_SIV
    
    form01_StudyDetail.Show False
End Sub

