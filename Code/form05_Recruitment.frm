VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form05_Recruitment 
   Caption         =   "Recruitment Plan"
   ClientHeight    =   7992
   ClientLeft      =   -504
   ClientTop       =   -2160
   ClientWidth     =   14004
   OleObjectBlob   =   "form05_Recruitment.frx":0000
End
Attribute VB_Name = "form05_Recruitment"
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
    
    'Turn off Settings to speed up
    'source: https://www.automateexcel.com/vba/turn-off-screen-updating/
    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayStatusBar = False
    Application.EnableEvents = False
    
    'Erase and initialise arrays
    ReDim OldValues(1 To 2)
    ReDim NxtOldValues(1 To 2)
                    
    'Clear user form
    'SOURCE: https://www.mrexcel.com/board/threads/loop-through-controls-on-a-userform.427103/
    For Each ctrl In Me.Controls
        Select Case True
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
        Me.txtDate_Plan.Value = ReadDate(.Range(38).Value)
        Me.txtReminder.Value = .Range(39).Value
    End With
    
    'Populate Old Values Array - for undo
    OldValues(1) = String_to_Date(Me.txtDate_Plan)
    OldValues(2) = Me.txtReminder
    
    'Initialize NxtOldValues
    NxtOldValues = OldValues
    
    'Access version control
    Call LogLastAccess
    
    'Depress and make toggle green on nav bar
    Me.tglRecruit.Value = True
    Me.tglRecruit.BackColor = vbGreen
    
    'Allocate tick box values
    Me.cbSaveonUnload.Value = SAG_Tick
    
    'Run date validation on data entered
    Call txtDate_Plan_AfterUpdate
    
    'Reinstate Settings
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayStatusBar = True
    Application.EnableEvents = True
    
End Sub

Private Sub txtDate_Plan_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtDate_Plan.Value)
    
    'Display error message
    Me.errDate_Plan.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtDate_Plan.Value) Then
        Me.txtDate_Plan.Value = ReadDate(CStr(Me.txtDate_Plan.Value))
    End If
    
End Sub

Private Sub cmdUndo_Click()
    'PURPOSE: Recall values read from register table when the form was loaded initially
    
    Me.txtDate_Plan.Value = ReadDate(CStr(OldValues(1)))
    Me.txtReminder.Value = OldValues(2)
    
End Sub

Private Sub cmdRedo_Click()
    'PURPOSE: Recall values replaced by undo
    
    Me.txtDate_Plan.Value = ReadDate(CStr(NxtOldValues(1)))
    Me.txtReminder.Value = NxtOldValues(2)
    
End Sub

Private Sub cmdClose_Click()
    'PURPOSE: Closes current form
    
    Unload Me
    
End Sub

Private Sub cmdEdit_Click()
    'PURPOSE: Apply changes into Register table when edit button clicked
    
    'Overwrite prev. old values with new backup values
    OldValues = NxtOldValues
    
    'Apply changes
    Call UpdateRegister
    DoEvents
    
    'Access version control
    Call LogLastAccess
    
    'Run date validation on data entered
    Call txtDate_Plan_AfterUpdate
    
End Sub

Private Sub Userform_Terminate()
    'PURPOSE: Update register when unloaded
        
    If cbSaveonUnload.Value Then
        'Apply changes
        Call UpdateRegister
        DoEvents
    End If
    
    'Access version control
    Call LogLastAccess
    
End Sub

Private Sub UpdateRegister()
    'PURPOSE: Apply changes into Register table
    Dim ReadRow(1 To 2) As Variant
    
    'Turn off Settings to speed up
    'source: https://www.automateexcel.com/vba/turn-off-screen-updating/
    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayStatusBar = False
    Application.EnableEvents = False
    
    With RegTable.ListRows(RowIndex)
        
        'Populate ReadRow Array - faster than double transpose
        ReadRow(1) = String_to_Date(Me.txtDate_Plan)
        ReadRow(2) = Me.txtReminder
        
        'Write to Register Table
        .Range(38) = ReadRow(1)
        .Range(39) = ReadRow(2)
         
        'Store next old values
        NxtOldValues = ReadRow
        
        'Check if values changed
        If Not ArraysSame(ReadRow, OldValues) Then
            'Update version control
            .Range(40) = Now
            .Range(41) = Username
            
            'Apply completion status
            Call Fill_Completion_Status
            DoEvents
        End If
        
        'Clear array elements
        Erase ReadRow
        
    End With
    
    'Reinstate Settings
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayStatusBar = True
    Application.EnableEvents = True
    
End Sub

Private Sub Fill_Completion_Status()

    'PURPOSE: Evaluate entry completion status
    
    Dim db As Range
    Dim ReadRow As Variant
    Dim i As Integer, cntTrue As Integer, cntEmpty As Integer
    
    'Turn off Settings to speed up
    'source: https://www.automateexcel.com/vba/turn-off-screen-updating/
    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayStatusBar = False
    Application.EnableEvents = False
    
    'Exit if register is empty
    If RegTable.DataBodyRange Is Nothing Then
        GoTo ErrHandler
    End If
    
    Set db = RegTable.DataBodyRange
    
    ReadRow = db.Cells(RowIndex, 38)
    
    'Apply test for completeness
    If ReadRow <> vbNullString Then
        ReadRow = IsDate(Format(ReadRow, "dd-mmm-yyyy"))
    End If
    
    'Completion status
    
    'Recruitment
    'Criteria - has to be date
    db.Cells(RowIndex, 133) = ReadRow
   
ErrHandler:
    'Reinstate Settings
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayStatusBar = True
    Application.EnableEvents = True
    
End Sub

Private Sub cbSaveonUnload_Click()
    'PURPOSE: Change value of SAG_Tick variable
    SAG_Tick = Me.cbSaveonUnload.Value
    
    If SAG_Tick Then
        Me.cbSaveonUnload.Caption = "Save via Navigation"
    Else
        Me.cbSaveonUnload.Caption = "Save via Button"
    End If
    
End Sub


'----------------- Navigation section Toggles ----------------

Private Sub tglNav_Click()
    'PURPOSE: Closes current form and open Nav form
    Unload form05_Recruitment
    
    form00_Nav.show False
End Sub

Private Sub tglCDA_Click()
    'PURPOSE: Closes current form and open CDA form
    Unload form05_Recruitment
    
    form02_CDA.show False
End Sub

Private Sub tglFS_Click()
    'PURPOSE: Closes current form and open Feasibility form
    Unload form05_Recruitment
    
    form03_FS.show False
End Sub

Private Sub tglSiteSelect_Click()
    'PURPOSE: Closes current form and open Site Select form
    Unload form05_Recruitment
    
    form04_SiteSelect.show False
End Sub

Private Sub tglStudyDetail_Click()
    'PURPOSE: Closes current form and open Study Detail form
    Unload form05_Recruitment
    
    form01_StudyDetail.show False
End Sub

Private Sub tglEthics_Click()
    'PURPOSE: Closes current form and open Ethics form
    Unload form05_Recruitment
    
    form06_Ethics.show False
    form06_Ethics.multiEthics.Value = 0
End Sub

Private Sub tglGov_Click()
    'PURPOSE: Closes current form and open Governance form
    Unload form05_Recruitment
    
    form07_Governance.show False
    form07_Governance.multiGov.Value = 0
End Sub

Private Sub tglBudget_Click()
    'PURPOSE: Closes current form and open Budget form
    Unload form05_Recruitment
    
    form08_Budget.show False
    form08_Budget.multiBudget.Value = 0
End Sub

Private Sub tglIndemnity_Click()
    'PURPOSE: Closes current form and open Indemnity form
    Unload form05_Recruitment
    
    form09_Indemnity.show False
End Sub

Private Sub tglCTRA_Click()
    'PURPOSE: Closes current form and open CTRA form
    Unload form05_Recruitment
    
    form10_CTRA.show False
End Sub

Private Sub tglFinDisc_Click()
    'PURPOSE: Closes current form and open Fin. Disc. form
    Unload form05_Recruitment
    
    form11_FinDisc.show False
End Sub

Private Sub tglSIV_Click()
    'PURPOSE: Closes current form and open SIV form
    Unload form05_Recruitment
    
    form12_SIV.show False
End Sub

Private Function ArraysSame(ArrX As Variant, ArrY As Variant) As Boolean
    'PURPOSE: Compare values of two 1D arrays
    
    Dim Check As Boolean
    Dim Upper As Long, i As Long
    
    Check = True
    Upper = UBound(ArrX)
    
    'Shift upper bound to smaller array
    If UBound(ArrX) >= UBound(ArrY) Then
        Upper = UBound(ArrY)
    End If
    
    For i = LBound(ArrX) To Upper
        If ArrX(i) <> ArrY(i) Then
            Check = False
            Exit For
        End If
    Next i
    
    ArraysSame = Check
End Function
