VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form03_FS 
   Caption         =   "Feasibility"
   ClientHeight    =   6624
   ClientLeft      =   -420
   ClientTop       =   -1836
   ClientWidth     =   14892
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
    
    'Turn off Settings to speed up
    'source: https://www.automateexcel.com/vba/turn-off-screen-updating/
    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayStatusBar = False
    Application.EnableEvents = False
    
    'Erase and initialise arrays
    ReDim OldValues(1 To 4)
    ReDim NxtOldValues(1 To 4)
                    
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
        Me.txtFS_Recv.Value = ReadDate(.Range(24).Value)
        Me.txtFS_Comp.Value = ReadDate(.Range(25).Value)
        Me.txtFS_Initials.Value = .Range(26).Value
        
        Me.txtReminder.Value = .Range(27).Value
    End With
    
    'Populate Old Values Array - for undo
    OldValues(1) = String_to_Date(Me.txtFS_Recv.Value)
    OldValues(2) = String_to_Date(Me.txtFS_Comp.Value)
    OldValues(3) = Me.txtFS_Initials.Value
    OldValues(4) = Me.txtReminder.Value
    
    'Initialize NxtOldValues
    NxtOldValues = OldValues
    
    'Access version control
    Call LogLastAccess
    
    'Depress and make toggle green on nav bar
    Me.tglFS.Value = True
    Me.tglFS.BackColor = vbGreen
    
    'Allocate tick box values
    Me.cbSaveonUnload.Value = SAG_Tick
    
    'Run date validation on data entered
    Call txtFS_Recv_AfterUpdate
    Call txtFS_Comp_AfterUpdate
    
    'Reinstate Settings
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayStatusBar = True
    Application.EnableEvents = True
    
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

Private Sub cmdUndo_Click()
    'PURPOSE: Recall values read from register table when the form was loaded initially
    
    Me.txtFS_Recv.Value = ReadDate(CStr(OldValues(1)))
    Me.txtFS_Comp.Value = ReadDate(CStr(OldValues(2)))
    Me.txtFS_Initials.Value = OldValues(3)
    
    Me.txtReminder.Value = OldValues(4)
    
End Sub

Private Sub cmdRedo_Click()
    'PURPOSE: Recall values replaced by undo
    
    Me.txtFS_Recv.Value = ReadDate(CStr(NxtOldValues(1)))
    Me.txtFS_Comp.Value = ReadDate(CStr(NxtOldValues(2)))
    Me.txtFS_Initials.Value = NxtOldValues(3)
    
    Me.txtReminder.Value = NxtOldValues(4)

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
    Call txtFS_Recv_AfterUpdate
    Call txtFS_Comp_AfterUpdate
    
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
    
    Dim ReadRow(1 To 4) As Variant
    
    'Turn off Settings to speed up
    'source: https://www.automateexcel.com/vba/turn-off-screen-updating/
    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayStatusBar = False
    Application.EnableEvents = False
    
    With RegTable.ListRows(RowIndex)
        
        'Populate ReadRow Array - faster than double transpose
        ReadRow(1) = String_to_Date(Me.txtFS_Recv.Value)
        ReadRow(2) = String_to_Date(Me.txtFS_Comp.Value)
        ReadRow(3) = Me.txtFS_Initials.Value
        ReadRow(4) = Me.txtReminder.Value
        
        'Write to Register Table
        .Range(24) = ReadRow(1)
        .Range(25) = ReadRow(2)
        .Range(26) = ReadRow(3)
        .Range(27) = ReadRow(4)
        
        'Store next old values
        NxtOldValues = ReadRow
        
        'Check if values changed
        If Not ArraysSame(ReadRow, OldValues) Then
            'Update version control
            .Range(28) = Now
            .Range(29) = Username
            
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
    
    'Tranpose twice to get 1D Array
    ReadRow = Application.Transpose(Application.Transpose(Range(db.Cells(RowIndex, 24), db.Cells(RowIndex, 26))))
                   
    'Apply correct test on each field
    For i = LBound(ReadRow) To UBound(ReadRow)
        If ReadRow(i) <> vbNullString Then
    
            Select Case Correct(i + 16)
                Case 0
                    ReadRow(i) = "Skip"
                Case 1
                    ReadRow(i) = Not (IsEmpty(ReadRow(i)))
                Case 2
                    ReadRow(i) = WorksheetFunction.IsText(ReadRow(i))
                Case 3
                    ReadRow(i) = IsDate(Format(ReadRow(i), "dd-mmm-yyyy"))
            End Select
            
        End If
    Next i
    
    'Completion status
    
    'Feasibility
    'Criteria - all fields filled with dates
    cntTrue = 0
    cntEmpty = 0
    For i = 1 To 3
        If ReadRow(i) = vbNullString Then
            cntEmpty = cntEmpty + 1
        ElseIf ReadRow(i) Then
            cntTrue = cntTrue + 1
        End If
    Next i
    
    If cntEmpty = 3 Then
        db.Cells(RowIndex, 131) = vbNullString
    ElseIf cntTrue = 3 Then
        db.Cells(RowIndex, 131) = True
    Else
        db.Cells(RowIndex, 131) = False
    End If
    
ErrHandler:
    'Clear array
    Erase ReadRow
    
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
    Unload form03_FS
    
    form00_Nav.show False
End Sub

Private Sub tglStudyDetail_Click()
    'PURPOSE: Closes current form and open Study Details form
    Unload form03_FS
    
    form01_StudyDetail.show False
End Sub

Private Sub tglCDA_Click()
    'PURPOSE: Closes current form and open CDA form
    Unload form03_FS
    
    form02_CDA.show False
End Sub

Private Sub tglSiteSelect_Click()
    'PURPOSE: Closes current form and open Site Select form
    Unload form03_FS
    
    form04_SiteSelect.show False
End Sub

Private Sub tglRecruit_Click()
    'PURPOSE: Closes current form and open Recruitment form
    Unload form03_FS
    
    form05_Recruitment.show False
End Sub

Private Sub tglEthics_Click()
    'PURPOSE: Closes current form and open Ethics form
    Unload form03_FS
    
    form06_Ethics.show False
    form06_Ethics.multiEthics.Value = 0
End Sub

Private Sub tglGov_Click()
    'PURPOSE: Closes current form and open Governance form
    Unload form03_FS
    
    form07_Governance.show False
    form07_Governance.multiGov.Value = 0
End Sub

Private Sub tglBudget_Click()
    'PURPOSE: Closes current form and open Budget form
    Unload form03_FS
    
    form08_Budget.show False
    form08_Budget.multiBudget.Value = 0
End Sub

Private Sub tglIndemnity_Click()
    'PURPOSE: Closes current form and open Indemnity form
    Unload form03_FS
    
    form09_Indemnity.show False
End Sub

Private Sub tglCTRA_Click()
    'PURPOSE: Closes current form and open CTRA form
    Unload form03_FS
    
    form10_CTRA.show False
End Sub

Private Sub tglFinDisc_Click()
    'PURPOSE: Closes current form and open Fin. Disc. form
    Unload form03_FS
    
    form11_FinDisc.show False
End Sub

Private Sub tglSIV_Click()
    'PURPOSE: Closes current form and open SIV form
    Unload form03_FS
    
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
