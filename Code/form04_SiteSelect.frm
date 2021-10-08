VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form04_SiteSelect 
   Caption         =   "Site Selection"
   ClientHeight    =   6492
   ClientLeft      =   -510
   ClientTop       =   -2130
   ClientWidth     =   12600
   OleObjectBlob   =   "form04_SiteSelect.frx":0000
End
Attribute VB_Name = "form04_SiteSelect"
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
    Dim cboList_StudyStatus As Variant, item As Variant
    
    'Turn off Settings to speed up
    'source: https://www.automateexcel.com/vba/turn-off-screen-updating/
    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayStatusBar = False
    Application.EnableEvents = False
    
    cboList_StudyStatus = Array("On-site", "Virtual")
    
    'Erase and initialise arrays
    ReDim OldValues(1 To 6)
    ReDim NxtOldValues(1 To 6)
                    
    'Clear user form
    'source: https://www.mrexcel.com/board/threads/loop-through-controls-on-a-userform.427103/
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
    
    'Fill combo box for study status
    For Each item In cboList_StudyStatus
        cboPrestudy_Type.AddItem item
        cboValidation_Type.AddItem item
    Next item
    
    'Read information from register table
    With RegTable.ListRows(RowIndex)
        Me.txtStudyName = .Range(9).Value
        Me.txtPrestudy_Date.Value = ReadDate(.Range(30).Value)
        Me.cboPrestudy_Type.Value = .Range(31).Value
        Me.txtValidation_Date.Value = ReadDate(.Range(32).Value)
        Me.cboValidation_Type.Value = .Range(33).Value
        Me.txtSiteSelect.Value = ReadDate(.Range(34).Value)
        
        Me.txtReminder.Value = .Range(35).Value
    End With
    
    'Populate Old Values Array - for undo
    OldValues(1) = String_to_Date(Me.txtPrestudy_Date.Value)
    OldValues(2) = Me.cboPrestudy_Type.Value
    OldValues(3) = String_to_Date(Me.txtValidation_Date.Value)
    OldValues(4) = Me.cboValidation_Type.Value
    OldValues(5) = String_to_Date(Me.txtSiteSelect.Value)
    OldValues(6) = Me.txtReminder.Value
    
    'Initialize NxtOldValues
    NxtOldValues = OldValues
    
    'Access version control
    Call LogLastAccess
    
    'Depress and make toggle green on nav bar
    Me.tglSiteSelect.Value = True
    Me.tglSiteSelect.BackColor = vbGreen
    
    'Allocate tick box values
    Me.cbSaveonUnload.Value = SAG_Tick
    
    'Run date validation on data entered
    Call txtPrestudy_Date_AfterUpdate
    Call txtValidation_Date_AfterUpdate
    Call txtSiteSelect_AfterUpdate
    
    'Reinstate Settings
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayStatusBar = True
    Application.EnableEvents = True
    
End Sub

Private Sub txtPrestudy_Date_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtPrestudy_Date.Value)
    
    'Display error message
    Me.errPrestudy_Date.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtPrestudy_Date.Value) Then
        Me.txtPrestudy_Date.Value = Format(Me.txtPrestudy_Date.Value, "dd-mmm-yyyy")
    End If
    
End Sub

Private Sub txtValidation_Date_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtValidation_Date.Value, Me.txtPrestudy_Date.Value, _
            "Date entered earlier than date of" & Chr(10) & "Pre-study visit")

    'Display error message
    Me.errValidation_Date.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtValidation_Date.Value) Then
        Me.txtValidation_Date.Value = Format(Me.txtValidation_Date.Value, "dd-mmm-yyyy")
    End If
     
End Sub

Private Sub txtSiteSelect_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtSiteSelect.Value, Me.txtValidation_Date.Value, _
            "Date entered earlier than date of" & Chr(10) & "Validation visit")

    'Display error message
    Me.errSiteSelect.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtSiteSelect.Value) Then
        Me.txtSiteSelect.Value = Format(Me.txtSiteSelect.Value, "dd-mmm-yyyy")
    End If
     
End Sub

Private Sub cmdUndo_Click()
    'PURPOSE: Recall values read from register table when the form was loaded initially
    
    Me.txtPrestudy_Date.Value = ReadDate(CStr(OldValues(1)))
    Me.cboPrestudy_Type.Value = OldValues(2)
    Me.txtValidation_Date.Value = ReadDate(CStr(OldValues(3)))
    Me.cboValidation_Type.Value = OldValues(4)
    Me.txtSiteSelect.Value = ReadDate(CStr(OldValues(5)))
    
    Me.txtReminder.Value = OldValues(6)
    
End Sub

Private Sub cmdRedo_Click()
    'PURPOSE: Recall values replaced by undo
    
    Me.txtPrestudy_Date.Value = ReadDate(CStr(NxtOldValues(1)))
    Me.cboPrestudy_Type.Value = NxtOldValues(2)
    Me.txtValidation_Date.Value = ReadDate(CStr(NxtOldValues(3)))
    Me.cboValidation_Type.Value = NxtOldValues(4)
    Me.txtSiteSelect.Value = ReadDate(CStr(NxtOldValues(5)))
    
    Me.txtReminder.Value = NxtOldValues(6)
    
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
    Call txtPrestudy_Date_AfterUpdate
    Call txtValidation_Date_AfterUpdate
    Call txtSiteSelect_AfterUpdate
    
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
    
    Dim ReadRow(1 To 6) As Variant
    
    'Turn off Settings to speed up
    'source: https://www.automateexcel.com/vba/turn-off-screen-updating/
    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayStatusBar = False
    Application.EnableEvents = False
    
    With RegTable.ListRows(RowIndex)
        
        'Populate ReadRow Array - faster than double transpose
        ReadRow(1) = String_to_Date(Me.txtPrestudy_Date.Value)
        ReadRow(2) = Me.cboPrestudy_Type.Value
        ReadRow(3) = String_to_Date(Me.txtValidation_Date.Value)
        ReadRow(4) = Me.cboValidation_Type.Value
        ReadRow(5) = String_to_Date(Me.txtSiteSelect.Value)
        ReadRow(6) = Me.txtReminder.Value
        
        'Write to Register Table
        .Range(30) = ReadRow(1)
        .Range(31) = ReadRow(2)
        .Range(32) = ReadRow(3)
        .Range(33) = ReadRow(4)
        .Range(34) = ReadRow(5)
        .Range(35) = WriteText(ReadRow(6))
         
        'Store next old values
        NxtOldValues = ReadRow
        
        'Check if values changed
        If Not ArraysSame(ReadRow, OldValues) Then
            'Update version control
            .Range(36) = Now
            .Range(37) = Username
            
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
    ReadRow = Application.Transpose(Application.Transpose(Range(db.Cells(RowIndex, 30), db.Cells(RowIndex, 34))))
                   
    'Apply correct test on each field
    For i = LBound(ReadRow) To UBound(ReadRow)
        If ReadRow(i) <> vbNullString Then
    
            Select Case Correct(i + 22)
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
    
    'Site Selection
    'Criteria - all fields filled with dates and text (for combo box)
    cntTrue = 0
    cntEmpty = 0
    For i = 1 To 5
        If ReadRow(i) = vbNullString Then
            cntEmpty = cntEmpty + 1
        ElseIf ReadRow(i) Then
            cntTrue = cntTrue + 1
        End If
    Next i
    
    If cntEmpty = 5 Then
        db.Cells(RowIndex, 132) = vbNullString
    ElseIf cntTrue = 5 Then
        db.Cells(RowIndex, 132) = True
    Else
        db.Cells(RowIndex, 132) = False
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
    Unload form04_SiteSelect
    
    form00_Nav.show False
End Sub

Private Sub tglCDA_Click()
    'PURPOSE: Closes current form and open CDA form
    Unload form04_SiteSelect
    
    form02_CDA.show False
End Sub

Private Sub tglFS_Click()
    'PURPOSE: Closes current form and open Feasibility form
    Unload form04_SiteSelect
    
    form03_FS.show False
End Sub

Private Sub tglStudyDetail_Click()
    'PURPOSE: Closes current form and open Study Detail form
    Unload form04_SiteSelect
    
    form01_StudyDetail.show False
End Sub

Private Sub tglRecruit_Click()
    'PURPOSE: Closes current form and open Recruitment form
    Unload form04_SiteSelect
    
    form05_Recruitment.show False
End Sub

Private Sub tglEthics_Click()
    'PURPOSE: Closes current form and open Ethics form
    Unload form04_SiteSelect
    
    form06_Ethics.show False
    form06_Ethics.multiEthics.Value = 0
End Sub

Private Sub tglGov_Click()
    'PURPOSE: Closes current form and open Governance form
    Unload form04_SiteSelect
    
    form07_Governance.show False
    form07_Governance.multiGov.Value = 0
End Sub

Private Sub tglBudget_Click()
    'PURPOSE: Closes current form and open Budget form
    Unload form04_SiteSelect
    
    form08_Budget.show False
    form08_Budget.multiBudget.Value = 0
End Sub

Private Sub tglIndemnity_Click()
    'PURPOSE: Closes current form and open Indemnity form
    Unload form04_SiteSelect
    
    form09_Indemnity.show False
End Sub

Private Sub tglCTRA_Click()
    'PURPOSE: Closes current form and open CTRA form
    Unload form04_SiteSelect
    
    form10_CTRA.show False
End Sub

Private Sub tglFinDisc_Click()
    'PURPOSE: Closes current form and open Fin. Disc. form
    Unload form04_SiteSelect
    
    form11_FinDisc.show False
End Sub

Private Sub tglSIV_Click()
    'PURPOSE: Closes current form and open SIV form
    Unload form04_SiteSelect
    
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
