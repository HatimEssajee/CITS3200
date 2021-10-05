VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form08_Budget 
   Caption         =   "Budget Review"
   ClientHeight    =   5508
   ClientLeft      =   -420
   ClientTop       =   -2088
   ClientWidth     =   8748.001
   OleObjectBlob   =   "form08_Budget.frx":0000
End
Attribute VB_Name = "form08_Budget"
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
    ReDim OldValues(1 To 9)
    ReDim NxtOldValues(1 To 9)
                    
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
    
    For Each pPage In Me.multiBudget.Pages
        For Each ctrl In pPage.Controls
            Select Case True
                Case TypeOf ctrl Is MSForms.CheckBox
                    ctrl.Value = False
                Case TypeOf ctrl Is MSForms.TextBox
                    ctrl.Value = ""
                Case TypeOf ctrl Is MSForms.ComboBox
                    ctrl.Value = ""
                Case TypeOf ctrl Is MSForms.ListBox
                    ctrl.Value = ""
            End Select
                
        Next ctrl
    Next pPage
    
    'Read information from register table
    With RegTable.ListRows(RowIndex)
        Me.txtStudyName.Value = .Range(9).Value
        Me.txtVTG_Date_Finalised.Value = ReadDate(.Range(94).Value)
        Me.txtVTG_Date_Submitted.Value = ReadDate(.Range(95).Value)
        Me.txtVTG_Date_Approved.Value = ReadDate(.Range(96).Value)
        Me.txtVTG_Reminder.Value = .Range(97).Value
        
        Me.txtTKI_Date_Approved.Value = ReadDate(.Range(98).Value)
        Me.txtTKI_Reminder.Value = .Range(99).Value
        
        Me.txtPharm_Date_Quote.Value = ReadDate(.Range(100).Value)
        Me.txtPharm_Date_Finalised.Value = ReadDate(.Range(101).Value)
        Me.txtPharm_Reminder.Value = .Range(102).Value
        
    End With
    
    'Populate Old Values Array - for undo
    OldValues(1) = String_to_Date(Me.txtVTG_Date_Finalised.Value)
    OldValues(2) = String_to_Date(Me.txtVTG_Date_Submitted.Value)
    OldValues(3) = String_to_Date(Me.txtVTG_Date_Approved.Value)
    OldValues(4) = Me.txtVTG_Reminder.Value
    
    OldValues(5) = String_to_Date(Me.txtTKI_Date_Approved.Value)
    OldValues(6) = Me.txtTKI_Reminder.Value
    
    OldValues(7) = String_to_Date(Me.txtPharm_Date_Quote.Value)
    OldValues(8) = String_to_Date(Me.txtPharm_Date_Finalised.Value)
    OldValues(9) = Me.txtPharm_Reminder.Value
    
    'Initialize NxtOldValues
    NxtOldValues = OldValues
    
    'Access version control
    Call LogLastAccess
    
    'Depress and make toggle green on nav bar
    Me.tglBudget.Value = True
    Me.tglBudget.BackColor = vbGreen
    
    'Allocate tick box values
    Me.cbSaveonUnload.Value = SAG_Tick
    
    'Run date validation on data entered
    Call txtVTG_Date_Submitted_AfterUpdate
    Call txtVTG_Date_Finalised_AfterUpdate
    Call txtVTG_Date_Approved_AfterUpdate
    
    Call txtTKI_Date_Approved_AfterUpdate
    
    Call txtPharm_Date_Quote_AfterUpdate
    Call txtPharm_Date_Finalised_AfterUpdate
    
    'Reinstate Settings
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayStatusBar = True
    Application.EnableEvents = True
    
End Sub

Private Sub txtVTG_Date_Finalised_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtVTG_Date_Finalised.Value)
    
    'Display error message
    Me.errVTG_Date_Finalised.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtVTG_Date_Finalised.Value) Then
        Me.txtVTG_Date_Finalised.Value = Format(Me.txtVTG_Date_Finalised.Value, "dd-mmm-yyyy")
    End If
    
End Sub

Private Sub txtVTG_Date_Submitted_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtVTG_Date_Submitted.Value, Me.txtVTG_Date_Finalised.Value, _
            "Date entered earlier than date Finalised")

    'Display error message
    Me.errVTG_Date_Submitted.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtVTG_Date_Submitted.Value) Then
        Me.txtVTG_Date_Submitted.Value = Format(Me.txtVTG_Date_Submitted.Value, "dd-mmm-yyyy")
    End If
     
End Sub

Private Sub txtVTG_Date_Approved_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtVTG_Date_Approved.Value, Me.txtVTG_Date_Submitted.Value, _
            "Date entered earlier than date Submitted")

    'Display error message
    Me.errVTG_Date_Approved.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtVTG_Date_Approved.Value) Then
        Me.txtVTG_Date_Approved.Value = Format(Me.txtVTG_Date_Approved.Value, "dd-mmm-yyyy")
    End If
     
End Sub

Private Sub txtTKI_Date_Approved_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtTKI_Date_Approved.Value)
    
    'Display error message
    Me.errTKI_Date_Approved.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtTKI_Date_Approved.Value) Then
        Me.txtTKI_Date_Approved.Value = Format(Me.txtTKI_Date_Approved.Value, "dd-mmm-yyyy")
    End If
    
End Sub

Private Sub txtPharm_Date_Quote_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtPharm_Date_Quote.Value)
    
    'Display error message
    Me.errPharm_Date_Quote.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtPharm_Date_Quote.Value) Then
        Me.txtPharm_Date_Quote.Value = Format(Me.txtPharm_Date_Quote.Value, "dd-mmm-yyyy")
    End If
    
End Sub

Private Sub txtPharm_Date_Finalised_AfterUpdate()
    'PURPOSE: Validate date entered
    Dim err As String
    
    err = Date_Validation(Me.txtPharm_Date_Finalised.Value, Me.txtPharm_Date_Quote.Value, _
            "Date entered earlier than date" & Chr(10) & "Quote was received")

    'Display error message
    Me.errPharm_Date_Finalised.Caption = err
    
    'Change date format displayed
    If IsDate(Me.txtPharm_Date_Finalised.Value) Then
        Me.txtPharm_Date_Finalised.Value = Format(Me.txtPharm_Date_Finalised.Value, "dd-mmm-yyyy")
    End If
     
End Sub

Private Sub cmdUndo_Click()
    'PURPOSE: Recall values read from register table when the form was loaded initially
    
    Me.txtVTG_Date_Finalised.Value = ReadDate(CStr(OldValues(1)))
    Me.txtVTG_Date_Submitted.Value = ReadDate(CStr(OldValues(2)))
    Me.txtVTG_Date_Approved.Value = ReadDate(CStr(OldValues(3)))
    Me.txtVTG_Reminder.Value = OldValues(4)
    
    Me.txtTKI_Date_Approved.Value = ReadDate(CStr(OldValues(5)))
    Me.txtTKI_Reminder.Value = OldValues(6)
    
    Me.txtPharm_Date_Quote.Value = ReadDate(CStr(OldValues(7)))
    Me.txtPharm_Date_Finalised.Value = ReadDate(CStr(OldValues(8)))
    Me.txtPharm_Reminder.Value = OldValues(9)
    
End Sub

Private Sub cmdRedo_Click()
    'PURPOSE: Recall values replaced by undo
    
    Me.txtVTG_Date_Finalised.Value = ReadDate(CStr(NxtOldValues(1)))
    Me.txtVTG_Date_Submitted.Value = ReadDate(CStr(NxtOldValues(2)))
    Me.txtVTG_Date_Approved.Value = ReadDate(CStr(NxtOldValues(3)))
    Me.txtVTG_Reminder.Value = NxtOldValues(4)
    
    Me.txtTKI_Date_Approved.Value = ReadDate(CStr(NxtOldValues(5)))
    Me.txtTKI_Reminder.Value = NxtOldValues(6)
    
    Me.txtPharm_Date_Quote.Value = ReadDate(CStr(NxtOldValues(7)))
    Me.txtPharm_Date_Finalised.Value = ReadDate(CStr(NxtOldValues(8)))
    Me.txtPharm_Reminder.Value = NxtOldValues(9)
    
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
    Call txtVTG_Date_Submitted_AfterUpdate
    Call txtVTG_Date_Finalised_AfterUpdate
    Call txtVTG_Date_Approved_AfterUpdate
    
    Call txtTKI_Date_Approved_AfterUpdate
    
    Call txtPharm_Date_Quote_AfterUpdate
    Call txtPharm_Date_Finalised_AfterUpdate
    
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
    Dim ReadRow(1 To 9) As Variant
    
    'Turn off Settings to speed up
    'source: https://www.automateexcel.com/vba/turn-off-screen-updating/
    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayStatusBar = False
    Application.EnableEvents = False
    
    With RegTable.ListRows(RowIndex)
        
        'Populate ReadRow Array - faster than double transpose
        ReadRow(1) = String_to_Date(Me.txtVTG_Date_Finalised.Value)
        ReadRow(2) = String_to_Date(Me.txtVTG_Date_Submitted.Value)
        ReadRow(3) = String_to_Date(Me.txtVTG_Date_Approved.Value)
        ReadRow(4) = Me.txtVTG_Reminder.Value
        
        ReadRow(5) = String_to_Date(Me.txtTKI_Date_Approved.Value)
        ReadRow(6) = Me.txtTKI_Reminder.Value
        
        ReadRow(7) = String_to_Date(Me.txtPharm_Date_Quote.Value)
        ReadRow(8) = String_to_Date(Me.txtPharm_Date_Finalised.Value)
        ReadRow(9) = Me.txtPharm_Reminder.Value
        
         'Write to Register table
        .Range(94) = ReadRow(1)
        .Range(95) = ReadRow(2)
        .Range(96) = ReadRow(3)
        .Range(97) = ReadRow(4)
        
        .Range(98) = ReadRow(5)
        .Range(99) = ReadRow(6)
        
        .Range(100) = ReadRow(7)
        .Range(101) = ReadRow(8)
        .Range(102) = ReadRow(9)
        
        'Store next old values
        NxtOldValues = ReadRow
        
        'Check if values changed
        If Not ArraysSame(ReadRow, OldValues) Then
                
            'Update version control
            .Range(103) = Now
            .Range(104) = Username
            
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
    ReadRow = Application.Transpose(Application.Transpose(Range(db.Cells(RowIndex, 94), db.Cells(RowIndex, 101))))
                   
    'Apply correct test on each field
    For i = LBound(ReadRow) To UBound(ReadRow)
        If ReadRow(i) <> vbNullString Then
    
            Select Case Correct(i + 86)
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
    
    'VTG Budget
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
        db.Cells(RowIndex, 146) = vbNullString
    ElseIf cntTrue = 3 Then
        db.Cells(RowIndex, 146) = True
    Else
        db.Cells(RowIndex, 146) = False
    End If
    
    'TKI Budget
    'Criteria - has to be date
    db.Cells(RowIndex, 147) = ReadRow(5)
    
    
    'Pharmacy Budget
    'Criteria - all fields filled with dates
    cntTrue = 0
    cntEmpty = 0
    For i = 7 To 8
        If ReadRow(i) = vbNullString Then
            cntEmpty = cntEmpty + 1
        ElseIf ReadRow(i) Then
            cntTrue = cntTrue + 1
        End If
    Next i
    
    If cntEmpty = 2 Then
        db.Cells(RowIndex, 148) = vbNullString
    ElseIf cntTrue = 2 Then
        db.Cells(RowIndex, 148) = True
    Else
        db.Cells(RowIndex, 148) = False
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
    Unload form08_Budget
    
    form00_Nav.show False
End Sub

Private Sub tglCDA_Click()
    'PURPOSE: Closes current form and open CDA form
    Unload form08_Budget
    
    form02_CDA.show False
End Sub

Private Sub tglFS_Click()
    'PURPOSE: Closes current form and open Feasibility form
    Unload form08_Budget
    
    form03_FS.show False
End Sub

Private Sub tglSiteSelect_Click()
    'PURPOSE: Closes current form and open Site Select form
    Unload form08_Budget
    
    form04_SiteSelect.show False
End Sub

Private Sub tglRecruit_Click()
    'PURPOSE: Closes current form and open Recruitment form
    Unload form08_Budget
    
    form05_Recruitment.show False
End Sub

Private Sub tglEthics_Click()
    'PURPOSE: Closes current form and open Ethics form
    Unload form08_Budget
    
    form06_Ethics.show False
    form06_Ethics.multiEthics.Value = 0
End Sub

Private Sub tglGov_Click()
    'PURPOSE: Closes current form and open Governance form
    Unload form08_Budget
    
    form07_Governance.show False
    form07_Governance.multiGov.Value = 0
End Sub

Private Sub tglStudyDetail_Click()
    'PURPOSE: Closes current form and open Study Detail form
    Unload form08_Budget
    
    form01_StudyDetail.show False
End Sub

Private Sub tglIndemnity_Click()
    'PURPOSE: Closes current form and open Indemnity form
    Unload form08_Budget
    
    form09_Indemnity.show False
End Sub

Private Sub tglCTRA_Click()
    'PURPOSE: Closes current form and open CTRA form
    Unload form08_Budget
    
    form10_CTRA.show False
End Sub

Private Sub tglFinDisc_Click()
    'PURPOSE: Closes current form and open Fin. Disc. form
    Unload form08_Budget
    
    form11_FinDisc.show False
End Sub

Private Sub tglSIV_Click()
    'PURPOSE: Closes current form and open SIV form
    Unload form08_Budget
    
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
