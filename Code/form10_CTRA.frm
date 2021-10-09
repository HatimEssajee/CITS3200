VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form10_CTRA 
   Caption         =   "CTRA"
   ClientHeight    =   6852
   ClientLeft      =   -468
   ClientTop       =   -2376
   ClientWidth     =   11460
   OleObjectBlob   =   "form10_CTRA.frx":0000
End
Attribute VB_Name = "form10_CTRA"
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
    
    'Turn off Settings to speed up
    'source: https://www.automateexcel.com/vba/turn-off-screen-updating/
    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayStatusBar = False
    Application.EnableEvents = False
    
    'Erase and initialise arrays
    ReDim OldValues(1 To 8)
    ReDim NxtOldValues(1 To 8)
                    
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
    
    'Read information from register table
    With RegTable.ListRows(RowIndex)
        Me.txtStudyName.Value = .Range(9).Value
        Me.txtDate_RGC.Value = ReadDate(.Range(111).Value)
        Me.txtDate_UWA.Value = ReadDate(.Range(112).Value)
        Me.txtDate_Finance.Value = ReadDate(.Range(113).Value)
        Me.txtDate_COO.Value = ReadDate(.Range(114).Value)
        Me.txtDate_VTG.Value = ReadDate(.Range(115).Value)
        Me.txtDate_Company.Value = ReadDate(.Range(116).Value)
        Me.txtDate_Finalised.Value = ReadDate(.Range(117).Value)
        Me.txtReminder.Value = .Range(118).Value
    End With
    
    'Populate Old Values Array - for undo
    OldValues(1) = String_to_Date(Me.txtDate_RGC.Value)
    OldValues(2) = String_to_Date(Me.txtDate_UWA.Value)
    OldValues(3) = String_to_Date(Me.txtDate_Finance.Value)
    OldValues(4) = String_to_Date(Me.txtDate_COO.Value)
    OldValues(5) = String_to_Date(Me.txtDate_VTG.Value)
    OldValues(6) = String_to_Date(Me.txtDate_Company.Value)
    OldValues(7) = String_to_Date(Me.txtDate_Finalised.Value)
    OldValues(8) = Me.txtReminder.Value
    
    'Initialize NxtOldValues
    NxtOldValues = OldValues
    
    'Access version control
    Call LogLastAccess
    
    'Depress and make toggle green on nav bar
    Me.tglCTRA.Value = True
    Me.tglCTRA.BackColor = vbGreen
    
    'Allocate tick box values
    Me.cbSaveonUnload.Value = SAG_Tick
    
    'Run date validation on data entered
    Call txtDate_RGC_AfterUpdate
    Call txtDate_UWA_AfterUpdate
    Call txtDate_Finance_AfterUpdate
    Call txtDate_COO_AfterUpdate
    Call txtDate_VTG_AfterUpdate
    Call txtDate_Company_AfterUpdate
    Call txtDate_Finalised_AfterUpdate
    
    'Reinstate Settings
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayStatusBar = True
    Application.EnableEvents = True
    
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

Private Sub cmdUndo_Click()
    'PURPOSE: Recall values read from register table when the form was loaded initially
    
    Me.txtDate_RGC.Value = ReadDate(CStr(OldValues(1)))
    Me.txtDate_UWA.Value = ReadDate(CStr(OldValues(2)))
    Me.txtDate_Finance.Value = ReadDate(CStr(OldValues(3)))
    Me.txtDate_COO.Value = ReadDate(CStr(OldValues(4)))
    Me.txtDate_VTG.Value = ReadDate(CStr(OldValues(5)))
    Me.txtDate_Company.Value = ReadDate(CStr(OldValues(6)))
    Me.txtDate_Finalised.Value = ReadDate(CStr(OldValues(7)))
    Me.txtReminder.Value = OldValues(8)
    
End Sub

Private Sub cmdRedo_Click()
    'PURPOSE: Recall values replaced by undo
    
    Me.txtDate_RGC.Value = ReadDate(CStr(NxtOldValues(1)))
    Me.txtDate_UWA.Value = ReadDate(CStr(NxtOldValues(2)))
    Me.txtDate_Finance.Value = ReadDate(CStr(NxtOldValues(3)))
    Me.txtDate_COO.Value = ReadDate(CStr(NxtOldValues(4)))
    Me.txtDate_VTG.Value = ReadDate(CStr(NxtOldValues(5)))
    Me.txtDate_Company.Value = ReadDate(CStr(NxtOldValues(6)))
    Me.txtDate_Finalised.Value = ReadDate(CStr(NxtOldValues(7)))
    Me.txtReminder.Value = NxtOldValues(8)
    
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
    Call txtDate_RGC_AfterUpdate
    Call txtDate_UWA_AfterUpdate
    Call txtDate_Finance_AfterUpdate
    Call txtDate_COO_AfterUpdate
    Call txtDate_VTG_AfterUpdate
    Call txtDate_Company_AfterUpdate
    Call txtDate_Finalised_AfterUpdate
    
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
    Dim ReadRow(1 To 8) As Variant
    
    'Turn off Settings to speed up
    'source: https://www.automateexcel.com/vba/turn-off-screen-updating/
    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayStatusBar = False
    Application.EnableEvents = False
    
    With RegTable.ListRows(RowIndex)
               
        'Populate ReadRow Array - faster than double transpose
        ReadRow(1) = String_to_Date(Me.txtDate_RGC.Value)
        ReadRow(2) = String_to_Date(Me.txtDate_UWA.Value)
        ReadRow(3) = String_to_Date(Me.txtDate_Finance.Value)
        ReadRow(4) = String_to_Date(Me.txtDate_COO.Value)
        ReadRow(5) = String_to_Date(Me.txtDate_VTG.Value)
        ReadRow(6) = String_to_Date(Me.txtDate_Company.Value)
        ReadRow(7) = String_to_Date(Me.txtDate_Finalised.Value)
        ReadRow(8) = Me.txtReminder.Value
        
        'Write to Register table
        .Range(111) = ReadRow(1)
        .Range(112) = ReadRow(2)
        .Range(113) = ReadRow(3)
        .Range(114) = ReadRow(4)
        .Range(115) = ReadRow(5)
        .Range(116) = ReadRow(6)
        .Range(117) = ReadRow(7)
        .Range(118) = WriteText(ReadRow(8))
        
        'Store next old values
        NxtOldValues = ReadRow
        
        'Check if values changed
        If Not ArraysSame(ReadRow, OldValues) Then
        
            'Update version control
            .Range(119) = Now
            .Range(120) = Username
            
            'Apply completion status
            Call Fill_Completion_Status
            DoEvents
        End If
        
        'Clear array
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
    ReadRow = Application.Transpose(Application.Transpose(Range(db.Cells(RowIndex, 111), db.Cells(RowIndex, 117))))
                   
    'Apply correct test on each field
    For i = LBound(ReadRow) To UBound(ReadRow)
        If ReadRow(i) <> vbNullString Then
    
            Select Case Correct(i + 103)
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
    
    'CTRA
    'Criteria - all fields filled with dates
    cntTrue = 0
    cntEmpty = 0
    For i = 1 To 7
        If ReadRow(i) = vbNullString Then
            cntEmpty = cntEmpty + 1
        ElseIf ReadRow(i) Then
            cntTrue = cntTrue + 1
        End If
    Next i
    
    If cntEmpty = 7 And db.Cells(RowIndex, 118).Value <> vbNullString Then
        db.Cells(RowIndex, 150) = False
    ElseIf cntEmpty = 7 Then
        db.Cells(RowIndex, 150) = vbNullString
    ElseIf cntTrue = 7 Then
        db.Cells(RowIndex, 150) = True
    Else
        db.Cells(RowIndex, 150) = False
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
    Unload form10_CTRA
    
    form00_Nav.show False
End Sub

Private Sub tglCDA_Click()
    'PURPOSE: Closes current form and open CDA form
    Unload form10_CTRA
    
    form02_CDA.show False
End Sub

Private Sub tglFS_Click()
    'PURPOSE: Closes current form and open Feasibility form
    Unload form10_CTRA
    
    form03_FS.show False
End Sub

Private Sub tglSiteSelect_Click()
    'PURPOSE: Closes current form and open Site Select form
    Unload form10_CTRA
    
    form04_SiteSelect.show False
End Sub

Private Sub tglRecruit_Click()
    'PURPOSE: Closes current form and open Recruitment form
    Unload form10_CTRA
    
    form05_Recruitment.show False
End Sub

Private Sub tglEthics_Click()
    'PURPOSE: Closes current form and open Ethics form
    Unload form10_CTRA
    
    form06_Ethics.show False
    form06_Ethics.multiEthics.Value = 0
End Sub

Private Sub tglGov_Click()
    'PURPOSE: Closes current form and open Governance form
    Unload form10_CTRA
    
    form07_Governance.show False
    form07_Governance.multiGov.Value = 0
End Sub

Private Sub tglBudget_Click()
    'PURPOSE: Closes current form and open Budget form
    Unload form10_CTRA
    
    form08_Budget.show False
    form08_Budget.multiBudget.Value = 0
End Sub

Private Sub tglIndemnity_Click()
    'PURPOSE: Closes current form and open Indemnity form
    Unload form10_CTRA
    
    form09_Indemnity.show False
End Sub

Private Sub tglStudyDetail_Click()
    'PURPOSE: Closes current form and open Study Detail form
    Unload form10_CTRA
    
    form01_StudyDetail.show False
End Sub

Private Sub tglFinDisc_Click()
    'PURPOSE: Closes current form and open Fin. Disc. form
    Unload form10_CTRA
    
    form11_FinDisc.show False
End Sub

Private Sub tglSIV_Click()
    'PURPOSE: Closes current form and open SIV form
    Unload form10_CTRA
    
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
