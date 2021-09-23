VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form00_Nav 
   Caption         =   "Vaccine Trial Study Start-up Tracker"
   ClientHeight    =   9096.001
   ClientLeft      =   -30
   ClientTop       =   -360
   ClientWidth     =   10920
   OleObjectBlob   =   "form00_Nav.frx":0000
End
Attribute VB_Name = "form00_Nav"
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
    'PURPOSE: Clear form on initialization and fill combo box with data from array
    'Source: https://www.contextures.com/xlUserForm02.html
    'Source: https://www.contextures.com/Excel-VBA-ComboBox-Lists.html
    Dim cboList_StudyStatus As Variant, item As Variant
    Dim ctrl As MSForms.Control
    
    'Load default values
    cboList_StudyStatus = Array("Pre-commencement", "Commenced", "Not Going Ahead")
    
    If Not RegTable.DataBodyRange Is Nothing Then
        StudyStatus = RegTable.DataBodyRange.Columns(8)
    End If
    
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
        Me.cboStudyStatus.AddItem item
    Next item
    
    'Allocate tick box values
    Me.cbOnlyCurrent.Value = Tick
    Me.cbFastCycle.Value = FC_Tick
    
    'Format fields
    If RowIndex > 0 Then
        Call Read_Table
        Me.cboStudyStatus.ForeColor = StudyStatus_Colour(Me.cboStudyStatus.Value)
    End If
    
    'Unload search display
    EraseIfArray (DisplayArr)
    
End Sub


Private Sub cmdClose_Click()
    'PURPOSE: Closes current form
    
    'Access version control
    Call LogLastAccess
        
    Unload Me
    
    'Empty Array as no longer needed
    EraseIfArray (StudyStatus)
    EraseIfArray (DisplayArr)
    
End Sub

Private Sub cmdClear_Click()
    
    'Reset Default values
    RowIndex = -1
    Tick = True
    
    'PURPOSE: Reinitialise form
    Call UserForm_Initialize
        
End Sub

Private Sub cmdNew_Click()
    'PURPOSE: Closes current form and open Study Detail form
    
    Dim FoundCell As Range
    Dim StudyName As String
    Dim ReadRow As Variant
    
    'Set Public Variable
    StudyName = Me.txtStudyName.Value
    
    'Check if study name is entered
    If StudyName = vbNullString Then
        Me.errSearch.Caption = "Please enter a study name to create a new record"
        Exit Sub
    End If
    
    'Check if study name already in Register table
    'Source: https://www.thespreadsheetguru.com/blog/2014/6/20/the-vba-guide-to-listobject-excel-tables
    On Error Resume Next
    Set FoundCell = RegTable.DataBodyRange.Columns(10).find(StudyName, LookAt:=xlWhole)
    On Error GoTo 0
    
    If Not FoundCell Is Nothing Then
        RowIndex = RegTable.ListRows(FoundCell.Row - RegTable.HeaderRowRange.Row).Index
        Me.errSearch.Caption = "Study already exists, consider edit instead"
        Exit Sub
    End If
    
    'Add Row to register table and repoint row references
    'Source: https://www.bluepecantraining.com/portfolio/excel-vba-how-to-add-rows-and-columns-to-excel-table-with-vba-macro/
    Set ReadRow = RegTable.ListRows.Add
    
    RowIndex = RegTable.ListRows.Count
    
    With ReadRow
    
        .Range(1) = RowIndex
        
        'Creation version control
        .Range(2) = Now
        .Range(3) = Username
        
        .Range(8) = "Pre-commencement"
        .Range(9) = Me.txtProtocolNum.Value
        .Range(10) = StudyName
        .Range(11) = Me.txtSponsor.Value
        .Range(140).Formula = "=IFERROR(MATCH(FALSE,Register[@[Study Details Complete]:[SIV Complete]],0),1)"
        
        'Update version control
        .Range(15) = .Range(2).Value
        .Range(16) = .Range(3).Value
    End With
        
    Unload form00_Nav
    
    form01_StudyDetail.Show False
    
    'Empty Array as no longer needed
    EraseIfArray (StudyStatus)
    EraseIfArray (DisplayArr)
    
End Sub

Private Sub cbOnlyCurrent_Click()
    'PURPOSE: Change value of Tick variable
    Tick = Me.cbOnlyCurrent.Value
End Sub

Private Sub cbFastCycle_Click()
    'PURPOSE: Change value of FC_Tick variable
    FC_Tick = Me.cbFastCycle.Value
End Sub

Private Sub cboStudyStatus_AfterUpdate()
    'PURPOSE: Change text color of combo box status based on value
    
    Dim SIVDate As String

    'Unique change events
    SIVDate = RegTable.DataBodyRange.Cells(RowIndex, 112).Value
    
    'Undeleting entry
    If OldStudyStatus = "DELETED" And Me.cboStudyStatus <> "DELETED" Then
        
        'Clear Deletion Log
        With RegTable.ListRows(RowIndex)
            'Deletion version control
            .Range(4) = vbNullString
            .Range(5) = vbNullString
            
            'Update version control
            .Range(15) = Now
            .Range(16) = Username
        End With
        
    End If
    
    'Swap to commenced if SIV before today
    If Me.cboStudyStatus.Value = "Pre-commencement" And SIVDate <> vbNullString And _
        String_to_Date(SIVDate) < Now Then
        
        Me.cboStudyStatus.Value = "Commenced"
        
        'Clear Deletion Log
        With RegTable.ListRows(RowIndex)
            'Update version control
            .Range(15) = Now
            .Range(16) = Username
            
        End With
        
    End If
    
    'Update value in table
    RegTable.DataBodyRange.Cells(RowIndex, 8).Value = Me.cboStudyStatus.Value
    Me.cboStudyStatus.ForeColor = StudyStatus_Colour(Me.cboStudyStatus.Value)
    StudyStatus = RegTable.DataBodyRange.Columns(8)
    
    'Update Access log
    Call LogLastAccess
    
End Sub

Private Sub cmdDelete_Click()
    'PURPOSE: Non-permanent delete of entry
    
    Dim confirm As Integer
    
    'Confirm deletion
    confirm = MsgBox("Are you sure you want to delete Project data?", vbYesNo, "WARNING!")

    'If select no then cancel deletion
    If confirm = vbNo Then
        Exit Sub
    End If

    'Change entry if RowIndex was found via search or new entry
    If RowIndex > 0 Then
        
        'Update deletion log
        With RegTable.ListRows(RowIndex)
            
            'Deletion version control
            .Range(4) = Now
            .Range(5) = Username
            .Range(8) = "DELETED"
            
            'Update version control
            .Range(15) = .Range(4).Value
            .Range(16) = .Range(5).Value
        End With
    
    
        'Change status
        With Me.cboStudyStatus
            .Value = "DELETED"
            .ForeColor = vbRed
        End With
        
        OldStudyStatus = "DELETED"
        
    End If
    
End Sub

Private Sub cmdChangeLog_Click()
    'PURPOSE: Open change log form
    
    If RowIndex < 0 Then
        errSearch.Caption = "Need study entry identified to view log"
    Else
        form13_ChangeLog.Show False
    End If
    
    'Store form position
    UserFormTopPos = Me.Top
    UserFormLeftPos = Me.Left
    
    'Start Position of Log
    UserFormTopPosC = Me.Top
    UserFormLeftPosC = Me.Left
End Sub

Private Sub cmdReminders_Click()
    'PURPOSE: Open reminder log form
    
    If RowIndex < 0 Then
        errSearch.Caption = "Need study entry identified to view log"
    Else
        form14_ReminderLog.Show False
    End If
    
    'Store form position
    UserFormTopPos = Me.Top
    UserFormLeftPos = Me.Left
    
    'Start Position of Log
    UserFormTopPosR = Me.Top
    UserFormLeftPosR = Me.Left
    
End Sub

Private Sub cmdEdit_Click()
    'PURPOSE: Closes current form and open Study Detail form
    
    'Redirect to new entry creation if no data
    If RegTable.DataBodyRange Is Nothing Then
        Call cmdNew_Click
        Exit Sub
    End If
    
    If RowIndex < 0 Then
        errSearch.Caption = "Could not locate entry in register, consider creating new record"
    Else
        
        'Write changes to register table
        With RegTable.ListRows(RowIndex)
            .Range(9) = Me.txtProtocolNum.Value
            .Range(10) = Me.txtStudyName.Value
            .Range(11) = Me.txtSponsor.Value
            
            'Update version control
            .Range(15) = Now
            .Range(16) = Username
        End With
        
        'Empty Array as no longer needed
        EraseIfArray (StudyStatus)
        EraseIfArray (DisplayArr)
        
        Call Fill_Completion_Status
        DoEvents
        
        Call Apply_FastCycle
        DoEvents
        
        Unload form00_Nav
    End If
    
End Sub

Private Sub Fill_Completion_Status()
    
    'PURPOSE: Evaluate entry completion status
    
    'Initialise by making all values to be false
    Range(RegTable.DataBodyRange.Cells(RowIndex, 116), RegTable.DataBodyRange.Cells(RowIndex, 139)).Value = False
    
    'Assess values
    With RegTable.ListRows(RowIndex)
            
        'Study details
        If Application.CountA(Range(RegTable.DataBodyRange.Cells(RowIndex, 9), _
            RegTable.DataBodyRange.Cells(RowIndex, 13))) = 5 Then
            .Range(116).Value = True
        End If
        
        'CDA
        If .Range(21).Value = vbNullString Then
            .Range(117).Value = vbNullString
        Else
            .Range(117).Value = IsDate(.Range(21).Value)
        End If
        
        'FS
        If .Range(23).Value = vbNullString Then
            .Range(118).Value = vbNullString
        Else
            .Range(118).Value = IsDate(.Range(23).Value)
        End If
        
        'Site selection
        If .Range(32).Value = vbNullString Then
            .Range(119).Value = vbNullString
        Else
            .Range(119).Value = IsDate(.Range(32).Value)
        End If
        
        'Recruitment
        If .Range(37).Value = vbNullString Then
            .Range(120).Value = vbNullString
        ElseIf .Range(37).Value = "Complete" Then
            .Range(120).Value = True
        End If
        
        'Ethics
        If Application.CountA(Range(RegTable.DataBodyRange.Cells(RowIndex, 41), _
            RegTable.DataBodyRange.Cells(RowIndex, 54))) = 0 Then
            .Range(121).Value = False
        ElseIf .Range(41).Value = vbNullString Then
            .Range(121).Value = vbNullString
        ElseIf IsDate(.Range(44).Value) Then
            .Range(121).Value = True
        End If
        
        If .Range(45).Value = vbNullString And .Range(46).Value = vbNullString Then
            .Range(122).Value = vbNullString
        ElseIf IsDate(.Range(47).Value) Then
            .Range(122).Value = True
        End If
        
        If .Range(48).Value = vbNullString Then
            .Range(123).Value = vbNullString
        ElseIf IsDate(.Range(49).Value) Then
            .Range(123).Value = True
        End If
        
        If .Range(50).Value = vbNullString Then
            .Range(124).Value = vbNullString
        ElseIf IsDate(.Range(51).Value) Then
            .Range(124).Value = True
        End If
        
        If .Range(52).Value = vbNullString And .Range(53).Value = vbNullString Then
            .Range(125).Value = vbNullString
        ElseIf IsDate(.Range(54).Value) Then
            .Range(125).Value = True
        End If
        
        'Governance
        If Application.CountA(Range(RegTable.DataBodyRange.Cells(RowIndex, 58), _
            RegTable.DataBodyRange.Cells(RowIndex, 79))) = 0 Then
            .Range(126).Value = False
        ElseIf .Range(58).Value = vbNullString Then
            .Range(126).Value = vbNullString
        ElseIf IsDate(.Range(60).Value) Then
            .Range(126).Value = True
        End If
        
        If .Range(61).Value = vbNullString Then
            .Range(127).Value = vbNullString
        ElseIf IsDate(.Range(63).Value) Then
            .Range(127).Value = True
        End If
        
        If .Range(64).Value = vbNullString Then
            .Range(128).Value = vbNullString
        ElseIf IsDate(.Range(66).Value) Then
            .Range(128).Value = True
        End If
        
        If .Range(67).Value = vbNullString Then
            .Range(129).Value = vbNullString
        ElseIf IsDate(.Range(69).Value) Then
            .Range(129).Value = True
        End If
    
        If .Range(70).Value = vbNullString Then
            .Range(130).Value = vbNullString
        ElseIf IsDate(.Range(72).Value) Then
            .Range(130).Value = True
        End If
        
        If .Range(73).Value = vbNullString Then
            .Range(131).Value = vbNullString
        ElseIf IsDate(.Range(75).Value) Then
            .Range(131).Value = True
        End If
        
        If .Range(76).Value = vbNullString And .Range(77).Value = vbNullString Then
            .Range(132).Value = vbNullString
        ElseIf IsDate(.Range(79).Value) Then
            .Range(132).Value = True
        End If
        
        'Budget
        If Application.CountA(Range(RegTable.DataBodyRange.Cells(RowIndex, 83), _
            RegTable.DataBodyRange.Cells(RowIndex, 88))) = 0 Then
            .Range(133).Value = False
        ElseIf IsDate(.Range(83).Value) And IsDate(.Range(85).Value) Then
            .Range(133).Value = True
        End If
        
        .Range(134).Value = IsDate(.Range(86).Value)
        
        If IsDate(.Range(87).Value) And IsDate(.Range(88).Value) Then
            .Range(135).Value = True
        End If
        
        'Indemnity
        If IsDate(.Range(92).Value) And IsDate(.Range(94).Value) Then
            .Range(136).Value = True
        End If
        
        'CTRA
        If IsDate(.Range(98).Value) And IsDate(.Range(104).Value) Then
            .Range(137).Value = True
        End If
        
        'Financial Disclosure
        .Range(138).Value = IsDate(.Range(108).Value)
        
        'SIV
        .Range(139).Value = IsDate(.Range(112).Value)
        
        If .Range(140).Value = vbNullString Then
            .Range(140).Formula = "=IFERROR(MATCH(FALSE,Register[@[Study Details Complete]:[SIV Complete]],0),1)"
        End If
        
    End With
    
End Sub
Private Sub Apply_FastCycle()
    'PURPOSE: Load next userform based on fast cycle value
    
    Dim loc As Long
    
    loc = RegTable.DataBodyRange.Cells(RowIndex, 140).Value
    
    'Unload form00_Nav
    
    'Apply fast cycle
        If FC_Tick Then
            Select Case loc
                Case 1
                    form01_StudyDetail.Show False
                Case 2
                    form02_CDA.Show False
                Case 3
                    form03_FS.Show False
                Case 4
                    form04_SiteSelect.Show False
                Case 5
                    form05_Recruitment.Show False
                Case 6
                    form06_Ethics.Show False
                    form06_Ethics.multiEthics.Value = 0
                Case 7
                    form06_Ethics.Show False
                    form06_Ethics.multiEthics.Value = 1
                Case 8
                    form06_Ethics.Show False
                    form06_Ethics.multiEthics.Value = 2
                Case 9
                    form06_Ethics.Show False
                    form06_Ethics.multiEthics.Value = 3
                Case 10
                    form06_Ethics.Show False
                    form06_Ethics.multiEthics.Value = 4
                Case 11
                    form07_Governance.Show False
                    form07_Governance.multiGov.Value = 0
                Case 12
                    form07_Governance.Show False
                    form07_Governance.multiGov.Value = 1
                Case 13
                    form07_Governance.Show False
                    form07_Governance.multiGov.Value = 2
                Case 14
                    form07_Governance.Show False
                    form07_Governance.multiGov.Value = 3
                Case 15
                    form07_Governance.Show False
                    form07_Governance.multiGov.Value = 4
                Case 16
                    form07_Governance.Show False
                    form07_Governance.multiGov.Value = 5
                Case 17
                    form07_Governance.Show False
                    form07_Governance.multiGov.Value = 6
                Case 18
                    form08_Budget.Show False
                    form08_Budget.multiBudget.Value = 0
                Case 19
                    form08_Budget.Show False
                    form08_Budget.multiBudget.Value = 1
                Case 20
                    form08_Budget.Show False
                    form08_Budget.multiBudget.Value = 2
                Case 21
                    form09_Indemnity.Show False
                Case 22
                    form10_CTRA.Show False
                Case 23
                    form11_FinDisc.Show False
                Case 24
                    form12_SIV.Show False
                Case Else
                    form01_StudyDetail.Show False
                End Select
                
        Else
            form01_StudyDetail.Show False
        End If
        
End Sub

Private Sub cmdSearch_Click()
    'PURPOSE: Populate list box with keyword search results
    'SOURCE: https://stackoverflow.com/questions/45356240/vba-for-selecting-a-number-of-columns-in-an-excel-table
    
    Dim Sponsor As String
    Dim ProtocolNum As String
    Dim SearchArr As Variant, TempArr() As Variant
    Dim SearchStatus As String
    Dim i As Integer, j As Integer
    Dim StudyName As String
    
    'Clear search array,list box and error message in memory
    EraseIfArray (DisplayArr)
    Me.lstSearch.Clear
    errSearch.Caption = vbNullString
    
    SearchArr = RegTable.ListColumns(8).DataBodyRange.Resize(, 4)
    If IsArrayEmpty(SearchArr) Then
        errSearch.Caption = "Study register is empty"
        Exit Sub
    End If
    
    j = 1
    
    'Store values in temporary variables
    Sponsor = Me.txtSponsor.Value
    ProtocolNum = Me.txtProtocolNum.Value
    StudyName = Me.txtStudyName.Value
    
    
    For i = 1 To UBound(SearchArr)
        If (Not (Tick) Or (Tick And SearchArr(i, 1) = "Pre-commencement")) And _
            (StudyName = vbNullString Or (Len(StudyName) > 0 And InStr(1, SearchArr(i, 3), StudyName, vbTextCompare) > 0)) And _
            (ProtocolNum = vbNullString Or (Len(ProtocolNum) > 0 And InStr(1, SearchArr(i, 2), ProtocolNum, vbTextCompare) > 0)) And _
            (Sponsor = vbNullString Or (Len(Sponsor) > 0 And InStr(1, SearchArr(i, 4), Sponsor, vbTextCompare) > 0)) Then
            
            'Grow display array
            ReDim Preserve TempArr(1 To 5, 1 To j)
            
            TempArr(1, j) = SearchArr(i, 4)
            TempArr(2, j) = SearchArr(i, 2)
            TempArr(3, j) = SearchArr(i, 3)
            TempArr(4, j) = SearchArr(i, 1)
            TempArr(5, j) = i
            
            j = j + 1
            
        End If
    Next i
    
    If IsArrayEmpty(TempArr) Then
        errSearch.Caption = "No records found matching query"
        Exit Sub
    End If
    
    'Transpose display array
    j = TransposeArray(TempArr, DisplayArr)
    
    EraseIfArray (SearchArr)
    EraseIfArray (TempArr)
    
    'Fill list box but retain shape and location
    'Source: https://www.mrexcel.com/board/threads/unexpected-changes-to-listbox-height.604737/
    With Me.lstSearch
        .Top = 205.8
        .Left = 12
        .Height = 88.45
        .Width = 406.55
        .IntegralHeight = False 'needed to stop list box changing position
        .List = DisplayArr
    End With
         
End Sub

Private Sub lstSearch_Click()
    'PURPOSE: Trigger populating input fields based on list box selection
    Dim i As Long, ListCount As Long
    
    'Determine no. of items in list box
    ListCount = Me.lstSearch.ListCount
    

    'Loop through items in list box until selected item found
    For i = 0 To ListCount - 1
        If Me.lstSearch.Selected(i) = True Then
            
            'Get RowIndex from hidden column
            RowIndex = DisplayArr(i + 1, 5)
            Exit For
        End If
    Next
    
    Call Read_Table
    
End Sub

Private Sub cmdJumpForw_Click()
    'PURPOSE: Redirect to newest
    
    Dim temp As Variant
    Dim r As Long
    
    r = RowIndex
    temp = StudyStatus
    

    'Check if got StudyStatus is a valid array and in the case of checkbox if it contains Pre-commencement
    If RegTable.DataBodyRange Is Nothing Or (Tick And Not Contains(StudyStatus, "Pre-commencement")) Then
        Call cmdClear_Click
        errSearch.Caption = "No data found in register"
        Exit Sub
    End If
    
    If IsArray(StudyStatus) Then
        RowIndex = UBound(StudyStatus)
    Else
        RowIndex = 1
        GoTo CallForm
    End If
    
    'Conditional stepping
    If Tick And IsArray(StudyStatus) Then
        'Loop through study status array
        Do While InStr(1, "Pre-commencement", StudyStatus(RowIndex, 1), vbTextCompare) = 0 And RowIndex > 1
            RowIndex = RowIndex - 1
        Loop
    End If

CallForm:
    'Clear form before bringing in new data
    Call UserForm_Initialize
    DoEvents
    
    
End Sub

Private Sub cmdNext_Click()
    'PURPOSE: Determine next entry row in register table depending on check box value
    
    Dim BtmRow As Long
    
    'Check if got StudyStatus is a valid array and in the case of checkbox if it contains Pre-commencement
    If RegTable.DataBodyRange Is Nothing Or (Tick And Not Contains(StudyStatus, "Pre-commencement")) Then
        Call cmdClear_Click
        errSearch.Caption = "No data found in register"
        Exit Sub
    End If
    
    'Repoint to RowIndex
    If IsArray(StudyStatus) Then
        BtmRow = UBound(StudyStatus)
    Else
        BtmRow = 1
    End If
    
    If RowIndex < 0 Or RowIndex = BtmRow Then
        RowIndex = 1
    Else
        RowIndex = RowIndex + 1
    End If
    
    'Conditional stepping
    If Tick And IsArray(StudyStatus) Then
        'Loop through study status array
        Do While InStr(1, "Pre-commencement", StudyStatus(RowIndex, 1), vbTextCompare) = 0
            RowIndex = RowIndex + 1
            If RowIndex > BtmRow Then
                RowIndex = 1
            End If
        Loop
    End If
        
    'Clear form before bringing in new data
    Call UserForm_Initialize
    DoEvents
    
End Sub

    
Private Sub cmdJumpBack_Click()
    'PURPOSE: Redirect to newest
    
    Dim BtmRow As Long
    
    'Check if got StudyStatus is a valid array and in the case of checkbox if it contains Pre-commencement
    If RegTable.DataBodyRange Is Nothing Or (Tick And Not Contains(StudyStatus, "Pre-commencement")) Then
        Call cmdClear_Click
        errSearch.Caption = "No data found in register"
        Exit Sub
    End If
    
    If IsArray(StudyStatus) Then
        RowIndex = LBound(StudyStatus)
        BtmRow = UBound(StudyStatus)
    Else
        RowIndex = 1
        BtmRow = 1
    End If
    
    'Conditional stepping
    If Tick And IsArray(StudyStatus) Then
        'Loop through study status array
        Do While InStr(1, "Pre-commencement", StudyStatus(RowIndex, 1), vbTextCompare) = 0 And RowIndex < BtmRow
            RowIndex = RowIndex + 1
        Loop
    End If
    
    'Clear form before bringing in new data
    Call UserForm_Initialize
    DoEvents
    
    
End Sub

Private Sub cmdPrevious_Click()
    'PURPOSE: Determine next entry row in register table depending on check box value
    
    Dim TopRow As Long
    Dim BtmRow As Long
    
    'Check if got StudyStatus is a valid array and in the case of checkbox if it contains Pre-commencement
    If RegTable.DataBodyRange Is Nothing Or (Tick And Not Contains(StudyStatus, "Pre-commencement")) Then
        Call cmdClear_Click
        errSearch.Caption = "No data found in register"
        Exit Sub
    End If
    
    'Repoint to RowIndex
     If IsArray(StudyStatus) Then
        TopRow = LBound(StudyStatus)
        BtmRow = UBound(StudyStatus)
    Else
        TopRow = 1
        BtmRow = 1
    End If
    
    If RowIndex < 0 Or RowIndex = TopRow Then
        RowIndex = BtmRow
    Else
        RowIndex = RowIndex - 1
    End If
    
    'Conditional stepping if check box ticked and Pre-commencement status in register
    'source: https://stackoverflow.com/questions/38267950/check-if-a-value-is-in-an-array-or-not-with-excel-vba
    If Tick And IsArray(StudyStatus) Then
        'Loop through study status array
        Do While InStr(1, "Pre-commencement", StudyStatus(RowIndex, 1), vbTextCompare) = 0
            RowIndex = RowIndex - 1
            
            If RowIndex < 1 Then
                RowIndex = BtmRow
            End If
        Loop
    End If
    
    'Clear form before bringing in new data
    Call UserForm_Initialize
    DoEvents
    
End Sub

Private Sub Read_Table()

    With RegTable.ListRows(RowIndex)
    
        Me.txtStudyName.Value = .Range(10).Value
        Me.txtProtocolNum.Value = .Range(9).Value
            
        'Check if site initiation visit passed and automatically reallocated status to commenced
        If .Range(112).Value <> vbNullString And String_to_Date(.Range(112).Value) < Now _
            And .Range(8).Value = "Pre-commencement" Then
            .Range(8).Value = "Commenced"
            
            'Update version control
            .Range(15).Value = Now
            .Range(16).Value = Username
            
            StudyStatus = RegTable.DataBodyRange.Columns(8)
        End If
            
        Me.txtSponsor.Value = .Range(11).Value
        Me.cboStudyStatus.Value = .Range(8).Value
        Me.cboStudyStatus.ForeColor = StudyStatus_Colour(.Range(10).Value)
        
        'Store value of old study status
        OldStudyStatus = Me.cboStudyStatus.Value
        
        'Access version control
        Call LogLastAccess
        
    End With
    
End Sub

Private Function StudyStatus_Colour(Status As String) As Long
    'PURPOSE: assigns RGB colour value depending on the Study Status
    Select Case (Status):
        Case "Pre-commencement"
            StudyStatus_Colour = RGB(0, 0, 0)
        Case "Commenced"
            StudyStatus_Colour = RGB(0, 128, 0)
        Case "Not Going Ahead"
            StudyStatus_Colour = RGB(255, 0, 255)
        Case "DELETED"
            StudyStatus_Colour = RGB(255, 0, 0)
    End Select
    
End Function

Private Function TransposeArray(InputArr As Variant, OutputArr As Variant) As Boolean
'PURPOSE: Transpose 2D array
'SOURCE: http://www.cpearson.com/excel/vbaarrays.htm

    Dim RowNdx As Long
    Dim ColNdx As Long
    Dim LB1 As Long
    Dim LB2 As Long
    Dim UB1 As Long
    Dim UB2 As Long
    
    '''''''''''''''''''''''''''''''''''''''
    ' Get the Lower and Upper bounds of
    ' InputArr.
    '''''''''''''''''''''''''''''''''''''''
    LB1 = LBound(InputArr, 1)
    LB2 = LBound(InputArr, 2)
    UB1 = UBound(InputArr, 1)
    UB2 = UBound(InputArr, 2)
    
    '''''''''''''''''''''''''''''''''''''''''
    ' Erase and ReDim OutputArr
    '''''''''''''''''''''''''''''''''''''''''
    Erase OutputArr
    ReDim OutputArr(LB2 To LB2 + UB2 - LB2, LB1 To LB1 + UB1 - LB1)
    
    For RowNdx = LBound(InputArr, 2) To UBound(InputArr, 2)
        For ColNdx = LBound(InputArr, 1) To UBound(InputArr, 1)
            OutputArr(RowNdx, ColNdx) = InputArr(ColNdx, RowNdx)
        Next ColNdx
    Next RowNdx
    
    TransposeArray = True

End Function

Private Function IsArrayEmpty(arr As Variant) As Boolean
'PURPOSE: Check if Array is empty
'SOURCE: http://www.cpearson.com/excel/vbaarrays.htm

Dim lb As Long
Dim ub As Long

err.Clear
On Error Resume Next
If IsArray(arr) = False Then
    ' we weren't passed an array, return True
    IsArrayEmpty = True
End If

' Attempt to get the UBound of the array. If the array is
' unallocated, an error will occur.
ub = UBound(arr, 1)
If (err.Number <> 0) Then
    IsArrayEmpty = True
Else
    ''''''''''''''''''''''''''''''''''''''''''
    ' On rare occassion, under circumstances I
    ' cannot reliably replictate, Err.Number
    ' will be 0 for an unallocated, empty array.
    ' On these occassions, LBound is 0 and
    ' UBound is -1.
    ' To accomodate the weird behavior, test to
    ' see if LB > UB. If so, the array is not
    ' allocated.
    ''''''''''''''''''''''''''''''''''''''''''
    err.Clear
    lb = LBound(arr)
    If lb > ub Then
        IsArrayEmpty = True
    Else
        IsArrayEmpty = False
    End If
End If

End Function


Private Function Contains(arr, v) As Boolean
'PURPOSE: Check if value is found in array
'Source: https://stackoverflow.com/questions/18754096/matching-values-in-string-array/18769246#18769246
Dim rv As Boolean, lb As Long, ub As Long, i As Long
    
    If IsArray(arr) Then
        lb = LBound(arr)
        ub = UBound(arr)
        For i = lb To ub
            If arr(i, 1) = v Then
                rv = True
                Exit For
            End If
        Next i
    ElseIf arr = v Then
        rv = True
    Else
        rv = False
    End If
    
    Contains = rv
End Function

Private Sub EraseIfArray(arr As Variant)
    'PURPOSE: Erase dynamic arrays
    
    If IsArray(arr) Then
        Erase arr
    End If
    
End Sub
