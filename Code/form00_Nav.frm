VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form00_Nav 
   Caption         =   "Vaccine Trial Study Start-up Tracker"
   ClientHeight    =   9096.001
   ClientLeft      =   -24
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
    cboList_StudyStatus = Array("Current", "Commenced", "Halted")
    StudyStatus = RegTable.DataBodyRange.Columns(8)
    
    'Clear user form
    'source: https://www.mrexcel.com/board/threads/loop-through-controls-on-a-userform.427103/
    For Each ctrl In Me.Controls
        Select Case True
                Case TypeOf ctrl Is MSForms.CheckBox
                    ctrl.value = Tick
                Case TypeOf ctrl Is MSForms.TextBox
                    ctrl.value = ""
                Case TypeOf ctrl Is MSForms.Label
                    'Empty error captions
                    If Left(ctrl.Name, 3) = "err" Then
                        ctrl.Caption = ""
                    End If
                Case TypeOf ctrl Is MSForms.ComboBox
                    ctrl.value = ""
                    ctrl.Clear
                Case TypeOf ctrl Is MSForms.ListBox
                    ctrl.value = ""
                    ctrl.Clear
            End Select
    Next ctrl
    
    'Fill combo box for study status
    For Each item In cboList_StudyStatus
        Me.cboStudyStatus.AddItem item
    Next item
    
    'Format fields
    Me.cboStudyStatus.TextAlign = fmTextAlignCenter
    Me.txtStudyName.TextAlign = fmTextAlignLeft
    Me.txtProtocolNum.TextAlign = fmTextAlignLeft
    Me.errSearch.TextAlign = fmTextAlignCenter
    
    If RowIndex > 0 Then
        Call Read_Table
        Me.cboStudyStatus.ForeColor = StudyStatus_Colour(Me.cboStudyStatus.value)
    End If
    
    'Unload search display
    Erase DisplayArr
    
End Sub


Private Sub cmdClose_Click()
    'PURPOSE: Closes current form
    
    'Access version control
    Call LogLastAccess
        
    Unload Me
    
    'Empty Array as no longer needed
    Erase StudyStatus
    Erase DisplayArr
    
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
    StudyName = Me.txtStudyName.value
    
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
        
        .Range(8) = "Current"
        .Range(9) = Me.txtProtocolNum.value
        .Range(10) = StudyName
        .Range(11) = Me.txtSponsor.value
        
        'Update version control
        .Range(15) = .Range(2).value
        .Range(16) = .Range(3).value
    End With
        
    Unload form00_Nav
    
    form01_StudyDetail.Show False
    
    'Empty Array as no longer needed
    Erase StudyStatus
    Erase DisplayArr
    
End Sub

Private Sub cbOnlyCurrent_Click()
    'PURPOSE: Change value of Tick variable
    Tick = Me.cbOnlyCurrent.value
End Sub

Private Sub cboStudyStatus_AfterUpdate()
    'PURPOSE: Change text color of combo box status based on value
    
    Dim SIVDate As String

    'Unique change events
    SIVDate = RegTable.DataBodyRange.Cells(RowIndex, 112).value
    
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
    If Me.cboStudyStatus.value = "Current" And SIVDate <> vbNullString And _
        String_to_Date(SIVDate) < Now Then
        
        Me.cboStudyStatus.value = "Commenced"
        
        'Clear Deletion Log
        With RegTable.ListRows(RowIndex)
            'Update version control
            .Range(15) = Now
            .Range(16) = Username
            
        End With
        
    End If
    
    'Update value in table
    RegTable.DataBodyRange.Cells(RowIndex, 8).value = Me.cboStudyStatus.value
    Me.cboStudyStatus.ForeColor = StudyStatus_Colour(Me.cboStudyStatus.value)
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
            .Range(15) = .Range(4).value
            .Range(16) = .Range(5).value
        End With
    
    
        'Change status
        With Me.cboStudyStatus
            .value = "DELETED"
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
        form08_ChangeLog.Show False
    End If
    
End Sub

Private Sub cmdReminders_Click()
    'PURPOSE: Open reminder log form
    
    If RowIndex < 0 Then
        errSearch.Caption = "Need study entry identified to view log"
    Else
        form09_ReminderLog.Show False
    End If
    
End Sub

Private Sub cmdEdit_Click()
    'PURPOSE: Closes current form and open Study Detail form
    
    If RowIndex < 0 Then
        errSearch.Caption = "Need study entry identified to proceed"
    Else
        
        'Write changes to register table
        With RegTable.ListRows(RowIndex)
            .Range(9) = Me.txtProtocolNum.value
            .Range(10) = Me.txtStudyName.value
            .Range(11) = Me.txtSponsor.value
            
            'Update version control
            .Range(15) = Now
            .Range(16) = Username
        End With
    
        Unload form00_Nav
        
        form01_StudyDetail.Show False
        
        'Empty Array as no longer needed
        Erase StudyStatus
        Erase DisplayArr
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
    
    SearchArr = RegTable.ListColumns(8).DataBodyRange.Resize(, 4)
    If IsArrayEmpty(SearchArr) Then Exit Sub
    
    j = 1
    
    Sponsor = Me.txtSponsor.value
    ProtocolNum = Me.txtProtocolNum.value
    StudyName = Me.txtStudyName.value
    
    
    For i = 1 To UBound(SearchArr)
        If (Not (Tick) Or (Tick And SearchArr(i, 1) = "Current")) And _
            (StudyName = vbNullString Or (Len(StudyName) > 0 And InStr(1, SearchArr(i, 3), StudyName) > 0)) And _
            (ProtocolNum = vbNullString Or (Len(ProtocolNum) > 0 And InStr(1, ProtocolNum, SearchArr(i, 2)) > 0)) And _
            (Sponsor = vbNullString Or (Len(Sponsor) > 0 And InStr(1, Sponsor, SearchArr(i, 4)) > 0)) Then

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
    
    If IsArrayEmpty(TempArr) Then Exit Sub
    
    'Transpose display array
    j = TransposeArray(TempArr, DisplayArr)
    
    Erase SearchArr
    Erase TempArr
    
    Me.lstSearch.List = DisplayArr
    
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
    
    'Check if got StudyStatus is a valid array and in the case of checkbox if it contains current
    If Not IsArray(StudyStatus) Or (Tick And Not Contains(StudyStatus, "Current")) Then
        errSearch.Caption = "No data found in register"
        Exit Sub
    End If
    
    RowIndex = UBound(StudyStatus)
    
    'Conditional stepping
    If Tick Then
        'Loop through study status array
        Do While InStr(1, StudyStatus(RowIndex, 1), "Current") = 0 And RowIndex > 1
            RowIndex = RowIndex - 1
        Loop
    End If
    
    'Clear form before bringing in new data
    Call UserForm_Initialize
    DoEvents
    
    
End Sub

Private Sub cmdNext_Click()
    'PURPOSE: Determine next entry row in register table depending on check box value

    'Check if got StudyStatus is a valid array and in the case of checkbox if it contains current
    If Not IsArray(StudyStatus) Or (Tick And Not Contains(StudyStatus, "Current")) Then
        errSearch.Caption = "No data found in register"
        Exit Sub
    End If
    
    'Repoint to RowIndex
    If RowIndex < 0 Or RowIndex = UBound(StudyStatus) Then
        RowIndex = 1
    Else
        RowIndex = RowIndex + 1
    End If
    
    'Conditional stepping
    If Tick Then
        'Loop through study status array
        Do While InStr(1, StudyStatus(RowIndex, 1), "Current") = 0
            RowIndex = RowIndex + 1
            If RowIndex > UBound(StudyStatus) Then
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
    
    'Check if got StudyStatus is a valid array and in the case of checkbox if it contains current
    If Not IsArray(StudyStatus) Or (Tick And Not Contains(StudyStatus, "Current")) Then
        errSearch.Caption = "No data found in register"
        Exit Sub
    End If
    
    RowIndex = LBound(StudyStatus)
    
    'Conditional stepping
    If Tick Then
        'Loop through study status array
        Do While InStr(1, StudyStatus(RowIndex, 1), "Current") = 0 And RowIndex < UBound(StudyStatus)
            RowIndex = RowIndex + 1
        Loop
    End If
    
    'Clear form before bringing in new data
    Call UserForm_Initialize
    DoEvents
    
    
End Sub

Private Sub cmdPrevious_Click()
    'PURPOSE: Determine next entry row in register table depending on check box value
    
    'Check if got StudyStatus is a valid array and in the case of checkbox if it contains current
    If Not IsArray(StudyStatus) Or (Tick And Not Contains(StudyStatus, "Current")) Then
        errSearch.Caption = "No data found in register"
        Exit Sub
    End If
    
    'Repoint to RowIndex
    If RowIndex < 0 Or RowIndex = LBound(StudyStatus) Then
        RowIndex = UBound(StudyStatus)
    Else
        RowIndex = RowIndex - 1
    End If
    
    'Conditional stepping if check box ticked and Current status in register
    'source: https://stackoverflow.com/questions/38267950/check-if-a-value-is-in-an-array-or-not-with-excel-vba
    If Tick Then
        'Loop through study status array
        Do While InStr(1, StudyStatus(RowIndex, 1), "Current") = 0
            RowIndex = RowIndex - 1
            
            If RowIndex < 1 Then
                RowIndex = UBound(StudyStatus)
            End If
        Loop
    End If
    
    'Clear form before bringing in new data
    Call UserForm_Initialize
    DoEvents
    
End Sub

Private Sub Read_Table()

    With RegTable.ListRows(RowIndex)
    
        Me.txtStudyName.value = .Range(10).value
        Me.txtProtocolNum.value = .Range(9).value
            
        'Check if site initiation visit passed and automatically reallocated status to commenced
        If .Range(112).value <> vbNullString And String_to_Date(.Range(112).value) < Now _
            And .Range(8).value = "Current" Then
            .Range(8).value = "Commenced"
            
            'Update version control
            .Range(15).value = Now
            .Range(16).value = Username
            
            StudyStatus = RegTable.DataBodyRange.Columns(8)
        End If
            
        Me.txtSponsor.value = .Range(11).value
        Me.cboStudyStatus.value = .Range(8).value
        Me.cboStudyStatus.ForeColor = StudyStatus_Colour(.Range(10).value)
        
        'Store value of old study status
        OldStudyStatus = Me.cboStudyStatus.value
        
        'Access version control
        Call LogLastAccess
        
    End With
    
End Sub

Private Function StudyStatus_Colour(Status As String) As Long
    'PURPOSE: assigns RGB colour value depending on the Study Status
    Select Case (Status):
        Case "Current"
            StudyStatus_Colour = RGB(0, 0, 0)
        Case "Commenced"
            StudyStatus_Colour = RGB(0, 128, 0)
        Case "Halted"
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
    Else
        rv = False
    End If
    
    Contains = rv
End Function
