VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form00_Nav 
   Caption         =   "Vaccine Trial Study Start-up Tracker"
   ClientHeight    =   8040
   ClientLeft      =   0
   ClientTop       =   -264
   ClientWidth     =   10092
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
    Me.StartUpPosition = 0
    Me.Top = Application.Top + 25
    Me.Left = Application.Left + 25
    Me.Height = UHeight
    Me.Width = UWidth

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
                    ctrl.Value = Tick
                Case TypeOf ctrl Is MSForms.TextBox
                    ctrl.Value = ""
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
    
    'Format fields
    Me.cboStudyStatus.TextAlign = fmTextAlignCenter
    Me.txtStudyName.TextAlign = fmTextAlignLeft
    Me.txtProtocolNum.TextAlign = fmTextAlignLeft
    Me.errSearch.TextAlign = fmTextAlignCenter
    Me.errSearch.Caption = vbNullString
    
    If RowIndex > 0 Then
        Call Read_Table
    End If
    
End Sub


Private Sub cmdClose_Click()
    'PURPOSE: Closes current form
    Unload Me
    
    'Empty Array as no longer needed
    Erase StudyStatus
    
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
        .Range(2) = Now
        .Range(3) = Username
        .Range(8) = "Current"
        .Range(9) = Me.txtProtocolNum.Value
        .Range(10) = StudyName
        .Range(11) = Me.txtSponsor.Value
        .Range(14) = .Range(2).Value
        .Range(15) = .Range(3).Value
    End With
    
    Unload form00_Nav
    
    form01_StudyDetail.Show False
    
    'Empty Array as no longer needed
    Erase StudyStatus
    
End Sub

Private Sub cbOnlyCurrent_Click()
    'PURPOSE: Change value of Tick variable
    Tick = Me.cbOnlyCurrent.Value
End Sub

Private Sub cboStudyStatus_Change()
    'PURPOSE: Change text color of combo box status based on value
    Me.cboStudyStatus.ForeColor = StudyStatus_Colour(Me.cboStudyStatus.Value)
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
            .Range(4) = Now
            .Range(5) = Username
            .Range(8) = "DELETED"
        End With
    
    
        'Change status
        With Me.cboStudyStatus
            .Value = "DELETED"
            .ForeColor = vbRed
        End With
        
    End If
    
End Sub

Private Sub cmdEdit_Click()
    'PURPOSE: Closes current form and open Study Detail form
    Unload form00_Nav
    
    form01_StudyDetail.Show False
    
    'Empty Array as no longer needed
    Erase StudyStatus
    
End Sub

Private Sub cmdSearch_Click()
    'PURPOSE: Populate list box with keyword search results
    'SOURCE: https://stackoverflow.com/questions/45356240/vba-for-selecting-a-number-of-columns-in-an-excel-table
    
    Dim Sponsor As String
    Dim ProtocolNum As String
    Dim SearchArr As Variant, TempArr() As Variant
    Dim SearchStatus As String
    Dim i As Integer, j As Integer
    
    SearchArr = RegTable.ListColumns(8).DataBodyRange.Resize(, 4)
    If IsArrayEmpty(SearchArr) Then Exit Sub
    
    j = 1
    
    Sponsor = Me.txtSponsor.Value
    ProtocolNum = Me.txtProtocolNum.Value
    StudyName = Me.txtStudyName.Value
    
    
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
    
    SearchArr = DisplayArr
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
    
    Dim SearchStr As String
    
    If Not IsArray(StudyStatus) Then Exit Sub
    
    RowIndex = UBound(StudyStatus)
    
    'Conditional stepping
    If Tick Then
        
        SearchStr = "Current"
        'Loop through study status array
        Do While InStr(1, StudyStatus(RowIndex, 1), SearchStr) = 0 And RowIndex > 1
            RowIndex = RowIndex - 1
        Loop
    End If
    
    'Clear form before bringing in new data
    Call UserForm_Initialize
    DoEvents
    
    
End Sub

Private Sub cmdNext_Click()
    'PURPOSE: Determine next entry row in register table depending on check box value

    Dim SearchStr As String
    
    If Not IsArray(StudyStatus) Then Exit Sub
    
    'Repoint to RowIndex
    If RowIndex < 0 Or RowIndex = UBound(StudyStatus) Then
        RowIndex = 1
    Else
        RowIndex = RowIndex + 1
    End If
    
    'Conditional stepping
    If Tick Then
        SearchStr = "Current"
        'Loop through study status array
        Do While InStr(1, StudyStatus(RowIndex, 1), SearchStr) = 0
            RowIndex = RowIndex + 1
        Loop
    End If
        
    'Clear form before bringing in new data
    Call UserForm_Initialize
    DoEvents
    
End Sub

    
Private Sub cmdJumpBack_Click()
    'PURPOSE: Redirect to newest
    
    Dim SearchStr As String
    
    If Not IsArray(StudyStatus) Then Exit Sub
    
    RowIndex = LBound(StudyStatus)
    
    'Conditional stepping
    If Tick Then
        SearchStr = "Current"
        'Loop through study status array
        Do While InStr(1, StudyStatus(RowIndex, 1), SearchStr) = 0 And RowIndex < UBound(StudyStatus)
            RowIndex = RowIndex + 1
        Loop
    End If
    
    'Clear form before bringing in new data
    Call UserForm_Initialize
    DoEvents
    
    
End Sub

Private Sub cmdPrevious_Click()
    'PURPOSE: Determine next entry row in register table depending on check box value

    Dim SearchStr As String
    
    If Not IsArray(StudyStatus) Then Exit Sub
    
    'Repoint to RowIndex
    If RowIndex < 0 Or RowIndex = LBound(StudyStatus) Then
        RowIndex = UBound(StudyStatus)
    Else
        RowIndex = RowIndex - 1
    End If
    
    'Conditional stepping
    If Tick Then
        SearchStr = "Current"
        'Loop through study status array
        Do While InStr(1, StudyStatus(RowIndex, 1), SearchStr) = 0
            RowIndex = RowIndex - 1
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
        Me.cboStudyStatus.Value = .Range(8).Value
        Me.txtSponsor.Value = .Range(11).Value
        Me.cboStudyStatus.ForeColor = StudyStatus_Colour(.Range(8).Value)
    End With
    
End Sub

Private Function StudyStatus_Colour(status As String) As Long
    'PURPOSE: assigns RGB colour value depending on the Study Status
    Select Case (status):
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

Private Function IsArrayEmpty(Arr As Variant) As Boolean
'PURPOSE: Check if Array is empty
'SOURCE: http://www.cpearson.com/excel/vbaarrays.htm

Dim LB As Long
Dim UB As Long

Err.Clear
On Error Resume Next
If IsArray(Arr) = False Then
    ' we weren't passed an array, return True
    IsArrayEmpty = True
End If

' Attempt to get the UBound of the array. If the array is
' unallocated, an error will occur.
UB = UBound(Arr, 1)
If (Err.Number <> 0) Then
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
    Err.Clear
    LB = LBound(Arr)
    If LB > UB Then
        IsArrayEmpty = True
    Else
        IsArrayEmpty = False
    End If
End If

End Function
