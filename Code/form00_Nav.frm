VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form00_Nav 
   Caption         =   "Vaccine Trial Study Start-up Tracker"
   ClientHeight    =   6720
   ClientLeft      =   60
   ClientTop       =   120
   ClientWidth     =   14580
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
    
    cboList_StudyStatus = Array("Current", "Commenced", "Halted")
    
    'Clear user form
    'source: https://www.mrexcel.com/board/threads/loop-through-controls-on-a-userform.427103/
    For Each ctrl In Me.Controls
        Select Case True
                Case TypeOf ctrl Is MSForms.CheckBox
                    ctrl.Value = False
                Case TypeOf ctrl Is MSForms.TextBox
                    ctrl.Value = ""
                Case TypeOf ctrl Is MSForms.ComboBox
                    ctrl.Value = "Current"
                    ctrl.Clear
                Case TypeOf ctrl Is MSForms.ListBox
                    ctrl.Value = ""
                    ctrl.Clear
            End Select
    Next ctrl
    
    'Fill combo box for study status
    For Each item In cboList_StudyStatus
        cboStudyStatus.AddItem item
    Next item

End Sub


Private Sub cmdClose_Click()
    'PURPOSE: Closes current form
    Unload Me
    
End Sub

Private Sub cmdNew_Click()
    'PURPOSE: Closes current form and open Study Detail form
    Unload form00_Nav
    
    form01_StudyDetail.Show False
End Sub

Private Sub cmdEdit_Click()
    'PURPOSE: Closes current form and open Study Detail form
    Unload form00_Nav
    
    form01_StudyDetail.Show False
End Sub






'Private Sub cboPrjChampTeam_AfterUpdate()
'    'Fill dependent combo box for project champion team members
'    Dim cPrjChamp As Range
'    Dim ws As Worksheet
'    Dim cboIndex As Integer
'    Dim strPrjChamp As String
'    Set ws = Worksheets("Lookup Lists")
'
'    'Fill Project Champion Team Persons combo box
'    'source: https://www.excel-easy.com/vba/examples/dependent-combo-boxes.html
'    cboIndex = cboPrjChampTeam.ListIndex
'
'    If cboIndex = 0 Then strPrjChamp = "TS_Team" Else strPrjChamp = "GI_Team"
'
'    cboPrjChamp.Clear
'
'    For Each cPrjChamp In ws.Range(strPrjChamp)
'      With Me.MultiPage1.tbPage3.cboPrjChamp
'        .AddItem cPrjChamp.Value
'      End With
'    Next cPrjChamp
'
'End Sub
'
'Private Sub cboPrjStatus_AfterUpdate()
'    'Copy Project Status to Main Page label
'    lblPrjStatusCapt.Caption = cboPrjStatus.Value
'
'End Sub
'
'
'Private Sub txtSF_Change()
''PURPOSE: Limit TextBox inputs to Postive Whole Numbers
''SOURCE: www.TheSpreadsheetGuru.com/the-code-vault
'
'    Dim myerror As Integer
'
'    If Not IsNumeric(txtSF.Text) And txtSF.Text <> "" Then
'      If Len(txtSF.Text) > 1 Then
'        'Remove Added Non-Numerical Character from Number
'          txtSF.Text = Abs(Round(Left(txtSF.Text, Len(txtSF.Text) - 1), 0))
'      Else
'        'Delete Single Non-Numerical Character
'          txtSF.Text = ""
'      End If
'    ElseIf txtSF.Text <> "" Then
'      'Ensure Positive and No Decimals
'        txtSF.Text = Abs(Round(txtSF.Text, 0))
'    End If
'
'    'Remove last digit if >100
'    If txtSF.Value > 100 And txtSF.Text <> "" Then
'        myerror = MsgBox("Error! Value cannot exceed 100", vbOKOnly, "WARNING!")
'        txtSF.Text = Left(txtSF.Text, 2)
'        Me.MultiPage1.tbPage1.txtSF.SetFocus
'    End If
'
'End Sub
'
'
''Code below pertains to Search Button and Toggles
''-------------------------------------------------
''Source: https://www.onlinepclearning.com/add-previous-and-next-buttons-userform-excel-vba/
'
'Private Sub cmdSearch_Click()
'    'Search for row in register with Project ID
'    Dim rowsearch As Range
'    Dim strSearch As String
'    Dim myerror As Integer
'
'    'Convert numbers to Allowed Project IDs
'    If IsNumeric(Me.txtPrjID.Value) Then
'        strSearch = "PRJ-" & Format(Me.txtPrjID.Value, "000000")
'        Me.txtPrjID.Value = strSearch
'    Else
'        strSearch = Me.txtPrjID.Value
'    End If
'
'    'error block
'    On Error GoTo ErrHandler:
'
'    'find cell with Project ID in register range
'    Set rowsearch = Sheets("Register").Range("Register[PROJECT ID]").Find(What:=strSearch, LookIn:=xlValues)
'
'    'Read values from register sheet
'    '--------------------------------------------
'    Call Read_from_sheet(rowsearch)
'
'    Exit Sub
'
'    'error block
'ErrHandler:
'        myerror = MsgBox("Error! Project Not Found - Check Project ID", vbOKOnly, "WARNING!")
'        Call UserForm_Initialize
'End Sub
'
'Private Sub cmdJumpBack_Click()
'
'    Dim rowsearch As Range
'    Dim cnt As Integer
'    Dim strSearch As String
'    Dim Jump As Integer
'    Dim TopRow As Long
'    Dim cRow As Long
'    Dim nRow As Long
'    Dim myerror As Integer
'
'    'Set Toggle interval and variables
'    Jump = 5
'    cnt = 0
'    TopRow = Sheets("Register").Range("Register[[#Headers],[PROJECT ID]]").Row + 1
'
'    'error block
'    On Error GoTo ErrHandler:
'
'    'Convert numbers to Allowed Project IDs
'    If Me.txtPrjID.Value = "" Then
'        Me.txtPrjID.Value = Sheets("Register").Cells(TopRow, 2).Value
'        cmdSearch_Click
'        Exit Sub
'    ElseIf IsNumeric(Me.txtPrjID.Value) Then
'        strSearch = "PRJ-" & Format(Me.txtPrjID.Value, "000000")
'        Me.txtPrjID.Value = strSearch
'    Else
'        strSearch = Me.txtPrjID.Value
'    End If
'
'    'find cell with Project ID in register range
'    Set rowsearch = Sheets("Register").Range("Register[PROJECT ID]").Find(What:=strSearch, LookIn:=xlValues)
'
'    'Loop back counting rows until Jump interval
'    cRow = rowsearch.Row
'    While cRow > (TopRow - 1) And nRow <= Jump:
'
'        cRow = rowsearch.Offset(-cnt, 0).Row
'        If Not (IsEmpty(rowsearch.Offset(-cnt, 0))) Then
'            nRow = nRow + 1
'        End If
'
'        'Break out of loop if Top Row reached
'        If cRow = TopRow Then
'            cnt = rowsearch.Row - TopRow + 1
'            nRow = Jump + 1
'        Else
'            cnt = cnt + 1
'        End If
'    Wend
'
'    'Redefine range selected from register
'    Set rowsearch = rowsearch.Offset(-cnt + 1, 0)
'
'
'    'Read values from register sheet
'    '--------------------------------------------
'    Call Read_from_sheet(rowsearch)
'
'    Exit Sub
'
'    'error block
'ErrHandler:
'        myerror = MsgBox("Error! Project Not Found - Check Project ID", vbOKOnly, "WARNING!")
'
'End Sub
'
'
'Private Sub cmdJumpForw_Click()
'
'    Dim rowsearch As Range
'    Dim cnt As Integer
'    Dim strSearch As String
'    Dim Jump As Integer
'    Dim BtmRow As Long
'    Dim cRow As Long
'    Dim nRow As Long
'    Dim myerror As Integer
'
'    'Set Toggle interval and variables
'    Jump = 5
'    cnt = 0
'
'    'Find last used Row in register sheet
'    'Source: https://www.contextures.com/rickrothsteinexcelvbasheet.html
'    BtmRow = Sheets("Register").Range("Register[PROJECT ID]").Rows.Count + Sheets("Register").Range("Register[[#Headers],[PROJECT ID]]").Row
'
'    'error block
'    On Error GoTo ErrHandler:
'
'    'Convert numbers to Allowed Project IDs
'    If Me.txtPrjID.Value = "" Then
'        Me.txtPrjID.Value = Sheets("Register").Cells(BtmRow, 2).Value
'        cmdSearch_Click
'        Exit Sub
'    ElseIf IsNumeric(Me.txtPrjID.Value) Then
'        strSearch = "PRJ-" & Format(Me.txtPrjID.Value, "000000")
'        Me.txtPrjID.Value = strSearch
'    Else
'        strSearch = Me.txtPrjID.Value
'    End If
'
'    'find cell with Project ID in register range
'    Set rowsearch = Sheets("Register").Range("Register[PROJECT ID]").Find(What:=strSearch, LookIn:=xlValues)
'
'    'Loop back counting rows until Jump interval
'    cRow = rowsearch.Row
'    While cRow < (BtmRow + 1) And nRow <= Jump:
'
'        cRow = rowsearch.Offset(cnt, 0).Row
'        If Not (IsEmpty(rowsearch.Offset(cnt, 0))) Then
'            nRow = nRow + 1
'        End If
'
'        'Break out of loop if Bottom Row Reached
'        If cRow = BtmRow Then
'            cnt = BtmRow + 1 - rowsearch.Row
'            nRow = Jump + 1
'        Else
'            cnt = cnt + 1
'        End If
'    Wend
'
'    'Redefine range selected from register
'    Set rowsearch = rowsearch.Offset(cnt - 1, 0)
'
'
'    'Read values from register sheet
'    '--------------------------------------------
'    Call Read_from_sheet(rowsearch)
'
'    Exit Sub
'
'    'error block
'ErrHandler:
'        myerror = MsgBox("Error! Project Not Found - Check Project ID", vbOKOnly, "WARNING!")
'
'End Sub
'
'Private Sub cmdPrevious_Click()
'
'    Dim rowsearch As Range
'    Dim cnt As Integer
'    Dim strSearch As String
'    Dim Jump As Integer
'    Dim TopRow As Long
'    Dim cRow As Long
'    Dim nRow As Long
'    Dim myerror As Integer
'
'    'Set Toggle interval and variables
'    Jump = 1
'    cnt = 0
'    TopRow = Sheets("Register").Range("Register[[#Headers],[PROJECT ID]]").Row + 1
'
'    'error block
'    On Error GoTo ErrHandler:
'
'    'Convert numbers to Allowed Project IDs
'    If Me.txtPrjID.Value = "" Then
'        Me.txtPrjID.Value = Sheets("Register").Cells(TopRow, 2).Value
'        cmdSearch_Click
'        Exit Sub
'    ElseIf IsNumeric(Me.txtPrjID.Value) Then
'        strSearch = "PRJ-" & Format(Me.txtPrjID.Value, "000000")
'        Me.txtPrjID.Value = strSearch
'    Else
'        strSearch = Me.txtPrjID.Value
'    End If
'
'    'find cell with Project ID in register range
'    Set rowsearch = Sheets("Register").Range("Register[PROJECT ID]").Find(What:=strSearch, LookIn:=xlValues)
'
'    'Loop back counting rows until Jump interval
'    cRow = rowsearch.Row
'    While cRow > (TopRow - 1) And nRow <= Jump:
'
'        cRow = rowsearch.Offset(-cnt, 0).Row
'        If Not (IsEmpty(rowsearch.Offset(-cnt, 0))) Then
'            nRow = nRow + 1
'        End If
'
'        'Break out of loop if Top Row reached
'        If cRow = TopRow Then
'            cnt = rowsearch.Row - TopRow + 1
'            nRow = Jump + 1
'        Else
'            cnt = cnt + 1
'        End If
'    Wend
'
'    'Redefine range selected from register
'    Set rowsearch = rowsearch.Offset(-cnt + 1, 0)
'
'
'    'Read values from register sheet
'    '--------------------------------------------
'    Call Read_from_sheet(rowsearch)
'
'    Exit Sub
'
'    'error block
'ErrHandler:
'        myerror = MsgBox("Error! Project Not Found - Check Project ID", vbOKOnly, "WARNING!")
'
'End Sub
'
'Private Sub cmdNext_Click()
'
'    Dim rowsearch As Range
'    Dim cnt As Integer
'    Dim strSearch As String
'    Dim Jump As Integer
'    Dim BtmRow As Long
'    Dim cRow As Long
'    Dim nRow As Long
'    Dim myerror As Integer
'
'    'Set Toggle interval and variables
'    Jump = 1
'    cnt = 0
'
'    'Find last used Row in register sheet
'    'Source: https://www.contextures.com/rickrothsteinexcelvbasheet.html
'    BtmRow = Sheets("Register").Range("Register[PROJECT ID]").Rows.Count + Sheets("Register").Range("Register[[#Headers],[PROJECT ID]]").Row
'
'    'error block
'    On Error GoTo ErrHandler:
'
'    'Convert numbers to Allowed Project IDs
'    If Me.txtPrjID.Value = "" Then
'        Me.txtPrjID.Value = Sheets("Register").Cells(BtmRow, 2).Value
'        cmdSearch_Click
'        Exit Sub
'    ElseIf IsNumeric(Me.txtPrjID.Value) Then
'        strSearch = "PRJ-" & Format(Me.txtPrjID.Value, "000000")
'        Me.txtPrjID.Value = strSearch
'    Else
'        strSearch = Me.txtPrjID.Value
'    End If
'
'    'find cell with Project ID in register range
'    Set rowsearch = Sheets("Register").Range("Register[PROJECT ID]").Find(What:=strSearch, LookIn:=xlValues)
'
'    'Loop back counting rows until Jump interval
'    cRow = rowsearch.Row
'    While cRow < (BtmRow + 1) And nRow <= Jump:
'
'        cRow = rowsearch.Offset(cnt, 0).Row
'        If Not (IsEmpty(rowsearch.Offset(cnt, 0))) Then
'            nRow = nRow + 1
'        End If
'
'        'Break out of loop if Bottom Row Reached
'        If cRow = BtmRow Then
'            cnt = BtmRow + 1 - rowsearch.Row
'            nRow = Jump + 1
'        Else
'            cnt = cnt + 1
'        End If
'    Wend
'
'    'Redefine range selected from register
'    Set rowsearch = rowsearch.Offset(cnt - 1, 0)
'
'
'    'Read values from register sheet
'    '--------------------------------------------
'    Call Read_from_sheet(rowsearch)
'
'    Exit Sub
'
'    'error block
'ErrHandler:
'        myerror = MsgBox("Error! Project Not Found - Check Project ID", vbOKOnly, "WARNING!")
'
'End Sub
'
'
'
''Code below pertains to selecting start and end dates
''----------------------------------------------------
'
'Private Sub cmdEndDate_Click()
'    'Run date picker if command button clicked
'    frmCalendar.lblCtrlName = "txtEndDate"
'    frmCalendar.lblUF = "frmProject"
'    frmCalendar.Show
'End Sub
'
'Private Sub cmdStartDate_Click()
'    'Run date picker if command button clicked
'    frmCalendar.lblCtrlName = "txtStartDate"
'    frmCalendar.lblUF = "frmProject"
'    frmCalendar.Show
'End Sub
'
'
'Private Sub txtStartDate_AfterUpdate()
'    'Check if date entered manually is valid format
'    Dim myerror As Integer
'
'    If txtStartDate.Value <> "" Then
'        If Not (IsDate(txtStartDate.Value)) Then
'            myerror = MsgBox("Enter valid date DD/MM/YYY", vbOKOnly, "WARNING!")
'            txtStartDate.Value = Null
'        End If
'    End If
'
'End Sub
'
'Private Sub txtEndDate_AfterUpdate()
'    'Check if date entered manually is valid format
'    Dim myerror As Integer
'
'    If txtEndDate.Value <> "" Then
'        If Not (IsDate(txtEndDate.Value)) Then
'            myerror = MsgBox("Enter valid date DD/MM/YYY", vbOKOnly, "WARNING!")
'            txtEndDate.Value = Null
'        End If
'    End If
'
'    'Change status to red tinge if over due
'    Me.lblPrjStatusCapt.BackColor = &HFFFFFF
'    Me.txtPrjTitle.BackColor = &HFFFFFF
'
'    If Me.MultiPage1.tbPage2.txtEndDate.Value <> "" Then
'        If CDate(Me.MultiPage1.tbPage2.txtEndDate.Value) < Date And Me.lblPrjStatusCapt.Caption <> "Monitor" Then
'            Me.lblPrjStatusCapt.BackColor = &HC0C0FF
'            Me.txtPrjTitle.BackColor = &HC0C0FF
'        End If
'    End If
'
'End Sub
'
'Private Sub txtStartDate_Change()
'    'Check if start date is less than end date
'    Dim SDateBool As Boolean
'    Dim EDateBool As Boolean
'    Dim myerror As Integer
'
'    SDateBool = IsDate(txtStartDate.Value)
'    EDateBool = IsDate(txtEndDate.Value)
'
''    'Check if dates are chronologically correct
''    'If not replace with today's date
''    If SDateBool And SDateBool = EDateBool Then
''        If DateDiff("d", txtStartDate.Value, txtEndDate.Value) < 0 Then
''            myerror = MsgBox("Start Date cannot be later than End Date", vbOKOnly, "WARNING!")
''            txtStartDate.Value = Date
''        End If
''    End If
'
'End Sub
'
'Private Sub txtEndDate_Change()
'    'Check if start date is less than end date
'    Dim SDateBool As Boolean
'    Dim EDateBool As Boolean
'    Dim myerror As Integer
'
'    SDateBool = IsDate(txtStartDate.Value)
'    EDateBool = IsDate(txtEndDate.Value)
'
''    'Check if dates are chronologically correct
''    'If not replace with Start Date
''    If EDateBool And EDateBool = SDateBool Then
''        If DateDiff("d", txtStartDate.Value, txtEndDate.Value) < 0 Then
''            myerror = MsgBox("End Date cannot be earlier than Start Date", vbOKOnly, "WARNING!")
''            txtEndDate.Value = txtStartDate.Value
''        End If
''    End If
'
'    'Change status to red tinge if over due
'    Me.lblPrjStatusCapt.BackColor = &HFFFFFF
'    Me.txtPrjTitle.BackColor = &HFFFFFF
'
'    If IsDate(Me.MultiPage1.tbPage2.txtEndDate.Value) Then
'        If CDate(Me.MultiPage1.tbPage2.txtEndDate.Value) < Date And Me.lblPrjStatusCapt.Caption <> "Monitor" Then
'            Me.lblPrjStatusCapt.BackColor = &HC0C0FF
'            Me.txtPrjTitle.BackColor = &HC0C0FF
'        End If
'    End If
'
'End Sub
'
''Code Below Pertains to Main Button controls
''--------------------------------------------
''source: https://www.contextures.com/xlUserForm02.html
'
'Private Sub cmdClose_Click()
'    'Closes Form - doesn't save data
'    Unload Me
'End Sub
'
'
'Private Sub cmdNew_Click()
'    'Adds New Project entry into register
'    'Note - Project ID cannot be customised
'    Dim rowsearch As Range
'    Dim BtmRow As Long
'    Dim LastID As Long
'    Dim NewID As String
'    Dim ws As Worksheet
'
'    Application.ScreenUpdating = False
'    Application.Calculation = xlCalculationManual
'
'    Set ws = Worksheets("Register")
'
'    'Find last used Row in register sheet
'    'Source: https://www.contextures.com/rickrothsteinexcelvbasheet.html
'    BtmRow = Sheets("Register").Range("Register[PROJECT ID]").Rows.Count + Sheets("Register").Range("Register[[#Headers],[PROJECT ID]]").Row
'
'     'Create new Project ID
'     If BtmRow = 6 And Sheets("Register").Cells(BtmRow, 2).Value = "" Then
'        NewID = "PRJ-000000"
'        BtmRow = BtmRow - 1
'     Else
'        LastID = Right(Sheets("Register").Cells(BtmRow, 2).Value, 6)
'        NewID = "PRJ-" & Format(LastID + 1, "000000")
'     End If
'
'     'find cell with Project ID in register range
'    Set rowsearch = Sheets("Register").Cells(BtmRow + 1, 2)
'
'    'Write values from register sheet
'    '--------------------------------------------
'    Sheets("Register").Unprotect
'
'    rowsearch.Value = NewID
'
'    Call Write_to_Sheet(rowsearch)
'
'    Sheets("Register").Protect
'
'    'Replace Project ID with newly generated Project ID
'    Me.txtPrjID.Value = NewID
'
'    'Refresh Pivots
'    Call Report_Builder
'
'    Application.ScreenUpdating = True
'    Application.Calculation = xlCalculationAutomatic
'End Sub
'
'
'Private Sub cmdChange_Click()
'    'Apply changes to existing project data
'    Dim rowsearch As Range
'    Dim strSearch As String
'    Dim myerror As Integer
'
'    Application.ScreenUpdating = False
'    Application.Calculation = xlCalculationManual
'
'    'error block
'    On Error GoTo ErrHandler:
'
'    'find cell with Project ID in register range
'    strSearch = Me.txtPrjID.Value
'    Set rowsearch = Sheets("Register").Range("Register[PROJECT ID]").Find(What:=strSearch, LookIn:=xlValues)
'
'    'Write values from register sheet
'    '--------------------------------------------
'    Sheets("Register").Unprotect
'
'    Call Write_to_Sheet(rowsearch)
'
'    Sheets("Register").Protect
'
'    'Refresh Pivots
'    Call Report_Builder
'
'    Application.ScreenUpdating = True
'    Application.Calculation = xlCalculationAutomatic
'
'    Exit Sub
'
'    'error block
'ErrHandler:
'        myerror = MsgBox("Error! Project Not Found - Check Project ID", vbOKOnly, "WARNING!")
'
'    Application.ScreenUpdating = True
'    Application.Calculation = xlCalculationAutomatic
'End Sub
'
'
'Private Sub cmdDelete_Click()
'    'Permanently deletes existing project entry
'    Dim rowsearch As Range
'    Dim strSearch As String
'    Dim nCol As Long
'    Dim Confirm As Integer
'    Dim ws As Worksheet
'    Dim myerror As Integer
'    Dim DelRow As Long
'
'    Application.ScreenUpdating = False
'    Application.Calculation = xlCalculationManual
'
'    Set ws = Worksheets("Register")
'
'    'error block
'    On Error GoTo ErrHandler:
'
'    'Confirm deletion
'    Confirm = MsgBox("Are you sure you want to delete Project data?", vbYesNo, "WARNING!")
'
'    'If select no then cancel deletion
'    If Confirm = vbNo Then
'        Exit Sub
'    End If
'
'    ws.Unprotect
'
'    'count number of columns in used range
'    'Source: https://www.contextures.com/rickrothsteinexcelvbasheet.html
'    nCol = ws.Range("Register[#Headers]").Columns.Count
'
'    'find cell with Project ID in register range
'    strSearch = Me.txtPrjID.Value
'    Set rowsearch = Sheets("Register").Range("Register[PROJECT ID]").Find(What:=strSearch, LookIn:=xlValues)
'
'    DelRow = rowsearch.Row - 1
'
'    'Delete Row Contents
'
'    rowsearch.Resize(1, nCol - 1).Delete Shift:=xlUp
'
'    'Replace project ID for search
'    Me.txtPrjID.Value = ws.Cells(DelRow, 2).Value
'
'    ws.Protect
'
'    'Display last project
'    Call cmdSearch_Click
'
'    'Refresh Pivots
'    Call Report_Builder
'
'    Exit Sub
'
'    'error block
'ErrHandler:
'        myerror = MsgBox("Error! Project Not Found - Check Project ID", vbOKOnly, "WARNING!")
'
'    Application.ScreenUpdating = True
'    Application.Calculation = xlCalculationAutomatic
'
'End Sub
'
'Private Sub cmdClear_Click()
'    'Clear complete form
'    Call UserForm_Initialize
'End Sub
'
'
'Sub ReadStringToList(lstObj As MSForms.ListBox, xStr As String)
'
'    'Read string into list box
'    'source: https://www.contextures.com/excel-data-validation-listbox.html
'    Dim lCountList As Long
'    Dim StrArr As Variant
'    Dim lStrCnt As Long
'
'    StrArr = Split(xStr, vbLf)
'
'    With lstObj
'      'Clear current selections in listbox
'      For lCountList = 0 To .ListCount - 1
'        .Selected(lCountList) = False
'      Next lCountList
'
'      'Select new items in listbox
'      For lCountList = 0 To .ListCount - 1
'        For lStrCnt = LBound(StrArr) To UBound(StrArr)
'            If CStr(.List(lCountList)) = StrArr(lStrCnt) Then
'              'On Error GoTo errHandler
'              .Selected(lCountList) = True
'              Exit For
'            End If
'        Next lStrCnt
'      Next lCountList
'    End With
'
'End Sub
'
'
'Sub Write_to_Sheet(rowsearch As Range)
'    'Write values to spreadsheet register tab
'
'    Dim strSelItems As String
'    Dim lCountList As Long
'    Dim strSep As String
'    Dim strAdd As String
'    Dim StartDate As Date
'    Dim EndDate As Date
'
'    strSep = vbLf
'
'    'Write values from register sheet
'    '--------------------------------------------
'    'Main Page
'    rowsearch.Offset(0, 1).Value = Me.txtPrjTitle.Value
'    rowsearch.Offset(0, 2).Value = Me.cbMET.Value
'
'    'Page 1
'    rowsearch.Offset(0, 5).Value = Me.MultiPage1.tbPage1.cboPrjOrigTeam.Value
'    rowsearch.Offset(0, 6).Value = Me.MultiPage1.tbPage1.txtPrjOrigPerson.Value
'    rowsearch.Offset(0, 7).Value = Me.MultiPage1.tbPage1.txtPrjContext.Value
'    rowsearch.Offset(0, 8).Value = Me.MultiPage1.tbPage1.txtPrjGoal.Value
'    rowsearch.Offset(0, 9).Value = Me.MultiPage1.tbPage1.cboImprovType.Value
'    rowsearch.Offset(0, 10).Value = Me.MultiPage1.tbPage1.txtAB.Value
'    rowsearch.Offset(0, 11).Value = Me.MultiPage1.tbPage1.cbVerified.Value
'    rowsearch.Offset(0, 12).Value = Me.MultiPage1.tbPage1.txtVB.Value
'    rowsearch.Offset(0, 13).Value = Me.MultiPage1.tbPage1.txtSF.Value
'
'    'Page 2
'    rowsearch.Offset(0, 14).Value = Me.MultiPage1.tbPage2.txtPrjMetric.Value
'    rowsearch.Offset(0, 15).Value = Me.MultiPage1.tbPage2.txtPrjBudget.Value
'    rowsearch.Offset(0, 16).Value = Me.MultiPage1.tbPage2.txtBudgetCC.Value
'    rowsearch.Offset(0, 17).Value = Me.MultiPage1.tbPage2.txtDeliverables.Value
'
'    'Convert string to date on writing
'    If Me.MultiPage1.tbPage2.txtStartDate.Value <> "" Then
'        rowsearch.Offset(0, 18).Value = DateValue(Me.MultiPage1.tbPage2.txtStartDate.Value)
'    Else
'        rowsearch.Offset(0, 18).Value = Me.MultiPage1.tbPage2.txtStartDate.Value
'    End If
'
'    If Me.MultiPage1.tbPage2.txtEndDate.Value <> "" Then
'        rowsearch.Offset(0, 19).Value = DateValue(Me.MultiPage1.tbPage2.txtEndDate.Value)
'    Else
'        rowsearch.Offset(0, 19).Value = Me.MultiPage1.tbPage2.txtEndDate.Value
'    End If
'
'
'    rowsearch.Offset(0, 20).Value = Me.MultiPage1.tbPage2.txtKN.Value
'    rowsearch.Offset(0, 21).Value = Me.MultiPage1.tbPage2.txtLink.Value
'
'    'Page 3
'    rowsearch.Offset(0, 22).Value = Me.MultiPage1.tbPage3.cboPrjChampTeam.Value
'    rowsearch.Offset(0, 23).Value = Me.MultiPage1.tbPage3.cboPrjChamp.Value
'    rowsearch.Offset(0, 25).Value = Me.MultiPage1.tbPage3.txtOtherTeam.Value
'
'    'concatenate list box selection into string
'    'source: https://www.contextures.com/excel-data-validation-listbox.html
'    With Me.MultiPage1.tbPage3.lstTeamSupp
'       For lCountList = 0 To .ListCount - 1
'
'          If .Selected(lCountList) Then
'             strAdd = .List(lCountList)
'          Else
'             strAdd = ""
'          End If
'
'          If strSelItems = "" Then
'             strSelItems = strAdd
'          Else
'             If strAdd <> "" Then
'                strSelItems = strSelItems & strSep & strAdd
'             End If
'          End If
'
'       Next lCountList
'    End With
'
'    rowsearch.Offset(0, 24).Value = strSelItems
'
'    'Page 4
'    rowsearch.Offset(0, 3).Value = Me.MultiPage1.tbPage4.cboPrjStatus.Value
'    If rowsearch.Offset(0, 3).Value = "" And Me.cbMET.Value = False Then
'         rowsearch.Offset(0, 3).Value = "Yet to be started"
'    End If
'    rowsearch.Offset(0, 4).Value = Me.MultiPage1.tbPage4.txtComments.Value
'
'    'Change due status
'    If Me.MultiPage1.tbPage2.txtEndDate.Value <> "" Then
'        If CDate(Me.MultiPage1.tbPage2.txtEndDate.Value) < Date And Me.lblPrjStatusCapt.Caption <> "Monitor" Then
'            rowsearch.Offset(0, 26).Value = "Yes"
'        Else
'            rowsearch.Offset(0, 26).Value = ""
'        End If
'    End If
'
'End Sub
'
'Sub Read_from_sheet(rowsearch As Range)
'    'Read values to spreadsheet register tab
'
'    'Main Page
'    Me.txtPrjID.Value = rowsearch
'    Me.txtPrjTitle.Value = rowsearch.Offset(0, 1)
'    Me.lblPrjStatusCapt.Caption = rowsearch.Offset(0, 3)
'    Me.cbMET.Value = rowsearch.Offset(0, 2)
'
'    'Page 1
'    Me.MultiPage1.tbPage1.cboPrjOrigTeam.Value = rowsearch.Offset(0, 5).Value
'    Me.MultiPage1.tbPage1.txtPrjOrigPerson.Value = rowsearch.Offset(0, 6).Value
'    Me.MultiPage1.tbPage1.txtPrjContext.Value = rowsearch.Offset(0, 7).Value
'    Me.MultiPage1.tbPage1.txtPrjGoal.Value = rowsearch.Offset(0, 8).Value
'    Me.MultiPage1.tbPage1.cboImprovType.Value = rowsearch.Offset(0, 9).Value
'    Me.MultiPage1.tbPage1.txtAB.Value = rowsearch.Offset(0, 10).Value
'    Me.MultiPage1.tbPage1.cbVerified.Value = rowsearch.Offset(0, 11).Value
'    Me.MultiPage1.tbPage1.txtVB.Value = rowsearch.Offset(0, 12).Value
'    Me.MultiPage1.tbPage1.txtSF.Value = rowsearch.Offset(0, 13).Value
'
'    'Page 2
'    Me.MultiPage1.tbPage2.txtPrjMetric.Value = rowsearch.Offset(0, 14).Value
'    Me.MultiPage1.tbPage2.txtPrjBudget.Value = rowsearch.Offset(0, 15).Value
'    Me.MultiPage1.tbPage2.txtBudgetCC.Value = rowsearch.Offset(0, 16).Value
'    Me.MultiPage1.tbPage2.txtDeliverables.Value = rowsearch.Offset(0, 17).Value
'
'    'Change date format shown
'    If rowsearch.Offset(0, 18).Value <> "" Then
'        Me.MultiPage1.tbPage2.txtStartDate.Value = Format(rowsearch.Offset(0, 18).Value, "dd/mm/yyyy")
'    Else
'        Me.MultiPage1.tbPage2.txtStartDate.Value = rowsearch.Offset(0, 18).Value
'    End If
'
'    If rowsearch.Offset(0, 19).Value <> "" Then
'        Me.MultiPage1.tbPage2.txtEndDate.Value = Format(rowsearch.Offset(0, 19).Value, "dd/mm/yyyy")
'    Else
'        Me.MultiPage1.tbPage2.txtEndDate.Value = rowsearch.Offset(0, 19).Value
'    End If
'
'    Me.MultiPage1.tbPage2.txtKN.Value = rowsearch.Offset(0, 20).Value
'    Me.MultiPage1.tbPage2.txtLink.Value = rowsearch.Offset(0, 21).Value
'
'    'Page 3
'    Me.MultiPage1.tbPage3.cboPrjChampTeam.Value = rowsearch.Offset(0, 22).Value
'    Me.MultiPage1.tbPage3.cboPrjChamp.Value = rowsearch.Offset(0, 23).Value
'    Me.MultiPage1.tbPage3.txtOtherTeam.Value = rowsearch.Offset(0, 25).Value
'
'    'Fill list box
'    Call ReadStringToList(Me.MultiPage1.tbPage3.lstTeamSupp, rowsearch.Offset(0, 24).Value)
'
'    'Page 4
'    Me.MultiPage1.tbPage4.cboPrjStatus.Value = rowsearch.Offset(0, 3)
'    Me.MultiPage1.tbPage4.txtComments.Value = rowsearch.Offset(0, 4)
'
'    'Change status to red tinge if over due
'    Me.lblPrjStatusCapt.BackColor = &HFFFFFF
'    Me.txtPrjTitle.BackColor = &HFFFFFF
'
'    Sheets("Register").Unprotect
'
'    If Me.MultiPage1.tbPage2.txtEndDate.Value <> "" Then
'        If CDate(Me.MultiPage1.tbPage2.txtEndDate.Value) < Date And Me.lblPrjStatusCapt.Caption <> "Monitor" Then
'            Me.lblPrjStatusCapt.BackColor = &HC0C0FF
'            Me.txtPrjTitle.BackColor = &HC0C0FF
'            rowsearch.Offset(0, 26).Value = "Yes"
'        Else
'            rowsearch.Offset(0, 26).Value = ""
'        End If
'    End If
'
'    Sheets("Register").Protect
'
'End Sub
