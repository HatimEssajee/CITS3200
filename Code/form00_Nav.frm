VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form00_Nav 
   Caption         =   "Vaccine Trial Study Start-up Tracker"
   ClientHeight    =   5784
   ClientLeft      =   36
   ClientTop       =   -72
   ClientWidth     =   6444
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
    
    cboStudyStatus.TextAlign = fmTextAlignCenter
        
End Sub


Private Sub cmdClose_Click()
    'PURPOSE: Closes current form
    Unload Me
        
End Sub

Private Sub cmdNew_Click()
    'PURPOSE: Closes current form and open Study Detail form
    
    Dim FoundCell As Range
    
    On Error GoTo ErrHandler:
    
    Call TurnEvents_OFF
    
    'Set Public Variable
    StudyName = Me.txtStudyName.Value
    
    'Check if study name already in table
    'Source: https://www.thespreadsheetguru.com/blog/2014/6/20/the-vba-guide-to-listobject-excel-tables
    On Error Resume Next
    Set FoundCell = RegTable.DataBodyRange.Columns(9).Find(StudyName, LookAt:=xlWhole)
    On Error GoTo 0
    
    If Not FoundCell Is Nothing Then
        RowIndex = RegTable.ListRows(FoundCell.Row - RegTable.HeaderRowRange.Row).Index
        Exit Sub
    End If
    
    
    'Add Row to register table and repoint row references
    'Source: https://www.bluepecantraining.com/portfolio/excel-vba-how-to-add-rows-and-columns-to-excel-table-with-vba-macro/
    Set ReadRow = RegTable.ListRows.Add
    
    RowIndex = RegTable.ListRows.Count
    
    With ReadRow
        .Range(1) = Now
        .Range(2) = Username
        .Range(7) = Me.cboStudyStatus.Value
        .Range(8) = Me.txtProtocolNum
        .Range(9) = StudyName
        .Range(13) = .Range(1).Value
        .Range(14) = .Range(2).Value
    End With
    
    Unload form00_Nav
    
    form01_StudyDetail.Show False
    
ErrHandler:
     Call TurnEvents_ON
     
End Sub

Private Sub cboStudyStatus_Change()
    
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
            .Range(3) = Now
            .Range(4) = Username
            .Range(7) = "DELETED"
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
    
End Sub


Private Sub cmdNext_Click()

    Dim Jump As Integer
    Dim BtmRow As Long
    
    'Set Toggle interval and variables
    Jump = 1
    
    BtmRow = RegTable.ListRows.Count
    
    'Increment row index
    If RowIndex < 0 Then
        RowIndex = BtmRow
    ElseIf RowIndex < BtmRow - Jump Then
        RowIndex = RowIndex + Jump
    Else
        RowIndex = BtmRow
    End If
    
    Call Read_Table

End Sub

    
Private Sub cmdPrevious_Click()

    Dim Jump As Integer

    'Set Toggle interval and variables
    Jump = 1
    
    'Increment row index
    If RowIndex < 0 Then
        RowIndex = 1
    ElseIf RowIndex > 2 - Jump Then
        RowIndex = RowIndex - Jump
    Else
        RowIndex = 1
    End If
    
    Call Read_Table

End Sub


Private Sub Read_Table()
    
    With RegTable.ListRows(RowIndex)
        Me.txtStudyName.Value = .Range(9).Value
        Me.txtProtocolNum.Value = .Range(8).Value
        Me.cboStudyStatus.Value = .Range(7).Value
        Me.cboStudyStatus.ForeColor = StudyStatus_Colour(.Range(7).Value)
    End With
    
End Sub

Private Function StudyStatus_Colour(status As String) As Long
       
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
