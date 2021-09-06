VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form09_ReminderLog 
   Caption         =   "Reminder Log"
   ClientHeight    =   10320
   ClientLeft      =   -360
   ClientTop       =   -1560
   ClientWidth     =   9228.001
   OleObjectBlob   =   "form09_ReminderLog.frx":0000
End
Attribute VB_Name = "form09_ReminderLog"
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


Private Sub UserForm_Initialize()
    'PURPOSE: Clear form on initialization
    'Source: https://www.contextures.com/xlUserForm02.html
    'Source: https://www.contextures.com/Excel-VBA-ComboBox-Lists.html
    Dim ctrl As MSForms.Control
    Dim ReadRow As Variant
    
    'Clear user form
    'source: https://www.mrexcel.com/board/threads/loop-through-controls-on-a-userform.427103/
    For Each ctrl In Me.Controls
        Select Case True
                Case TypeOf ctrl Is MSForms.CheckBox
                    ctrl.value = False
                Case TypeOf ctrl Is MSForms.TextBox
                    ctrl.value = ""
                Case TypeOf ctrl Is MSForms.ComboBox
                    ctrl.value = ""
                    ctrl.Clear
                Case TypeOf ctrl Is MSForms.ListBox
                    ctrl.value = ""
                    ctrl.Clear
            End Select
    Next ctrl
    
    'Pull in data from register table
    ReadRow = RegTable.DataBodyRange.Rows(RowIndex)
    
    
    'Fill text fields
    Me.remStudyDetails.value = ReadRow(1, 14)

    Me.remCDA_FS.value = ReadRow(1, 25)

    Me.remSiteSelect.value = ReadRow(1, 33)

    Me.remRecruitment.value = ReadRow(1, 38)

    Me.remEthics.value = ReadRow(1, 55)

    Me.remGovernance.value = ReadRow(1, 80)

    Me.remBudget.value = ReadRow(1, 89)

    Me.remIndemnity.value = ReadRow(1, 95)

    Me.remCTRA.value = ReadRow(1, 105)

    Me.remFinDisc.value = ReadRow(1, 109)

    Me.remSIV.value = ReadRow(1, 113)


    'Assess stage status

    'Has to have age range filled
    If ReadRow(1, 13) <> vbNullString Then
        Me.statStudyDetails.BackColor = vbGreen
    End If

    'Has to have CDA Finalised and Feasibility Study completed filled
    If ReadRow(1, 21) <> vbNullString And ReadRow(1, 23) <> vbNullString Then
        Me.statCDA_FS.BackColor = vbGreen
    End If

    'Has to have Site Selection Date filled
    If ReadRow(1, 32) <> vbNullString Then
        Me.statSiteSelect.BackColor = vbGreen
    End If

    'Has to have Recruitment status as complete
    If ReadRow(1, 37) = "Complete" Then
        Me.statRecruitment.BackColor = vbGreen
    End If

    'Has to have at least one committee approving ethics review
    'and all submitted reviews are approved with dates filled
    If (ReadRow(1, 44) <> vbNullString Or ReadRow(1, 47) <> vbNullString Or _
        ReadRow(1, 49) <> vbNullString Or ReadRow(1, 51) <> vbNullString Or _
        ReadRow(1, 54) <> vbNullString) And _
        ((ReadRow(1, 41) <> vbNullString And ReadRow(1, 44) <> vbNullString) Or _
        (ReadRow(1, 46) <> vbNullString And ReadRow(1, 47) <> vbNullString) Or _
        (ReadRow(1, 48) <> vbNullString And ReadRow(1, 49) <> vbNullString) Or _
        (ReadRow(1, 50) <> vbNullString And ReadRow(1, 51) <> vbNullString) Or _
        (ReadRow(1, 53) <> vbNullString And ReadRow(1, 54) <> vbNullString)) _
        Then
        Me.statEthics.BackColor = vbGreen
    End If

    'Has to have at least one committee approving governance review
    'and all submitted reviews are approved with dates filled
    If (ReadRow(1, 60) <> vbNullString Or ReadRow(1, 63) <> vbNullString Or _
        ReadRow(1, 66) <> vbNullString Or ReadRow(1, 69) <> vbNullString Or _
        ReadRow(1, 72) <> vbNullString Or ReadRow(1, 75) <> vbNullString Or _
        ReadRow(1, 79) <> vbNullString) And _
        ((ReadRow(1, 58) <> vbNullString And ReadRow(1, 60) <> vbNullString) Or _
        (ReadRow(1, 61) <> vbNullString And ReadRow(1, 63) <> vbNullString) Or _
        (ReadRow(1, 64) <> vbNullString And ReadRow(1, 66) <> vbNullString) Or _
        (ReadRow(1, 67) <> vbNullString And ReadRow(1, 69) <> vbNullString) Or _
        (ReadRow(1, 70) <> vbNullString And ReadRow(1, 72) <> vbNullString) Or _
        (ReadRow(1, 73) <> vbNullString And ReadRow(1, 75) <> vbNullString) Or _
        (ReadRow(1, 77) <> vbNullString And ReadRow(1, 79) <> vbNullString)) _
        Then
        Me.statGovernance.BackColor = vbGreen
    End If

    'Has to have all parties approve Budget with dates filled
    If ReadRow(1, 85) <> vbNullString And ReadRow(1, 86) <> vbNullString And _
        ReadRow(1, 88) <> vbNullString Then
        Me.statBudget.BackColor = vbGreen
    End If

    'Has to have Date Completed filled
    If ReadRow(1, 94) <> vbNullString Then
        Me.statIndemnity.BackColor = vbGreen
    End If

    'Has to have Date Finalised filled
    If ReadRow(1, 104) <> vbNullString Then
        Me.statCTRA.BackColor = vbGreen
    End If

    'Has to have Date Completed filled
    If ReadRow(1, 108) <> vbNullString Then
        Me.statFinDisc.BackColor = vbGreen
    End If

    'Has to have Site Initiation Visit Date filled
    If ReadRow(1, 112) <> vbNullString Then
        Me.statSIV.BackColor = vbGreen
    End If

End Sub

Private Sub cmdCloseLog_Click()
    'PURPOSE: Closes current form
    Unload form09_ReminderLog
    
End Sub


