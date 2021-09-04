VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form08_ChangeLog 
   Caption         =   "Change Log"
   ClientHeight    =   10320
   ClientLeft      =   -360
   ClientTop       =   -1560
   ClientWidth     =   9228.001
   OleObjectBlob   =   "form08_ChangeLog.frx":0000
End
Attribute VB_Name = "form08_ChangeLog"
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
    
    Me.timeCreatedOn.value = Format(ReadRow(RowIndex, 2), "dd-mmm-yyyy hh:mm:ss AM/PM")
    Me.perCreatedOn.value = ReadRow(RowIndex, 3)
    Me.timeDeletedOn.value = Format(ReadRow(RowIndex, 4), "dd-mmm-yyyy hh:mm:ss AM/PM")
    Me.perDeletedOn.value = ReadRow(RowIndex, 5)
    Me.timeLastAccessed.value = Format(ReadRow(RowIndex, 6), "dd-mmm-yyyy hh:mm:ss AM/PM")
    Me.perLastAccessed.value = ReadRow(RowIndex, 7)
        
    Me.timeStudyDetails.value = Format(ReadRow(RowIndex, 15), "dd-mmm-yyyy hh:mm:ss AM/PM")
    Me.perStudyDetails.value = ReadRow(RowIndex, 16)
        
    Me.timeCDA_FS.value = Format(ReadRow(RowIndex, 26), "dd-mmm-yyyy hh:mm:ss AM/PM")
    Me.perCDA_FS.value = ReadRow(RowIndex, 27)
        
    Me.timeSiteSelect.value = Format(ReadRow(RowIndex, 34), "dd-mmm-yyyy hh:mm:ss AM/PM")
    Me.perSiteSelect.value = ReadRow(RowIndex, 35)
    
    Me.timeRecruitment.value = Format(ReadRow(RowIndex, 39), "dd-mmm-yyyy hh:mm:ss AM/PM")
    Me.perRecruitment.value = ReadRow(RowIndex, 40)
    
    Me.timeEthics.value = Format(ReadRow(RowIndex, 56), "dd-mmm-yyyy hh:mm:ss AM/PM")
    Me.perEthics.value = ReadRow(RowIndex, 57)
    
    Me.timeGovernance.value = Format(ReadRow(RowIndex, 81), "dd-mmm-yyyy hh:mm:ss AM/PM")
    Me.perGovernance.value = ReadRow(RowIndex, 82)
              
    Me.timeBudget.value = Format(ReadRow(RowIndex, 90), "dd-mmm-yyyy hh:mm:ss AM/PM")
    Me.perBudget.value = ReadRow(RowIndex, 91)
    
    Me.timeIndemnity.value = Format(ReadRow(RowIndex, 96), "dd-mmm-yyyy hh:mm:ss AM/PM")
    Me.perIndemnity.value = ReadRow(RowIndex, 97)
    
    Me.timeCTRA.value = Format(ReadRow(RowIndex, 106), "dd-mmm-yyyy hh:mm:ss AM/PM")
    Me.perCTRA.value = ReadRow(RowIndex, 107)
    
    Me.timeFinDisc.value = Format(ReadRow(RowIndex, 110), "dd-mmm-yyyy hh:mm:ss AM/PM")
    Me.perFinDisc.value = ReadRow(RowIndex, 111)
    
    Me.timeSIV.value = Format(ReadRow(RowIndex, 114), "dd-mmm-yyyy hh:mm:ss AM/PM")
    Me.perSIV.value = ReadRow(RowIndex, 115)
    
End Sub

Private Sub cmdCloseLog_Click()
    'PURPOSE: Closes current form
    Unload form08_ChangeLog
    
End Sub


