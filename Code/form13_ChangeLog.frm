VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form13_ChangeLog 
   Caption         =   "Change Log"
   ClientHeight    =   10320
   ClientLeft      =   -360
   ClientTop       =   -1545
   ClientWidth     =   9225.001
   OleObjectBlob   =   "form13_ChangeLog.frx":0000
End
Attribute VB_Name = "form13_ChangeLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub UserForm_Activate()
    'PURPOSE: Reposition userform to Top Left of application Window and fix size
    'source: https://www.mrexcel.com/board/threads/userform-startup-position.671108/
    Me.Top = UserFormTopPosC
    Me.Left = UserFormLeftPosC
    Me.Height = UHeight
    Me.Width = UWidth
    
    Call UserForm_Initialize
    
End Sub

Private Sub UserForm_Deactivate()
    'Store form position
    UserFormTopPosC = Me.Top
    UserFormLeftPosC = Me.Left
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
                    ctrl.Value = False
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
    
    'Pull in data from register table
    ReadRow = RegTable.DataBodyRange.Rows(RowIndex)
    
    Me.timeCreatedOn.Value = Format(ReadRow(1, 2), "dd-mmm-yyyy hh:mm:ss AM/PM")
    Me.perCreatedOn.Value = ReadRow(1, 3)
    Me.timeDeletedOn.Value = Format(ReadRow(1, 4), "dd-mmm-yyyy hh:mm:ss AM/PM")
    Me.perDeletedOn.Value = ReadRow(1, 5)
    Me.timeLastAccessed.Value = Format(ReadRow(1, 6), "dd-mmm-yyyy hh:mm:ss AM/PM")
    Me.perLastAccessed.Value = ReadRow(1, 7)
        
    Me.timeStudyDetails.Value = Format(ReadRow(1, 15), "dd-mmm-yyyy hh:mm:ss AM/PM")
    Me.perStudyDetails.Value = ReadRow(1, 16)
        
    Me.timeCDA_FS.Value = Format(ReadRow(1, 26), "dd-mmm-yyyy hh:mm:ss AM/PM")
    Me.perCDA_FS.Value = ReadRow(1, 27)
        
    Me.timeSiteSelect.Value = Format(ReadRow(1, 34), "dd-mmm-yyyy hh:mm:ss AM/PM")
    Me.perSiteSelect.Value = ReadRow(1, 35)
    
    Me.timeRecruitment.Value = Format(ReadRow(1, 39), "dd-mmm-yyyy hh:mm:ss AM/PM")
    Me.perRecruitment.Value = ReadRow(1, 40)
    
    Me.timeEthics.Value = Format(ReadRow(1, 56), "dd-mmm-yyyy hh:mm:ss AM/PM")
    Me.perEthics.Value = ReadRow(1, 57)
    
    Me.timeGovernance.Value = Format(ReadRow(1, 81), "dd-mmm-yyyy hh:mm:ss AM/PM")
    Me.perGovernance.Value = ReadRow(1, 82)
              
    Me.timeBudget.Value = Format(ReadRow(1, 90), "dd-mmm-yyyy hh:mm:ss AM/PM")
    Me.perBudget.Value = ReadRow(1, 91)
    
    Me.timeIndemnity.Value = Format(ReadRow(1, 96), "dd-mmm-yyyy hh:mm:ss AM/PM")
    Me.perIndemnity.Value = ReadRow(1, 97)
    
    Me.timeCTRA.Value = Format(ReadRow(1, 106), "dd-mmm-yyyy hh:mm:ss AM/PM")
    Me.perCTRA.Value = ReadRow(1, 107)
    
    Me.timeFinDisc.Value = Format(ReadRow(1, 110), "dd-mmm-yyyy hh:mm:ss AM/PM")
    Me.perFinDisc.Value = ReadRow(1, 111)
    
    Me.timeSIV.Value = Format(ReadRow(1, 114), "dd-mmm-yyyy hh:mm:ss AM/PM")
    Me.perSIV.Value = ReadRow(1, 115)
    
End Sub

Private Sub cmdCloseLog_Click()
    'PURPOSE: Closes current form
    Unload form08_ChangeLog
    
End Sub


