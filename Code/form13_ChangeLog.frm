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
    Me.Height = 520 'UHeight
    Me.Width = 500 'UWidth
    
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
    
    'Turn off Settings to speed up
    'source: https://www.automateexcel.com/vba/turn-off-screen-updating/
    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayStatusBar = False
    Application.EnableEvents = False
    
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
    
    Me.timeCreatedOn.Value = Format(ReadRow(1, 1), "dd-mmm-yyyy hh:mm:ss AM/PM")
    Me.perCreatedOn.Value = ReadRow(1, 2)
    Me.timeDeletedOn.Value = Format(ReadRow(1, 3), "dd-mmm-yyyy hh:mm:ss AM/PM")
    Me.perDeletedOn.Value = ReadRow(1, 4)
    Me.timeLastAccessed.Value = Format(ReadRow(1, 5), "dd-mmm-yyyy hh:mm:ss AM/PM")
    Me.perLastAccessed.Value = ReadRow(1, 6)
        
    Me.timeStudyDetails.Value = Format(ReadRow(1, 14), "dd-mmm-yyyy hh:mm:ss AM/PM")
    Me.perStudyDetails.Value = ReadRow(1, 15)
        
    Me.timeCDA.Value = Format(ReadRow(1, 22), "dd-mmm-yyyy hh:mm:ss AM/PM")
    Me.perCDA.Value = ReadRow(1, 23)
     
    Me.timeFS.Value = Format(ReadRow(1, 28), "dd-mmm-yyyy hh:mm:ss AM/PM")
    Me.perFS.Value = ReadRow(1, 29)
    
    Me.timeSiteSelect.Value = Format(ReadRow(1, 36), "dd-mmm-yyyy hh:mm:ss AM/PM")
    Me.perSiteSelect.Value = ReadRow(1, 37)
    
    Me.timeRecruitment.Value = Format(ReadRow(1, 40), "dd-mmm-yyyy hh:mm:ss AM/PM")
    Me.perRecruitment.Value = ReadRow(1, 41)
    
    Me.timeEthics.Value = Format(ReadRow(1, 61), "dd-mmm-yyyy hh:mm:ss AM/PM")
    Me.perEthics.Value = ReadRow(1, 62)
    
    Me.timeGovernance.Value = Format(ReadRow(1, 92), "dd-mmm-yyyy hh:mm:ss AM/PM")
    Me.perGovernance.Value = ReadRow(1, 93)
              
    Me.timeBudget.Value = Format(ReadRow(1, 103), "dd-mmm-yyyy hh:mm:ss AM/PM")
    Me.perBudget.Value = ReadRow(1, 104)
    
    Me.timeIndemnity.Value = Format(ReadRow(1, 109), "dd-mmm-yyyy hh:mm:ss AM/PM")
    Me.perIndemnity.Value = ReadRow(1, 110)
    
    Me.timeCTRA.Value = Format(ReadRow(1, 119), "dd-mmm-yyyy hh:mm:ss AM/PM")
    Me.perCTRA.Value = ReadRow(1, 120)
    
    Me.timeFinDisc.Value = Format(ReadRow(1, 123), "dd-mmm-yyyy hh:mm:ss AM/PM")
    Me.perFinDisc.Value = ReadRow(1, 124)
    
    Me.timeSIV.Value = Format(ReadRow(1, 127), "dd-mmm-yyyy hh:mm:ss AM/PM")
    Me.perSIV.Value = ReadRow(1, 128)
    
    'Reinstate Settings
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayStatusBar = True
    Application.EnableEvents = True
    
End Sub

Private Sub cmdCloseLog_Click()
    'PURPOSE: Closes current form
    Unload form13_ChangeLog
    
End Sub
