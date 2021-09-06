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
    
    Me.timeCreatedOn.value = Format(ReadRow(1, 2), "dd-mmm-yyyy hh:mm:ss AM/PM")
    Me.perCreatedOn.value = ReadRow(1, 3)
    Me.timeDeletedOn.value = Format(ReadRow(1, 4), "dd-mmm-yyyy hh:mm:ss AM/PM")
    Me.perDeletedOn.value = ReadRow(1, 5)
    Me.timeLastAccessed.value = Format(ReadRow(1, 6), "dd-mmm-yyyy hh:mm:ss AM/PM")
    Me.perLastAccessed.value = ReadRow(1, 7)
        
    Me.timeStudyDetails.value = Format(ReadRow(1, 15), "dd-mmm-yyyy hh:mm:ss AM/PM")
    Me.perStudyDetails.value = ReadRow(1, 16)
        
    Me.timeCDA_FS.value = Format(ReadRow(1, 26), "dd-mmm-yyyy hh:mm:ss AM/PM")
    Me.perCDA_FS.value = ReadRow(1, 27)
        
    Me.timeSiteSelect.value = Format(ReadRow(1, 34), "dd-mmm-yyyy hh:mm:ss AM/PM")
    Me.perSiteSelect.value = ReadRow(1, 35)
    
    Me.timeRecruitment.value = Format(ReadRow(1, 39), "dd-mmm-yyyy hh:mm:ss AM/PM")
    Me.perRecruitment.value = ReadRow(1, 40)
    
    Me.timeEthics.value = Format(ReadRow(1, 56), "dd-mmm-yyyy hh:mm:ss AM/PM")
    Me.perEthics.value = ReadRow(1, 57)
    
    Me.timeGovernance.value = Format(ReadRow(1, 81), "dd-mmm-yyyy hh:mm:ss AM/PM")
    Me.perGovernance.value = ReadRow(1, 82)
              
    Me.timeBudget.value = Format(ReadRow(1, 90), "dd-mmm-yyyy hh:mm:ss AM/PM")
    Me.perBudget.value = ReadRow(1, 91)
    
    Me.timeIndemnity.value = Format(ReadRow(1, 96), "dd-mmm-yyyy hh:mm:ss AM/PM")
    Me.perIndemnity.value = ReadRow(1, 97)
    
    Me.timeCTRA.value = Format(ReadRow(1, 106), "dd-mmm-yyyy hh:mm:ss AM/PM")
    Me.perCTRA.value = ReadRow(1, 107)
    
    Me.timeFinDisc.value = Format(ReadRow(1, 110), "dd-mmm-yyyy hh:mm:ss AM/PM")
    Me.perFinDisc.value = ReadRow(1, 111)
    
    Me.timeSIV.value = Format(ReadRow(1, 114), "dd-mmm-yyyy hh:mm:ss AM/PM")
    Me.perSIV.value = ReadRow(1, 115)
    
End Sub

Private Sub cmdCloseLog_Click()
    'PURPOSE: Closes current form
    Unload form08_ChangeLog
    
End Sub


