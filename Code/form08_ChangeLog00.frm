VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form08_ChangeLog00 
   Caption         =   "Change Log"
   ClientHeight    =   10320
   ClientLeft      =   -360
   ClientTop       =   -1560
   ClientWidth     =   9228.001
   OleObjectBlob   =   "form08_ChangeLog00.frx":0000
End
Attribute VB_Name = "form08_ChangeLog00"
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
    With RegTable.ListRows(RowIndex)
    
        Me.timeCreatedOn.Value = Format(.Range(2).Value, "dd-mmm-yyyy hh:mm:ss AM/PM")
        Me.perCreatedOn.Value = .Range(3).Value
        Me.timeDeletedOn.Value = Format(.Range(4).Value, "dd-mmm-yyyy hh:mm:ss AM/PM")
        Me.perDeletedOn.Value = .Range(5).Value
        Me.timeLastAccessed.Value = Format(.Range(6).Value, "dd-mmm-yyyy hh:mm:ss AM/PM")
        Me.perLastAccessed.Value = .Range(7).Value
        
        Me.timeStudyDetails.Value = Format(.Range(14).Value, "dd-mmm-yyyy hh:mm:ss AM/PM")
        Me.perStudyDetails.Value = .Range(15).Value
        
        Me.timeCDA_FS.Value = Format(.Range(24).Value, "dd-mmm-yyyy hh:mm:ss AM/PM")
        Me.perCDA_FS.Value = .Range(25).Value
        
        Me.timeSiteSelect.Value = Format(.Range(32).Value, "dd-mmm-yyyy hh:mm:ss AM/PM")
        Me.perSiteSelect.Value = .Range(33).Value
    
        Me.timeRecruitment.Value = Format(.Range(37).Value, "dd-mmm-yyyy hh:mm:ss AM/PM")
        Me.perRecruitment.Value = .Range(38).Value
        
        Me.timeEthics.Value = Format(.Range(54).Value, "dd-mmm-yyyy hh:mm:ss AM/PM")
        Me.perEthics.Value = .Range(55).Value
        
        Me.timeGovernance.Value = Format(.Range(79).Value, "dd-mmm-yyyy hh:mm:ss AM/PM")
        Me.perGovernance.Value = .Range(80).Value
              
        Me.timeBudget.Value = Format(.Range(88).Value, "dd-mmm-yyyy hh:mm:ss AM/PM")
        Me.perBudget.Value = .Range(89).Value
        
        Me.timeIndemnity.Value = Format(.Range(94).Value, "dd-mmm-yyyy hh:mm:ss AM/PM")
        Me.perIndemnity.Value = .Range(95).Value
        
        Me.timeCTRA.Value = Format(.Range(104).Value, "dd-mmm-yyyy hh:mm:ss AM/PM")
        Me.perCTRA.Value = .Range(105).Value
        
        Me.timeFinDisc.Value = Format(.Range(108).Value, "dd-mmm-yyyy hh:mm:ss AM/PM")
        Me.perFinDisc.Value = .Range(109).Value
        
        Me.timeSIV.Value = Format(.Range(112).Value, "dd-mmm-yyyy hh:mm:ss AM/PM")
        Me.perSIV.Value = .Range(113).Value
        
    End With
    
    'Load data into Change Log
End Sub

Private Sub cmdCloseLog_Click()
    'PURPOSE: Closes current form
    Unload form08_ChangeLog
    
End Sub


