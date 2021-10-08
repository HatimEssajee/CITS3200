VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form14_ReminderLog 
   Caption         =   "Reminder Log"
   ClientHeight    =   10350
   ClientLeft      =   -390
   ClientTop       =   -1815
   ClientWidth     =   25920
   OleObjectBlob   =   "form14_ReminderLog.frx":0000
End
Attribute VB_Name = "form14_ReminderLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub UserForm_Activate()
    'PURPOSE: Reposition userform to Top Left of application Window and fix size
    'source: https://www.mrexcel.com/board/threads/userform-startup-position.671108/
    Me.Top = UserFormTopPosR
    Me.Left = UserFormLeftPosR
    Me.Height = 550 '435 'UHeight
    Me.Width = 980 '940 'UWidth
    
    Call UserForm_Initialize
End Sub

Private Sub UserForm_Deactivate()
    'Store form position
    UserFormTopPosR = Me.Top
    UserFormLeftPosR = Me.Left
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
    
    
    'Fill text fields
    
    Me.lblRLStudyName = ReadRow(1, 9)
    
    Me.remStudyDetails.Value = ReadRow(1, 13)

    Me.remCDA.Value = ReadRow(1, 21)
    
    Me.remFS.Value = ReadRow(1, 27)
    
    Me.remSiteSelect.Value = ReadRow(1, 35)

    Me.remRecruitment.Value = ReadRow(1, 39)

    Me.remCAHS_Ethics.Value = ReadRow(1, 46)
    Me.remNMA_Ethics.Value = ReadRow(1, 50)
    Me.remWNHS_Ethics.Value = ReadRow(1, 53)
    Me.remSJOG_Ethics.Value = ReadRow(1, 56)
    Me.remOthers_Ethics.Value = ReadRow(1, 60)
    
    Me.remPCH_Gov.Value = ReadRow(1, 66)
    Me.remTKI_Gov.Value = ReadRow(1, 70)
    Me.remKEMH_Gov.Value = ReadRow(1, 74)
    Me.remSJOG_S_Gov.Value = ReadRow(1, 78)
    Me.remSJOG_L_Gov.Value = ReadRow(1, 82)
    Me.remSJOG_M_Gov.Value = ReadRow(1, 86)
    Me.remOthers_Gov.Value = ReadRow(1, 91)

    Me.remVTG_Budget.Value = ReadRow(1, 97)
    Me.remTKI_Budget.Value = ReadRow(1, 99)
    Me.remPharm_Budget.Value = ReadRow(1, 102)

    Me.remIndemnity.Value = ReadRow(1, 108)

    Me.remCTRA.Value = ReadRow(1, 118)

    Me.remFinDisc.Value = ReadRow(1, 122)

    Me.remSIV.Value = ReadRow(1, 126)


    'Assess stage status
    
    'Study Details
    If ReadRow(1, 129) Then
        Me.remStudyDetails.BackColor = &H80FF80
    Else
        Me.remStudyDetails.BackColor = &H80000005
    End If
    
    'CDA
    If ReadRow(1, 130) Then
        Me.remCDA.BackColor = &H80FF80
    Else
        Me.remCDA.BackColor = &H80000005
    End If
    
    'Feasibility
    If ReadRow(1, 131) Then
        Me.remFS.BackColor = &H80FF80
    Else
        Me.remFS.BackColor = &H80000005
    End If
    
    'Site Selection
    If ReadRow(1, 132) Then
        Me.remSiteSelect.BackColor = &H80FF80
    Else
        Me.remSiteSelect.BackColor = &H80000005
    End If
    
    'Recruitment
    If ReadRow(1, 133) Then
        Me.remRecruitment.BackColor = &H80FF80
    Else
        Me.remRecruitment.BackColor = &H80000005
    End If
    
    'CAHS Ethics
    If ReadRow(1, 134) Then
        Me.remCAHS_Ethics.BackColor = &H80FF80
    Else
        Me.remCAHS_Ethics.BackColor = &H80000005
    End If
    
    'NMA Ethics
    If ReadRow(1, 135) Then
        Me.remNMA_Ethics.BackColor = &H80FF80
    Else
        Me.remNMA_Ethics.BackColor = &H80000005
    End If
    
    'WNHS Ethics
    If ReadRow(1, 136) Then
        Me.remWNHS_Ethics.BackColor = &H80FF80
    Else
        Me.remWNHS_Ethics.BackColor = &H80000005
    End If
    
    'SJOG Ethics
    If ReadRow(1, 137) Then
        Me.remSJOG_Ethics.BackColor = &H80FF80
    Else
        Me.remSJOG_Ethics.BackColor = &H80000005
    End If
    
    'Others Ethics
    If ReadRow(1, 138) Then
        Me.remOthers_Ethics.BackColor = &H80FF80
    Else
        Me.remOthers_Ethics.BackColor = &H80000005
    End If
    
    'PCH Governance
    If ReadRow(1, 139) Then
        Me.remPCH_Gov.BackColor = &H80FF80
    Else
        Me.remPCH_Gov.BackColor = &H80000005
    End If
    
    'TKI Governance
    If ReadRow(1, 140) Then
        Me.remTKI_Gov.BackColor = &H80FF80
    Else
        Me.remTKI_Gov.BackColor = &H80000005
    End If
    
    'KEMH Governance
    If ReadRow(1, 141) Then
        Me.remKEMH_Gov.BackColor = &H80FF80
    Else
        Me.remKEMH_Gov.BackColor = &H80000005
    End If
    
    'SJOG_S Governance
    If ReadRow(1, 142) Then
        Me.remSJOG_S_Gov.BackColor = &H80FF80
    Else
        Me.remSJOG_S_Gov.BackColor = &H80000005
    End If
    
    'SJOG_L Governance
    If ReadRow(1, 143) Then
        Me.remSJOG_L_Gov.BackColor = &H80FF80
    Else
        Me.remSJOG_L_Gov.BackColor = &H80000005
    End If
    
    'SJOG_M Governance
    If ReadRow(1, 144) Then
        Me.remSJOG_M_Gov.BackColor = &H80FF80
    Else
        Me.remSJOG_M_Gov.BackColor = &H80000005
    End If
    
    'Others Governance
    If ReadRow(1, 145) Then
        Me.remOthers_Gov.BackColor = &H80FF80
    Else
        Me.remOthers_Gov.BackColor = &H80000005
    End If
    
    'VTG Budget
    If ReadRow(1, 146) Then
        Me.remVTG_Budget.BackColor = &H80FF80
    Else
        Me.remVTG_Budget.BackColor = &H80000005
    End If
    
    'TKI Budget
    If ReadRow(1, 147) Then
        Me.remTKI_Budget.BackColor = &H80FF80
    Else
        Me.remTKI_Budget.BackColor = &H80000005
    End If
    
    'Pharmacy Budget
    If ReadRow(1, 148) Then
        Me.remPharm_Budget.BackColor = &H80FF80
    Else
        Me.remPharm_Budget.BackColor = &H80000005
    End If
    
    'Indemnity
    If ReadRow(1, 149) Then
        Me.remIndemnity.BackColor = &H80FF80
    Else
        Me.remIndemnity.BackColor = &H80000005
    End If
    
    'CTRA
    If ReadRow(1, 150) Then
        Me.remCTRA.BackColor = &H80FF80
    Else
        Me.remCTRA.BackColor = &H80000005
    End If

    'Financial Disclosure
    If ReadRow(1, 151) Then
        Me.remFinDisc.BackColor = &H80FF80
    Else
        Me.remFinDisc.BackColor = &H80000005
    End If

    'Site Initiation Visit
    If ReadRow(1, 152) Then
        Me.remSIV.BackColor = &H80FF80
    Else
        Me.remSIV.BackColor = &H80000005
    End If
    
    'Reinstate Settings
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayStatusBar = True
    Application.EnableEvents = True
    
End Sub

Private Sub cmdCloseLog_Click()
    'PURPOSE: Closes current form
    Unload form14_ReminderLog
    
End Sub


