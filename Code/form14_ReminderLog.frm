VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form14_ReminderLog 
   Caption         =   "Reminder Log"
   ClientHeight    =   8268.001
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
    
    Call Userform_Initialize
End Sub

Private Sub UserForm_Deactivate()
    'Store form position
    UserFormTopPosR = Me.Top
    UserFormLeftPosR = Me.Left
End Sub

Private Sub Userform_Initialize()
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
    
    'Swap caption if Ethics committee known
    If ReadRow(1, 47) <> vbNullString Then
        Me.lblRLNMA_Ethics.Caption = ReadRow(1, 47)
    Else
        Me.lblRLNMA_Ethics.Caption = "NMA"
    End If
    
    
    Me.remNMA_Ethics.Value = ReadRow(1, 50)
    Me.remWNHS_Ethics.Value = ReadRow(1, 53)
    Me.remSJOG_Ethics.Value = ReadRow(1, 56)
    Me.remOthers_Ethics.Value = ReadRow(1, 60)
    
    'Swap caption if Ethics committee known
    If ReadRow(1, 57) <> vbNullString Then
        Me.lblRLOthers_Ethics.Caption = ReadRow(1, 57)
    Else
        Me.lblRLOthers_Ethics.Caption = "Other"
    End If
    
    Me.remPCH_Gov.Value = ReadRow(1, 66)
    Me.remTKI_Gov.Value = ReadRow(1, 70)
    Me.remKEMH_Gov.Value = ReadRow(1, 74)
    Me.remSJOG_S_Gov.Value = ReadRow(1, 78)
    Me.remSJOG_L_Gov.Value = ReadRow(1, 82)
    Me.remSJOG_M_Gov.Value = ReadRow(1, 86)
    
    'Swap caption if Ethics committee known
    If ReadRow(1, 87) <> vbNullString Then
        Me.lblRLOthers_Gov.Caption = ReadRow(1, 87)
    Else
        Me.lblRLOthers_Gov.Caption = "Other"
    End If
        
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
        Me.compStudyDetails.Visible = True
        Me.compStudyDetails.BackColor = &H80FF80
        Me.compStudyDetails.Caption = "Sponsor, CRO, Protocol & Age Range filled"
    Else
        Me.compStudyDetails.Visible = False
    End If
    
    'CDA
    If ReadRow(1, 130) Then
        Me.compCDA.Visible = True
        Me.compCDA.BackColor = &H80FF80
        Me.compCDA.Caption = "Date Finalised = " & Format(ReadRow(1, 20), "DD-MMM-YYYY")
    Else
        Me.compCDA.Visible = False
    End If
    
    'Feasibility
    If ReadRow(1, 131) Then
        Me.compFS.Visible = True
        Me.compFS.BackColor = &H80FF80
        Me.compFS.Caption = "Date Completed = " & Format(ReadRow(1, 25), "DD-MMM-YYYY")
    Else
        Me.compFS.Visible = False
    End If
    
    'Site Selection
    If ReadRow(1, 132) Then
        Me.compSiteSelect.Visible = True
        Me.compSiteSelect.BackColor = &H80FF80
        Me.compSiteSelect.Caption = "Site Selected = " & Format(ReadRow(1, 34), "DD-MMM-YYYY")
    Else
        Me.compSiteSelect.Visible = False
    End If
    
    'Recruitment
    If ReadRow(1, 133) Then
        Me.compRecruitment.Visible = True
        Me.compRecruitment.BackColor = &H80FF80
        Me.compRecruitment.Caption = "Planning Date = " & Format(ReadRow(1, 38), "DD-MMM-YYYY")
    Else
        Me.compRecruitment.Visible = False
    End If
    
    'CAHS Ethics
    If ReadRow(1, 134) Then
        Me.compCAHS_Ethics.Visible = True
        Me.compCAHS_Ethics.BackColor = &H80FF80
        Me.compCAHS_Ethics.Caption = "Date Approved = " & Format(ReadRow(1, 45), "DD-MMM-YYYY")
    Else
        Me.compCAHS_Ethics.Visible = False
    End If
    
    'NMA Ethics
    If ReadRow(1, 135) Then
        Me.compNMA_Ethics.Visible = True
        Me.compNMA_Ethics.BackColor = &H80FF80
        Me.compNMA_Ethics.Caption = "Date Approved = " & Format(ReadRow(1, 49), "DD-MMM-YYYY")
    Else
        Me.compNMA_Ethics.Visible = False
    End If
    
    'WNHS Ethics
    If ReadRow(1, 136) Then
        Me.compWNHS_Ethics.Visible = True
        Me.compWNHS_Ethics.BackColor = &H80FF80
        Me.compWNHS_Ethics.Caption = "Date Approved = " & Format(ReadRow(1, 52), "DD-MMM-YYYY")
    Else
        Me.compWNHS_Ethics.Visible = False
    End If
    
    'SJOG Ethics
    If ReadRow(1, 137) Then
        Me.compSJOG_Ethics.Visible = True
        Me.compSJOG_Ethics.BackColor = &H80FF80
        Me.compSJOG_Ethics.Caption = "Date Approved = " & Format(ReadRow(1, 55), "DD-MMM-YYYY")
    Else
        Me.compSJOG_Ethics.Visible = False
    End If
    
    'Others Ethics
    If ReadRow(1, 138) Then
        Me.compOthers_Ethics.Visible = True
        Me.compOthers_Ethics.BackColor = &H80FF80
        Me.compOthers_Ethics.Caption = "Date Approved = " & Format(ReadRow(1, 59), "DD-MMM-YYYY")
    Else
        Me.compOthers_Ethics.Visible = False
    End If
    
    'PCH Governance
    If ReadRow(1, 139) Then
        Me.compPCH_Gov.Visible = True
        Me.compPCH_Gov.BackColor = &H80FF80
        Me.compPCH_Gov.Caption = "Date Approved = " & Format(ReadRow(1, 65), "DD-MMM-YYYY")
    Else
        Me.compPCH_Gov.Visible = False
    End If
    
    'TKI Governance
    If ReadRow(1, 140) Then
        Me.compTKI_Gov.Visible = True
        Me.compTKI_Gov.BackColor = &H80FF80
        Me.compTKI_Gov.Caption = "Date Approved = " & Format(ReadRow(1, 69), "DD-MMM-YYYY")
    Else
        Me.compTKI_Gov.Visible = False
    End If
    
    'KEMH Governance
    If ReadRow(1, 141) Then
        Me.compKEMH_Gov.Visible = True
        Me.compKEMH_Gov.BackColor = &H80FF80
        Me.compKEMH_Gov.Caption = "Date Approved = " & Format(ReadRow(1, 73), "DD-MMM-YYYY")
    Else
        Me.compKEMH_Gov.Visible = False
    End If
    
    'SJOG_S Governance
    If ReadRow(1, 142) Then
        Me.compSJOG_S_Gov.Visible = True
        Me.compSJOG_S_Gov.BackColor = &H80FF80
        Me.compSJOG_S_Gov.Caption = "Date Approved = " & Format(ReadRow(1, 77), "DD-MMM-YYYY")
    Else
        Me.compSJOG_S_Gov.Visible = False
    End If
    
    'SJOG_L Governance
    If ReadRow(1, 143) Then
        Me.compSJOG_L_Gov.Visible = True
        Me.compSJOG_L_Gov.BackColor = &H80FF80
        Me.compSJOG_L_Gov.Caption = "Date Approved = " & Format(ReadRow(1, 81), "DD-MMM-YYYY")
    Else
        Me.compSJOG_L_Gov.Visible = False
    End If
    
    'SJOG_M Governance
    If ReadRow(1, 144) Then
        Me.compSJOG_M_Gov.Visible = True
        Me.compSJOG_M_Gov.BackColor = &H80FF80
        Me.compSJOG_M_Gov.Caption = "Date Approved = " & Format(ReadRow(1, 85), "DD-MMM-YYYY")
    Else
        Me.compSJOG_M_Gov.Visible = False
    End If
    
    'Others Governance
    If ReadRow(1, 145) Then
        Me.compOthers_Gov.Visible = True
        Me.compOthers_Gov.BackColor = &H80FF80
        Me.compOthers_Gov.Caption = "Date Approved = " & Format(ReadRow(1, 90), "DD-MMM-YYYY")
    Else
        Me.compOthers_Gov.Visible = False
    End If
    
    'VTG Budget
    If ReadRow(1, 146) Then
        Me.compVTG_Budget.Visible = True
        Me.compVTG_Budget.BackColor = &H80FF80
        Me.compVTG_Budget.Caption = "Date Approved = " & Format(ReadRow(1, 96), "DD-MMM-YYYY")
    Else
        Me.compVTG_Budget.Visible = False
    End If
    
    'TKI Budget
    If ReadRow(1, 147) Then
        Me.compTKI_Budget.Visible = True
        Me.compTKI_Budget.BackColor = &H80FF80
        Me.compTKI_Budget.Caption = "Date Approved = " & Format(ReadRow(1, 98), "DD-MMM-YYYY")
    Else
        Me.compTKI_Budget.Visible = False
    End If
    
    'Pharmacy Budget
    If ReadRow(1, 148) Then
        Me.compPharm_Budget.Visible = True
        Me.compPharm_Budget.BackColor = &H80FF80
        Me.compPharm_Budget.Caption = "PO Finalised = " & Format(ReadRow(1, 101), "DD-MMM-YYYY")
    Else
        Me.compPharm_Budget.Visible = False
    End If
    
    'Indemnity
    If ReadRow(1, 149) Then
        Me.compIndemnity.Visible = True
        Me.compIndemnity.BackColor = &H80FF80
        Me.compIndemnity.Caption = "Date Completed = " & Format(ReadRow(1, 107), "DD-MMM-YYYY")
    Else
        Me.compIndemnity.Visible = False
        Me.compIndemnity.BackColor = &H80000005
        Me.compIndemnity.Caption = vbNullString
    End If
    
    'CTRA
    If ReadRow(1, 150) Then
        Me.compCTRA.Visible = True
        Me.compCTRA.BackColor = &H80FF80
        Me.compCTRA.Caption = "Date Finalised = " & Format(ReadRow(1, 117), "DD-MMM-YYYY")
    Else
        Me.compCTRA.Visible = False
    End If

    'Financial Disclosure
    If ReadRow(1, 151) Then
        Me.compFinDisc.Visible = True
        Me.compFinDisc.BackColor = &H80FF80
        Me.compFinDisc.Caption = "Date Completed = " & Format(ReadRow(1, 121), "DD-MMM-YYYY")
    Else
        Me.compFinDisc.Visible = False
    End If

    'Site Initiation Visit
    If ReadRow(1, 152) Then
        Me.compSIV.Visible = True
        Me.compSIV.BackColor = &H80FF80
        Me.compSIV.Caption = "SIV Date = " & Format(ReadRow(1, 125), "DD-MMM-YYYY")
    Else
        Me.compSIV.Visible = False
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


