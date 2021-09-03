Attribute VB_Name = "Public_Variables"
Option Explicit

'List of public variables to share between userforms
'---------------------------------------------------
Public StudyName As String
Public RowIndex As Long
Public Username As String
Public LastUpdate As Date
Public Tick As Boolean
Public StudyStatus As Variant
Public DisplayArr() As Variant

Public RegTable As ListObject
Public ReadRow As ListRow

'List of fixed value constants to set fixed values
'-------------------------------------------------
Public Const UHeight As Long = 440
Public Const UWidth As Long = 500


'list of public functions to share between userforms
'---------------------------------------------------

Sub LogLastAccess()
    
    'PURPOSE: Log last time entry was accessed
    
    If RowIndex > 0 Then
        With RegTable.ListRows(RowIndex)
            .Range(6) = Now
            .Range(7) = Username
        End With
    End If
End Sub


Public Sub TurnEvents_ON()

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    
End Sub

Public Sub TurnEvents_OFF()

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    
End Sub
