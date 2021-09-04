Attribute VB_Name = "Public_Variables"
Option Explicit

'List of public variables to share between userforms
'---------------------------------------------------
Public RowIndex As Long
Public Username As String
Public LastUpdate As Date
Public Tick As Boolean
Public StudyStatus As Variant
Public DisplayArr() As Variant

Public RegTable As ListObject
Public ReadRow As ListRow
Public UserFormLeftPos As Long
Public UserFormTopPos As Long


'List of fixed value constants to set fixed values
'-------------------------------------------------
Public Const UHeight As Long = 470
Public Const UWidth As Long = 500


'list of public functions to share between userforms
'---------------------------------------------------

Public Sub LogLastAccess()
    
    'PURPOSE: Log last time entry was accessed
    
    If RowIndex > 0 Then
        With RegTable.ListRows(RowIndex)
            .Range(6) = Now
            .Range(7) = Username
        End With
    End If
End Sub

Public Function String_to_Date(Txt As String)
    'PURPOSE: Convert string input to date value if it is a valid date
    
    If IsDate(Txt) Then
        String_to_Date = DateValue(Txt)
    Else
        String_to_Date = Txt
    End If
    
End Function

Public Function Date_Validation(CurrDate As String, Optional PrevDate As String = "", Optional err2 As String = "") As String
    'PURPOSE: Assess data input is in correct format and output error message string
    
    Dim err As String
    Dim d1 As Variant
    Dim d2 As Variant
    
    err = vbNullString
    
    If CurrDate <> vbNullString And Not IsDate(CurrDate) Then
        err = "Please enter a valid date:" & Chr(10) & "DD-MMM-YYYY"
    End If
    
    d1 = String_to_Date(PrevDate)
    d2 = String_to_Date(CurrDate)
    
    'If no date entry issue, check date for chronology
    If err = "" And d1 <> "" And d2 <> "" And d2 < d1 Then
        err = err2
    End If
    
    Date_Validation = err
    
End Function

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
