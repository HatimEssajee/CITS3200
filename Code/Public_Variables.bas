Attribute VB_Name = "Public_Variables"
Option Explicit

'List of public variables to share between userforms
'---------------------------------------------------

'Row of table being read
Public RowIndex As Long
Public RegTable As ListObject

'Version control
Public Username As String
Public LastUpdate As Date

'Conditional Navigation
Public Tick As Boolean
Public FC_Tick As Boolean
Public SAG_Tick As Boolean
Public StudyStatus As Variant
Public Correct As Variant
Public OldValues As Variant
Public NxtOldValues As Variant

'Search
Public DisplayArr() As Variant

'Undo delete
Public OldStudyStatus As String

'Userform Position storage
Public UserFormLeftPos As Double
Public UserFormTopPos As Double
Public UserFormLeftPosC As Double
Public UserFormTopPosC As Double
Public UserFormLeftPosR As Double
Public UserFormTopPosR As Double
Public UHeight As Double
Public UWidth As Double

'Userform dimension control
Public Const DUHeight As Long = 470 '610
Public Const DUWidth As Long = 650 '500


'list of public functions to share between userforms
'---------------------------------------------------

Public Sub LogLastAccess()
    
    'PURPOSE: Log last time entry was accessed
    
    If RowIndex > 0 Then
        With RegTable.ListRows(RowIndex)
            .Range(5) = Now
            .Range(6) = Username
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
    If err = "" And d1 <> "" And d2 <> "" And _
        IsDate(d1) And IsDate(d2) And d2 < d1 Then
        err = err2
    End If
    
    Date_Validation = err
    
End Function

Public Function ArraysSame(ArrX As Variant, ArrY As Variant) As Boolean
    'PURPOSE: Compare values of two 1D arrays
    
    Dim Check As Boolean
    Dim Upper As Long, i As Long
    
    Check = True
    Upper = UBound(ArrX)
    
    'Shift upper bound to smaller array
    If UBound(ArrX) >= UBound(ArrY) Then
        Upper = UBound(ArrY)
    End If
    
    For i = LBound(ArrX) To Upper
        If ArrX(i) <> ArrY(i) Then
            Check = False
            Exit For
        End If
    Next i
    
    ArraysSame = Check
End Function

Public Function ReadDate(dstr As String) As String
    
    'PURPOSE: Check date fields are valid dates and not numbers when read from excel table
    If IsDate(dstr) Then
       dstr = Format(dstr, "dd-mmm-yyyy")
    End If
    
    ReadDate = dstr
    
End Function

Public Function WriteText(Field As Variant) As String
    
    Dim str As String
    'PURPOSE: Check field is a number and add an apostrophe infront to store as text when written into userform
    If IsNumeric(Field) Then
        str = "'" & CStr(Field)
    Else
        str = CStr(Field)
    End If
    
    WriteText = str
    
End Function
