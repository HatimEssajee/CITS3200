Attribute VB_Name = "StressTest"
Sub FillStudies()
Attribute FillStudies.VB_ProcData.VB_Invoke_Func = " \n14"
    
    Dim Comp As Range, Incomp As Range, CopyR As Range
    Dim lastRow As Range
    Dim j As Integer
    
    Set RegTable = ThisWorkbook.Sheets("Register").ListObjects("Register")
    
    Set Comp = RegTable.ListRows(4).Range
    Set Incomp = RegTable.ListRows(7).Range
    
    For j = 1 To 2
        Set lastRow = RegTable.ListRows.Add.Range
        
        If j = 1 Then
            Set CopyR = Comp
        Else
            Set CopyR = Incomp
        End If
        
        CopyR.Copy lastRow
    Next j
    
End Sub
