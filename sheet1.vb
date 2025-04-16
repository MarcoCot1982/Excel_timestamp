Private Sub Worksheet_Change(ByVal Target As Range)
    Dim cell As Range
    Dim ws As Worksheet
    
    'Define the working sheet as the active sheet
    Set ws = Me  '"Me" refers to the sheet where the change occurred
    
    'Check if the change is within column I (column 9)
    If Not Intersect(Target, ws.Range("I2:I99")) Is Nothing Then
        Application.EnableEvents = False 'Prevent event looping
        
        'Unprotect the sheet
        ws.Unprotect Password:="NeverEdit"
        
        For Each cell In Target
            If IsNumeric(cell.Value) And cell.Value <> "" Then
                'Force timestamp format with slashes
                cell.Offset(0, 2).Value = Replace(Format(Now, "dd/mm/yyyy HH:mm:ss"), "-", "/") & " - " & EmployeeName()
            ElseIf cell.Value = "" Then
                'Clear column K if column I is emptied
                cell.Offset(0, 2).ClearContents
            End If
        Next cell
        
        'Protect the sheet again
        ws.Protect Password:="NeverEdit"
        
        Application.EnableEvents = True 'Re-enable events
    End If
End Sub
