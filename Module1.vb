Function DAS() As String
    Dim objShell As Object
    Set objShell = CreateObject("WScript.Network")
'automatically retrieves DAS from Windows Login
    DAS = Application.WorksheetFunction.Proper(objShell.UserName)
End Function


Function EmployeeName() As String
'automatically retrieves username with proper caps from Windows login
    EmployeeName = Application.WorksheetFunction.Proper(Application.UserName)
End Function
