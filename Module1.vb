Function DAS() As String
    Dim objShell As Object
    Set objShell = CreateObject("WScript.Network")
    DAS = Application.WorksheetFunction.Proper(objShell.UserName)
End Function


Function EmployeeName() As String
    EmployeeName = Application.WorksheetFunction.Proper(Application.UserName)
End Function
