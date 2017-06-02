Option Explicit
Private Const ModuleName As String = "Module1"

Sub DoSomething(ByVal value1 As Integer, ByVal value2 As Integer, ByVal value3 As String)
    CallStack.Push ModuleName, "DoSomething", value1, value2, value3
    TestSomethingElse value1
    CallStack.Pop
End Sub

Private Sub TestSomethingElse(ByVal value1 As Integer)
    CallStack.Push ModuleName, "TestSomethingElse", value1
    On Error GoTo CleanFail

    Debug.Print value1 / 0

CleanExit:
    CallStack.Pop
    Exit Sub
CleanFail:
    PrintErrorInfo
    Resume CleanExit
End Sub

Public Sub PrintErrorInfo()
    Debug.Print "Runtime error " & Err.Number & ": " & Err.Description & vbNewLine & CallStack.ToString
End Sub
