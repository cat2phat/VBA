VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CallStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private frames As New Collection

Public Sub Push(ByVal module As String, ByVal member As String, ParamArray parameterValues() As Variant)
Attribute Push.VB_Description = "Pushes a new stack frame onto the call stack. Call once at the entry point of each procedure to trace."
    Dim values() As Variant
    values = parameterValues
    frames.Add StackFrame.Create(module, member, values)
End Sub

Public Function Pop() As IStackFrame
Attribute Pop.VB_Description = "Removes the last stack frame from the top of the stack. Call once at the exit point of each traced procedure."
    Set Pop = Peek
    frames.Remove frames.Count
End Function

Public Function Peek() As IStackFrame
Attribute Peek.VB_Description = "Returns the top-most stack frame."
    Set Peek = frames(frames.Count)
End Function

Public Property Get Count() As Long
Attribute Count.VB_Description = "Gets the depth of the call stack."
    Count = frames.Count
End Property

Public Function ToString() As String
Attribute ToString.VB_Description = "Returns a String containing the stack trace."
    Dim result As String
    Dim index As Long
    For index = frames.Count To 1 Step -1
        result = result & "at " & frames(index).ToString & IIf(index = 1, vbNullString, vbNewLine)
    Next
    ToString = result
End Function
