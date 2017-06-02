VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IStackFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get ModuleName() As String
Attribute ModuleName.VB_Description = "Gets the name of the module for this instance."
End Property

Public Property Get MemberName() As String
Attribute ModuleName.VB_Description = "Gets the name of the member for this instance."
End Property

Public Property Get ParameterValue(ByVal index As Integer) As Variant
Attribute ModuleName.VB_Description = "Gets the value of the parameter at the specified index."
End Property

Public Function ToString() As String
Attribute ToString.VB_Description = "Returns a string representation of the member and its arguments."
End Function
